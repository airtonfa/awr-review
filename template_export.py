#!/usr/bin/env python3
"""
Template-native AWR PPT exporter.

Uses an Oracle .potx template and only Light/white layouts.
Input can be either:
1) Raw AWR JSON (instances/cohort_rollups/database_statistics/...); or
2) Saved app report JSON containing { raw_data, ui_state }.
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import math
import os
import re
import struct
import tempfile
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor


ORACLE_COLORS = {
    "text": RGBColor(0x2A, 0x2F, 0x2F),
    "muted": RGBColor(0x6B, 0x74, 0x7A),
    "accent1": RGBColor(0x04, 0x53, 0x6F),
    "accent2": RGBColor(0x6C, 0x3F, 0x49),
    "accent3": RGBColor(0xC7, 0x46, 0x34),
    "accent4": RGBColor(0xF0, 0xCC, 0x71),
    "accent5": RGBColor(0x89, 0xB2, 0xB0),
    "accent6": RGBColor(0x86, 0xB5, 0x96),
    "panel": RGBColor(0xF8, 0xFA, 0xFC),
    "line": RGBColor(0xD9, 0xDE, 0xDE),
}
PALETTE = [
    ORACLE_COLORS["accent3"],
    ORACLE_COLORS["accent1"],
    ORACLE_COLORS["accent2"],
    ORACLE_COLORS["accent5"],
    ORACLE_COLORS["accent6"],
    ORACLE_COLORS["accent4"],
]


def short_label(s: Any, n: int = 34) -> str:
    t = str(s or "")
    return t if len(t) <= n else f"{t[: n - 1]}…"


def to_num(v: Any, default: float = 0.0) -> float:
    try:
        return float(v)
    except Exception:
        return default


def fmt_num(v: Any, dec: int = 1) -> str:
    try:
        x = float(v)
    except Exception:
        return "N/A"
    if dec == 0:
        return f"{int(round(x)):,}"
    return f"{x:,.{dec}f}"


def decode_f64_b64(s: str) -> list[float]:
    raw = base64.b64decode(s + "=" * ((4 - (len(s) % 4)) % 4))
    out = []
    for i in range(0, len(raw), 8):
        out.append(struct.unpack("<d", raw[i : i + 8])[0])
    return out


def _plotly_payload_tail(file_content: str) -> str:
    txt = file_content or ""
    i = txt.rfind("Plotly.newPlot(")
    return txt[i:] if i >= 0 else txt


def _extract_plotly_traces(file_content: str) -> list[dict[str, Any]]:
    src = _plotly_payload_tail(file_content)
    if not src:
        return []
    i0 = src.find("(")
    if i0 < 0:
        return []
    c1 = src.find(",", i0 + 1)  # after div id
    if c1 < 0:
        return []
    a0 = src.find("[", c1 + 1)  # data traces array
    if a0 < 0:
        return []

    level = 0
    in_str = False
    esc = False
    a1 = -1
    for idx, ch in enumerate(src[a0:], a0):
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "[":
            level += 1
        elif ch == "]":
            level -= 1
            if level == 0:
                a1 = idx
                break
    if a1 <= a0:
        return []
    try:
        parsed = json.loads(src[a0 : a1 + 1])
    except Exception:
        return []
    return parsed if isinstance(parsed, list) else []


def parse_time_axis_from_plot_html(file_content: str) -> list[str]:
    traces = _extract_plotly_traces(file_content)
    if not traces:
        return []
    x = traces[0].get("x")
    if isinstance(x, list):
        return [str(v) for v in x]
    return []


def parse_numeric_axis_from_plot_html(file_content: str) -> list[float]:
    traces = _extract_plotly_traces(file_content)
    if not traces:
        return []
    y = traces[0].get("y")
    if isinstance(y, dict):
        dtype = str(y.get("dtype") or "").lower()
        bdata = y.get("bdata")
        if isinstance(bdata, str) and dtype == "f8":
            try:
                return decode_f64_b64(bdata)
            except Exception:
                return []
    if isinstance(y, list):
        try:
            return [float(v) for v in y]
        except Exception:
            return []
    return []


def parse_timeseries_from_plot_html(file_content: str) -> tuple[list[str], list[float]] | None:
    xs = parse_time_axis_from_plot_html(file_content)
    ys = parse_numeric_axis_from_plot_html(file_content)
    if not xs or not ys:
        return None
    n = min(len(xs), len(ys))
    return xs[:n], ys[:n]


def quantile(sorted_vals: list[float], q: float) -> float:
    if not sorted_vals:
        return 0.0
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    pos = (len(sorted_vals) - 1) * q
    base = int(math.floor(pos))
    frac = pos - base
    if base + 1 < len(sorted_vals):
        return sorted_vals[base] + frac * (sorted_vals[base + 1] - sorted_vals[base])
    return sorted_vals[base]


def stats_from_values(vals: list[float]) -> dict[str, float] | None:
    clean = sorted([float(v) for v in vals if isinstance(v, (int, float)) and math.isfinite(float(v))])
    if not clean:
        return None
    return {
        "min": clean[0],
        "p30": quantile(clean, 0.30),
        "p50": quantile(clean, 0.50),
        "p70": quantile(clean, 0.70),
        "p95": quantile(clean, 0.95),
        "max": clean[-1],
    }


@dataclass
class Row:
    cohort: str
    db: str
    instance: str
    host: str
    init_cpu_count: float
    logical_cpu_count: float
    mem_gb: float
    sga_gb: float
    pga_gb: float


class Model:
    def __init__(self, raw: dict[str, Any], ui_state: dict[str, Any] | None = None):
        self.raw = raw
        self.ui_state = ui_state or {}
        self.rows: list[Row] = self._normalize_rows()
        self.cpu_by_cohort = self._parse_cpu_series()
        self.snap_time_by_cohort = self._parse_snap_time_axis()
        self.db_stats = self._normalize_db_stats()
        self.db_versions = self._normalize_db_versions()
        self.db_names = self._normalize_db_names()
        self.db_meta = self._normalize_db_meta()
        self.db_params = self._normalize_db_params()
        self.cohort_rollups = self._normalize_cohort_rollups()
        self.selection = self._resolve_selection()

    def _normalize_rows(self) -> list[Row]:
        out: list[Row] = []
        for r in self.raw.get("instances", []) or []:
            cohort = r.get("cdb_cohort") or r.get("db_cohort") or "UNASSIGNED"
            db = r.get("cdb") or r.get("WHICH_DB") or "UNKNOWN_DB"
            inst = r.get("db") or f"{db}${r.get('host') or 'UNKNOWN_HOST'}"
            out.append(
                Row(
                    cohort=cohort,
                    db=db,
                    instance=inst,
                    host=r.get("host") or "UNKNOWN_HOST",
                    init_cpu_count=to_num(r.get("init_cpu_count")),
                    logical_cpu_count=to_num(r.get("logical_cpu_count")),
                    mem_gb=to_num(r.get("mem_gb")),
                    sga_gb=to_num(r.get("sga_size_gb")),
                    pga_gb=to_num(r.get("pga_size_gb")),
                )
            )
        return out

    def _parse_cpu_series(self) -> dict[str, tuple[list[str], list[float]]]:
        out: dict[str, tuple[list[str], list[float]]] = {}
        for p in self.raw.get("plots_html", []) or []:
            name = p.get("file_name") or ""
            if not name.endswith("_DB_vCPU_unadjusted.html"):
                continue
            cohort = name.replace("_DB_vCPU_unadjusted.html", "")
            parsed = parse_timeseries_from_plot_html(p.get("file_content") or "")
            if parsed:
                out[cohort] = parsed
                c2 = p.get("cohort")
                if c2 and c2 not in out:
                    out[c2] = parsed
        return out

    def _parse_snap_time_axis(self) -> dict[str, list[str]]:
        out: dict[str, list[str]] = {}
        for p in self.raw.get("plots_html", []) or []:
            name = p.get("file_name") or ""
            if not name.endswith("_awr_snap_coverage.html"):
                continue
            cohort = name.replace("_awr_snap_coverage.html", "")
            xs = parse_time_axis_from_plot_html(p.get("file_content") or "")
            if xs:
                out[cohort] = xs
                c2 = p.get("cohort")
                if c2 and c2 not in out:
                    out[c2] = xs
        return out

    def _normalize_db_stats(self) -> list[dict[str, Any]]:
        out = []
        for r in self.raw.get("database_statistics", []) or []:
            out.append(
                {
                    "cohort": r.get("cdb_cohort") or "UNASSIGNED",
                    "db": r.get("cdb") or "UNKNOWN_DB",
                    "metric": r.get("metric") or "",
                    "min": to_num(r.get("min")),
                    "p30": to_num(r.get("p30")),
                    "p50": to_num(r.get("p50")),
                    "p70": to_num(r.get("p70")),
                    "p95": to_num(r.get("p95")),
                    "p99": to_num(r.get("p99")),
                    "max": to_num(r.get("max")),
                    "summary_of": r.get("summary_of") or "",
                }
            )
        return out

    def _normalize_db_versions(self) -> dict[str, str]:
        out: dict[str, str] = {}
        for d in self.raw.get("databases", []) or []:
            db = d.get("cdb") or d.get("WHICH_DB")
            if not db:
                continue
            out[db] = d.get("cdb_version") or d.get("db_version") or "Unknown"
        return out

    def _normalize_db_names(self) -> dict[str, str]:
        out: dict[str, str] = {}
        for r in self.raw.get("properties_database", []) or []:
            db = r.get("WHICH_DB") or r.get("cdb")
            if not db:
                continue
            nm = r.get("DB_NAME")
            if nm:
                out[db] = str(nm)
        return out

    def _normalize_db_meta(self) -> dict[str, dict[str, Any]]:
        out: dict[str, dict[str, Any]] = {}
        for d in self.raw.get("databases", []) or []:
            db = d.get("cdb") or d.get("WHICH_DB")
            if not db:
                continue
            out[db] = {
                "version": d.get("cdb_version") or d.get("db_version") or "Unknown",
                "instance_count": int(to_num(d.get("cdb_instance_count"), 0)),
                "clustered": d.get("clustered") or "",
            }
        return out

    def _normalize_db_params(self) -> dict[str, dict[str, str]]:
        out: dict[str, dict[str, str]] = defaultdict(dict)
        for r in self.raw.get("database_parameters", []) or []:
            db = r.get("cdb")
            pn = str(r.get("PARAMETER_NAME") or "").strip().lower()
            if not db or not pn:
                continue
            if pn in {"enable_pluggable_database", "target_pdbs", "max_pdbs", "noncdb_compatible"}:
                out[db][pn] = str(r.get("VALUE") or "").strip()
        return dict(out)

    def _normalize_cohort_rollups(self) -> dict[str, dict[str, float]]:
        out: dict[str, dict[str, float]] = {}
        for r in self.raw.get("cohort_rollups", []) or []:
            c = r.get("Name")
            if not c:
                continue
            out[c] = {
                "allocated": to_num(r.get("Allocated Storage (GB)")),
                "used": to_num(r.get("Used Storage (GB)")),
                "db_iops": to_num(r.get("DB IOPS")),
            }
        return out

    def _resolve_selection(self) -> dict[str, set[str]]:
        rows = self.rows
        all_cohorts = {r.cohort for r in rows}
        all_dbs = {r.db for r in rows}
        all_inst = {r.instance for r in rows}
        ui = self.ui_state.get("selection", {}).get("ppt", {}) if isinstance(self.ui_state, dict) else {}
        cohorts = set(ui.get("cohorts", []) or all_cohorts)
        dbs = set(ui.get("dbs", []) or all_dbs)
        inst = set(ui.get("instances", []) or all_inst)
        return {"cohorts": cohorts, "dbs": dbs, "instances": inst}

    def selected_rows(self) -> list[Row]:
        sel = self.selection
        return [r for r in self.rows if r.cohort in sel["cohorts"] and r.db in sel["dbs"] and r.instance in sel["instances"]]

    def cohorts(self) -> list[str]:
        return sorted({r.cohort for r in self.selected_rows()})

    def rows_for_cohort(self, cohort: str) -> list[Row]:
        return [r for r in self.selected_rows() if r.cohort == cohort]

    def display_db_name(self, db: str) -> str:
        nm = self.db_names.get(db)
        if nm:
            return nm
        token = str(db or "UNKNOWN_DB")
        if token.startswith("db_"):
            token = token[3:]
        token = token.split("_")[0] if "_" in token else token
        return token or "UNKNOWN_DB"

    def instance_label(self, inst: str) -> str:
        parts = str(inst or "").split("$")
        if len(parts) >= 3:
            return f"{parts[-2]}${parts[-1]}"
        return str(inst or "UNKNOWN_INSTANCE")

    def db_type_and_pdb_count(self, db: str) -> tuple[str, str]:
        params = self.db_params.get(db, {})
        epd = params.get("enable_pluggable_database", "").upper()
        target_pdbs = params.get("target_pdbs", "")
        db_type = "Non-CDB"
        if epd == "TRUE":
            db_type = "CDB"
        elif epd == "FALSE":
            db_type = "Non-CDB"

        if db_type == "CDB":
            try:
                return db_type, str(max(0, int(float(target_pdbs))))
            except Exception:
                return db_type, "N/A"
        return db_type, "0"

    def db_version(self, db: str) -> str:
        meta = self.db_meta.get(db, {})
        v = meta.get("version") or self.db_versions.get(db) or "Unknown"
        return str(v)

    def db_is_rac(self, db: str) -> bool:
        meta = self.db_meta.get(db, {})
        clustered = str(meta.get("clustered") or "").upper()
        return "CLUSTERED" in clustered and "NON" not in clustered

    def summary(self, rows: list[Row] | None = None) -> dict[str, float]:
        rs = rows if rows is not None else self.selected_rows()
        cohorts = {r.cohort for r in rs}
        allocated = sum((self.cohort_rollups.get(c, {}).get("allocated") or 0.0) for c in cohorts)
        used = sum((self.cohort_rollups.get(c, {}).get("used") or 0.0) for c in cohorts)
        return {
            "db_count": len({r.db for r in rs}),
            "host_count": len({r.host for r in rs}),
            "instance_count": len(rs),
            "vcpu_total": sum((r.init_cpu_count or r.logical_cpu_count) for r in rs),
            "memory_gb_total": sum(r.mem_gb for r in rs),
            "sga_total": sum(r.sga_gb for r in rs),
            "pga_total": sum(r.pga_gb for r in rs),
            "allocated_storage_gb": allocated,
            "used_storage_gb": used,
        }

    def by_cohort_table(self) -> list[dict[str, Any]]:
        groups: dict[str, dict[str, Any]] = {}
        for r in self.selected_rows():
            g = groups.setdefault(
                r.cohort,
                {"cohort": r.cohort, "dbs": set(), "hosts": set(), "instances": 0, "vcpu": 0.0, "mem": 0.0},
            )
            g["dbs"].add(r.db)
            g["hosts"].add(r.host)
            g["instances"] += 1
            g["vcpu"] += r.init_cpu_count or r.logical_cpu_count
            g["mem"] += r.mem_gb
        out = []
        for g in groups.values():
            out.append(
                {
                    "cohort": g["cohort"],
                    "dbs": len(g["dbs"]),
                    "hosts": len(g["hosts"]),
                    "instances": g["instances"],
                    "vcpu": g["vcpu"],
                    "mem": g["mem"],
                }
            )
        return sorted(out, key=lambda x: x["cohort"])

    def version_counts(self) -> list[tuple[str, int]]:
        cnt: dict[str, int] = defaultdict(int)
        selected_db = sorted({r.db for r in self.selected_rows()})
        for db in selected_db:
            cnt[self.db_versions.get(db, "Unknown")] += 1
        return sorted(cnt.items(), key=lambda kv: (-kv[1], kv[0]))

    def cpu_series_global(self) -> list[tuple[str, list[str], list[float]]]:
        out = []
        for c in self.cohorts():
            s = self.cpu_by_cohort.get(c)
            if s:
                out.append((c, s[0], s[1]))
        return out

    def cpu_stats_by_cohort(self) -> list[tuple[str, dict[str, float]]]:
        out: list[tuple[str, dict[str, float]]] = []
        for c, _xs, ys in self.cpu_series_global():
            st = stats_from_values(ys)
            if st:
                out.append((c, st))
        return out

    def memory_by_cohort(self) -> list[tuple[str, float]]:
        rows = self.by_cohort_table()
        return sorted([(r["cohort"], float(r["mem"])) for r in rows if float(r["mem"]) > 0], key=lambda kv: kv[0])

    def iops_by_cohort(self) -> list[tuple[str, float]]:
        out: list[tuple[str, float]] = []
        for c in self.cohorts():
            v = float((self.cohort_rollups.get(c, {}) or {}).get("db_iops") or 0.0)
            if v > 0:
                out.append((c, v))
        return sorted(out, key=lambda kv: kv[0])

    def cpu_series_instances_in_cohort(self, cohort: str) -> list[tuple[str, list[str], list[float]]]:
        base = self.cpu_by_cohort.get(cohort)
        if not base:
            return []
        xs, ys = base
        rows = self.rows_for_cohort(cohort)
        if not rows:
            return []
        total = sum(r.init_cpu_count for r in rows)
        share_default = 1.0 / len(rows)
        out = []
        for r in rows:
            share = (r.init_cpu_count / total) if total > 0 else share_default
            out.append((r.instance, xs, [v * share for v in ys]))
        return out

    def mem_by_db(self, cohort: str, db_scope: set[str] | None = None) -> list[tuple[str, float]]:
        rows = [
            x
            for x in self.db_stats
            if x["cohort"] == cohort
            and x["metric"] == "DB Memory (MB)"
            and (not x["summary_of"] or x["summary_of"] == "cdb")
            and (db_scope is None or x["db"] in db_scope)
        ]
        groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for r in rows:
            groups[r["db"]].append(r)
        out = []
        for db, arr in groups.items():
            p95 = sum(x["p95"] for x in arr) / max(len(arr), 1) / 1024.0
            mx = sum(x["max"] for x in arr) / max(len(arr), 1) / 1024.0
            out.append((db, p95 if p95 > 0 else mx))
        return sorted(out, key=lambda kv: (-kv[1], kv[0]))

    def metric_by_db(self, cohort: str, metric: str, db_scope: set[str] | None = None, scale: float = 1.0) -> list[tuple[str, float]]:
        rows = [
            x
            for x in self.db_stats
            if x["cohort"] == cohort
            and x["metric"] == metric
            and (not x["summary_of"] or x["summary_of"] == "cdb")
            and (db_scope is None or x["db"] in db_scope)
        ]
        groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for r in rows:
            groups[r["db"]].append(r)
        out = []
        for db, arr in groups.items():
            p95 = (sum(x["p95"] for x in arr) / max(len(arr), 1)) * scale
            mx = (sum(x["max"] for x in arr) / max(len(arr), 1)) * scale
            out.append((db, p95 if p95 > 0 else mx))
        return sorted(out, key=lambda kv: kv[0])

    def db_metric_stats(self, cohort: str, db: str, metric: str) -> dict[str, float] | None:
        rows = [
            x
            for x in self.db_stats
            if x["cohort"] == cohort
            and x["db"] == db
            and x["metric"] == metric
            and (not x["summary_of"] or x["summary_of"] == "cdb")
        ]
        if not rows:
            rows = [x for x in self.db_stats if x["cohort"] == cohort and x["db"] == db and x["metric"] == metric]
        if not rows:
            return None
        keys = ["min", "p30", "p50", "p70", "p95", "p99", "max"]
        out: dict[str, float] = {}
        for k in keys:
            out[k] = sum(to_num(r.get(k)) for r in rows) / max(len(rows), 1)
        return out

    def _cohort_time_range(self, cohort: str) -> tuple[datetime, datetime] | None:
        keys = list(self.cpu_by_cohort.keys())
        snap_keys = list(self.snap_time_by_cohort.keys())

        def resolve_cpu_for_cohort(name: str):
            if name in self.cpu_by_cohort:
                return self.cpu_by_cohort[name]
            up = name.upper()
            for k in keys:
                if str(k).upper() == up:
                    return self.cpu_by_cohort[k]
            for k in keys:
                ks = str(k).upper()
                if ks.startswith(up) or up.startswith(ks):
                    return self.cpu_by_cohort[k]
            return None

        def resolve_snap_for_cohort(name: str):
            if name in self.snap_time_by_cohort:
                return self.snap_time_by_cohort[name]
            up = name.upper()
            for k in snap_keys:
                if str(k).upper() == up:
                    return self.snap_time_by_cohort[k]
            for k in snap_keys:
                ks = str(k).upper()
                if ks.startswith(up) or up.startswith(ks):
                    return self.snap_time_by_cohort[k]
            return None

        def parse_dt(xs: list[str]) -> list[datetime]:
            ts: list[datetime] = []
            for x in xs:
                txt = str(x).replace("Z", "")
                try:
                    ts.append(datetime.fromisoformat(txt))
                except Exception:
                    continue
            return ts

        s = resolve_cpu_for_cohort(cohort)
        ts = parse_dt(s[0]) if s else []
        if not ts:
            xs = resolve_snap_for_cohort(cohort) or []
            ts = parse_dt(xs)
        if not ts:
            return None
        return (min(ts), max(ts))

    def freshness_ranges_by_instance_in_cohort(self, cohort: str) -> list[tuple[str, datetime, datetime]]:
        span = self._cohort_time_range(cohort)
        if not span:
            return []
        start_dt, end_dt = span
        instances = sorted({r.instance for r in self.rows_for_cohort(cohort)})
        return [(inst, start_dt, end_dt) for inst in instances]

    def freshness_ranges(self) -> list[tuple[str, datetime, datetime]]:
        out: list[tuple[str, datetime, datetime]] = []

        for c in self.cohorts():
            span = self._cohort_time_range(c)
            if not span:
                continue
            out.append((c, span[0], span[1]))
        # Fallback: if cohort matching fails, still provide global timeline by available keys.
        if not out:
            for k, s in list(self.cpu_by_cohort.items())[:8]:
                ts = parse_dt(s[0])
                if ts:
                    out.append((str(k), min(ts), max(ts)))
            if not out:
                for k, xs in list(self.snap_time_by_cohort.items())[:8]:
                    ts = parse_dt(xs)
                    if ts:
                        out.append((str(k), min(ts), max(ts)))
        out.sort(key=lambda t: t[0])
        return out


def find_layout(prs: Presentation, preferred_names: list[str], fallback_light: bool = True):
    names = [n.lower() for n in preferred_names]
    for lay in prs.slide_layouts:
        n = (lay.name or "").lower()
        if any(p in n for p in names):
            return lay
    if fallback_light:
        for lay in prs.slide_layouts:
            n = (lay.name or "").lower()
            if "light" in n and "dark" not in n:
                return lay
    return prs.slide_layouts[0]


def presentation_from_template(path: Path) -> Presentation:
    # python-pptx rejects .potx content-type directly; create a temporary
    # .pptx clone with presentation main content-type.
    if path.suffix.lower() != ".potx":
        return Presentation(str(path))
    with zipfile.ZipFile(path, "r") as zin:
        tmpdir = Path(tempfile.mkdtemp(prefix="potx_as_pptx_"))
        zin.extractall(tmpdir)
    ct = tmpdir / "[Content_Types].xml"
    txt = ct.read_text(encoding="utf-8")
    txt = txt.replace(
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    )
    ct.write_text(txt, encoding="utf-8")
    out = tmpdir / "template_as_presentation.pptx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for p in sorted(tmpdir.rglob("*")):
            if p == out or p.is_dir():
                continue
            zout.write(p, p.relative_to(tmpdir))
    return Presentation(str(out))


def clear_existing_slides(prs: Presentation) -> None:
    """
    Remove any slides already present in the template so export contains only
    generated AWR slides.
    """
    slide_ids = list(prs.slides._sldIdLst)  # pylint: disable=protected-access
    for sld_id in slide_ids:
        r_id = sld_id.rId
        prs.part.drop_rel(r_id)
        prs.slides._sldIdLst.remove(sld_id)  # pylint: disable=protected-access


def set_title_and_subtitle(slide, title: str, subtitle: str = ""):
    if slide.shapes.title is not None:
        slide.shapes.title.text = title
        tf = slide.shapes.title.text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                r.font.name = "Oracle Sans Tab"
                r.font.color.rgb = ORACLE_COLORS["text"]
    # First body placeholder as subtitle when available
    body_ph = None
    for ph in slide.placeholders:
        if ph.placeholder_format.type == 2:  # BODY
            body_ph = ph
            break
    if body_ph is not None and subtitle:
        body_ph.text = subtitle
        tf = body_ph.text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                r.font.name = "Oracle Sans Tab"
                r.font.color.rgb = ORACLE_COLORS["muted"]


def add_kpi_boxes(slide, kpis: list[tuple[str, str]]):
    # Content region in white template is roughly 0.84..12.5 x 1.75..6.68
    cols = 4
    x0 = 0.9
    y0 = 1.75
    w = 2.85
    h = 0.72
    gx = 0.18
    gy = 0.1
    for i, (k, v) in enumerate(kpis[:8]):
        c = i % cols
        r = i // cols
        x = x0 + c * (w + gx)
        y = y0 + r * (h + gy)
        shp = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))  # MSO_AUTO_SHAPE_TYPE.RECTANGLE
        shp.fill.solid()
        shp.fill.fore_color.rgb = ORACLE_COLORS["panel"]
        shp.line.color.rgb = ORACLE_COLORS["line"]
        tf = shp.text_frame
        tf.clear()
        p1 = tf.paragraphs[0]
        p1.text = k
        p1.font.name = "Oracle Sans Tab"
        p1.font.size = Pt(9)
        p1.font.color.rgb = ORACLE_COLORS["muted"]
        p2 = tf.add_paragraph()
        p2.text = v
        p2.font.name = "Oracle Sans Tab"
        p2.font.bold = True
        p2.font.size = Pt(13)
        p2.font.color.rgb = ORACLE_COLORS["text"]


def add_exec_cards(slide, cards: list[dict[str, str]]):
    x0 = 0.8
    y0 = 1.55
    count = min(len(cards), 5)
    gap = 0.14
    usable_w = 12.0
    w = (usable_w - (gap * (count - 1))) / max(count, 1)
    h = 1.35
    for i, c in enumerate(cards[:5]):
        x = x0 + i * (w + gap)
        box = slide.shapes.add_shape(1, Inches(x), Inches(y0), Inches(w), Inches(h))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0xF7, 0xF9, 0xFC)
        box.line.color.rgb = ORACLE_COLORS["line"]
        # Top accent strip
        strip = slide.shapes.add_shape(1, Inches(x), Inches(y0), Inches(w), Inches(0.12))
        strip.fill.solid()
        strip.fill.fore_color.rgb = c.get("color", ORACLE_COLORS["accent1"])
        strip.line.fill.background()
        tf = box.text_frame
        tf.clear()
        p0 = tf.paragraphs[0]
        p0.text = c.get("label", "")
        p0.font.name = "Oracle Sans Tab"
        p0.font.bold = True
        p0.font.size = Pt(10)
        p0.font.color.rgb = ORACLE_COLORS["muted"]
        p0.alignment = PP_ALIGN.LEFT
        p1 = tf.add_paragraph()
        p1.text = c.get("value", "")
        p1.font.name = "Oracle Sans Tab"
        p1.font.bold = True
        p1.font.size = Pt(22)
        p1.font.color.rgb = c.get("color", ORACLE_COLORS["accent1"])
        p1.alignment = PP_ALIGN.LEFT
        p2 = tf.add_paragraph()
        p2.text = c.get("sub", "")
        p2.font.name = "Oracle Sans Tab"
        p2.font.size = Pt(9)
        p2.font.color.rgb = ORACLE_COLORS["muted"]
        p2.alignment = PP_ALIGN.LEFT


def add_big_versions_card(slide, title: str, lines: list[str], x: float, y: float, w: float, h: float):
    box = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF7, 0xF9, 0xFC)
    box.line.color.rgb = ORACLE_COLORS["line"]
    strip = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(0.12))
    strip.fill.solid()
    strip.fill.fore_color.rgb = ORACLE_COLORS["accent2"]
    strip.line.fill.background()
    tf = box.text_frame
    tf.clear()
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Oracle Sans Tab"
    p0.font.bold = True
    p0.font.size = Pt(18)
    p0.font.color.rgb = ORACLE_COLORS["text"]
    p0.alignment = PP_ALIGN.LEFT
    for line in lines[:5]:
        p = tf.add_paragraph()
        p.text = line
        p.font.name = "Oracle Sans Tab"
        p.font.size = Pt(14)
        p.font.color.rgb = ORACLE_COLORS["muted"]
        p.alignment = PP_ALIGN.LEFT


def add_panel(
    slide,
    title: str,
    lines: list[str],
    x: float,
    y: float,
    w: float,
    h: float,
    dark: bool = False,
    body_font_size: int = 12,
):
    box = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x23, 0x2B, 0x62) if dark else RGBColor(0xF7, 0xF9, 0xFC)
    box.line.color.rgb = ORACLE_COLORS["line"]
    if not dark:
        left = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(0.12), Inches(h))
        left.fill.solid()
        left.fill.fore_color.rgb = ORACLE_COLORS["accent1"]
        left.line.fill.background()
    tf = box.text_frame
    tf.margin_left = Cm(0.5)
    tf.clear()
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Oracle Sans Tab"
    p0.font.bold = True
    p0.font.size = Pt(18)
    p0.font.color.rgb = RGBColor(0xE7, 0xEE, 0xFF) if dark else ORACLE_COLORS["accent1"]
    p0.alignment = PP_ALIGN.LEFT
    for ln in lines[:8]:
        p = tf.add_paragraph()
        p.text = ln
        p.font.name = "Oracle Sans Tab"
        p.font.size = Pt(body_font_size)
        p.font.color.rgb = RGBColor(0xE7, 0xEE, 0xFF) if dark else ORACLE_COLORS["text"]
        p.level = 0
        p.alignment = PP_ALIGN.LEFT


def add_footer_note(slide, text: str, x: float, y: float, w: float):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.2))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Oracle Sans Tab"
    p.font.size = Pt(7)
    p.font.color.rgb = ORACLE_COLORS["muted"]
    p.alignment = PP_ALIGN.LEFT


def add_section_label(slide, text: str, x: float, y: float):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(3.8), Inches(0.25))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Oracle Sans Tab"
    p.font.bold = True
    p.font.size = Pt(10)
    p.font.color.rgb = ORACLE_COLORS["muted"]


def add_insight_callout(slide, title: str, lines: list[str], x: float, y: float, w: float, h: float):
    shp = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = ORACLE_COLORS["panel"]
    shp.line.color.rgb = ORACLE_COLORS["line"]
    tf = shp.text_frame
    tf.clear()
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Oracle Sans Tab"
    p0.font.bold = True
    p0.font.size = Pt(10)
    p0.font.color.rgb = ORACLE_COLORS["text"]
    for line in lines[:3]:
        p = tf.add_paragraph()
        p.text = f"- {line}"
        p.font.name = "Oracle Sans Tab"
        p.font.size = Pt(9)
        p.font.color.rgb = ORACLE_COLORS["muted"]


def add_notes(slide, lines: list[str], x: float, y: float, w: float, h: float):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    for i, line in enumerate(lines[:3]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"- {line}"
        p.font.name = "Oracle Sans Tab"
        p.font.size = Pt(9)
        p.font.color.rgb = ORACLE_COLORS["muted"]


def add_table(
    slide,
    headers: list[str],
    rows: list[list[str]],
    x: float,
    y: float,
    w: float,
    h: float,
    numeric_cols: set[int] | None = None,
):
    numeric_cols = numeric_cols or set()
    t = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(x), Inches(y), Inches(w), Inches(h)).table
    for j, htxt in enumerate(headers):
        cell = t.cell(0, j)
        cell.text = htxt
        cell.fill.solid()
        cell.fill.fore_color.rgb = ORACLE_COLORS["accent1"]
        p = cell.text_frame.paragraphs[0]
        p.font.name = "Oracle Sans Tab"
        p.font.bold = True
        p.font.size = Pt(9)
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = PP_ALIGN.RIGHT if j in numeric_cols else PP_ALIGN.LEFT
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            cell = t.cell(i, j)
            cell.text = v
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xF2, 0xF5, 0xF7) if i % 2 else RGBColor(0xE8, 0xED, 0xF1)
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Oracle Sans Tab"
            p.font.size = Pt(8)
            p.font.color.rgb = ORACLE_COLORS["text"]
            p.alignment = PP_ALIGN.RIGHT if j in numeric_cols else PP_ALIGN.LEFT


def downsample(xs: list[str], ys: list[float], max_points: int = 36) -> tuple[list[str], list[float]]:
    n = min(len(xs), len(ys))
    if n <= max_points:
        return xs[:n], ys[:n]
    out_i = [round(i * (n - 1) / (max_points - 1)) for i in range(max_points)]
    out_i = sorted(set(out_i))
    return [xs[i] for i in out_i], [ys[i] for i in out_i]


def add_line_chart(
    slide,
    series: list[tuple[str, list[str], list[float]]],
    x: float,
    y: float,
    w: float,
    h: float,
    show_legend: bool = True,
    y_min: float | None = None,
    y_max: float | None = None,
    x_tick_font_size: int = 8,
    y_tick_font_size: int = 8,
):
    if not series:
        return
    chart_data = CategoryChartData()
    base_x = series[0][1]
    if not base_x:
        return
    labels, _ = downsample(base_x, series[0][2], 36)
    chart_data.categories = [str(t).replace("T", " ")[5:16] for t in labels]
    top = sorted(series, key=lambda s: max(s[2]) if s[2] else 0.0, reverse=True)[:8]
    for name, xs, ys in top:
        sx, sy = downsample(xs, ys, 36)
        # align by downsample count
        m = min(len(labels), len(sx), len(sy))
        chart_data.add_series(short_label(name, 24), sy[:m])
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        chart_data,
    ).chart
    chart.has_title = False
    chart.has_legend = bool(show_legend)
    if show_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        try:
            chart.legend.font.size = Pt(10)
        except Exception:
            pass
    if y_min is not None:
        chart.value_axis.minimum_scale = float(y_min)
    if y_max is not None:
        chart.value_axis.maximum_scale = float(y_max)
    chart.value_axis.has_major_gridlines = False
    chart.value_axis.has_minor_gridlines = False
    chart.category_axis.has_major_gridlines = False
    chart.category_axis.has_minor_gridlines = False
    chart.value_axis.tick_labels.font.size = Pt(y_tick_font_size)
    chart.category_axis.tick_labels.font.size = Pt(x_tick_font_size)
    for i, s in enumerate(chart.series):
        s.format.line.color.rgb = PALETTE[i % len(PALETTE)]
        s.format.line.width = Pt(2)


def add_bar_chart(
    slide,
    items: list[tuple[str, float]],
    x: float,
    y: float,
    w: float,
    h: float,
    series_name: str = "Value",
    x_tick_font_size: int = 8,
):
    vals = [(short_label(k, 22), float(v)) for k, v in items if float(v) > 0][:10]
    if not vals:
        return
    data = CategoryChartData()
    data.categories = [k for k, _ in vals]
    data.add_series(series_name, [v for _, v in vals])
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        data,
    ).chart
    chart.has_title = False
    chart.has_legend = False
    chart.value_axis.minimum_scale = 0
    chart.value_axis.has_major_gridlines = False
    chart.value_axis.has_minor_gridlines = False
    chart.category_axis.has_major_gridlines = False
    chart.category_axis.has_minor_gridlines = False
    chart.value_axis.tick_labels.font.size = Pt(8)
    chart.category_axis.tick_labels.font.size = Pt(x_tick_font_size)
    plot = chart.plots[0]
    plot.vary_by_categories = True


def add_vcpu_boxplot_chart(
    slide,
    stats: dict[str, float] | None,
    x: float,
    y: float,
    w: float,
    h: float,
    title_text: str = "CPU Consumption (Box Plot)",
    value_scale: float = 1.0,
):
    title = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.18))
    tf = title.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Oracle Sans Tab"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = ORACLE_COLORS["muted"]
    p.alignment = PP_ALIGN.LEFT

    if not stats:
        add_freshness_unavailable_note(slide, x, y + 0.28, w)
        return

    vmin = to_num(stats.get("min")) * value_scale
    p30 = to_num(stats.get("p30")) * value_scale
    p50 = to_num(stats.get("p50")) * value_scale
    p70 = to_num(stats.get("p70")) * value_scale
    vmax = to_num(stats.get("max")) * value_scale
    p95 = to_num(stats.get("p95")) * value_scale
    # Real Python boxplot (matplotlib), inserted as a rendered chart image.
    try:
        os.environ.setdefault("MPLCONFIGDIR", "/tmp/mplconfig")
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        fig_w = max(3.0, w * 1.6)
        fig_h = max(1.2, (h - 0.24) * 1.6)
        fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=180)
        bxpstats = [
            {
                "label": "",
                "whislo": vmin,
                "q1": p30,
                "med": p50,
                "q3": p70,
                "whishi": vmax,
                "fliers": [p95] if (p95 < vmin or p95 > vmax) else [],
            }
        ]
        b = ax.bxp(bxpstats, patch_artist=True, showfliers=True, widths=0.45)
        for box in b.get("boxes", []):
            box.set_facecolor("#DCECF0")
            box.set_edgecolor("#04536F")
            box.set_linewidth(1.2)
        for med in b.get("medians", []):
            med.set_color("#C74634")
            med.set_linewidth(1.4)
        for wk in b.get("whiskers", []) + b.get("caps", []):
            wk.set_color("#04536F")
            wk.set_linewidth(1.1)
        for fl in b.get("fliers", []):
            fl.set_marker("D")
            fl.set_markerfacecolor("#6C3F49")
            fl.set_markeredgecolor("#6C3F49")
            fl.set_markersize(4)

        ax.set_xticks([])
        ax.set_ylabel("")
        ax.tick_params(axis="y", labelsize=8, colors="#6B747A")
        for tl in ax.get_yticklabels():
            tl.set_fontname("Oracle Sans Tab")
        ax.grid(axis="y", linestyle="-", linewidth=0.6, color="#D9DEDE", alpha=0.85)
        for sp in ("top", "right"):
            ax.spines[sp].set_visible(False)
        ax.spines["left"].set_color("#D9DEDE")
        ax.spines["bottom"].set_color("#D9DEDE")
        ax.set_facecolor((1, 1, 1, 0))
        fig.patch.set_alpha(0.0)
        fig.tight_layout(pad=0.4)

        png = io.BytesIO()
        fig.savefig(png, format="png", dpi=180, transparent=True)
        plt.close(fig)
        png.seek(0)
        slide.shapes.add_picture(
            png,
            Inches(x),
            Inches(y + 0.22),
            Inches(w),
            Inches(max(0.2, h - 0.24)),
        )
    except Exception:
        # Fully native editable fallback chart with percentile distribution.
        data = CategoryChartData()
        data.categories = ["Min", "P30", "P50", "P70", "P95", "Max"]
        data.add_series("CPU", [vmin, p30, p50, p70, p95, vmax])
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(x),
            Inches(y + 0.22),
            Inches(w),
            Inches(max(0.2, h - 0.24)),
            data,
        ).chart
        chart.has_legend = False
        try:
            chart.value_axis.tick_labels.font.size = Pt(8)
            chart.category_axis.tick_labels.font.size = Pt(8)
        except Exception:
            pass

    # Show factual reference values so users can interpret the OHLC mapping.
    note = slide.shapes.add_textbox(Inches(x), Inches(y + h - 0.16), Inches(w), Inches(0.16))
    ntf = note.text_frame
    ntf.clear()
    np = ntf.paragraphs[0]
    np.text = f"Min {fmt_num(vmin,3)} | P30 {fmt_num(p30,3)} | Median {fmt_num(p50,3)} | P70 {fmt_num(p70,3)} | Max {fmt_num(vmax,3)}"
    np.font.name = "Oracle Sans Tab"
    np.font.size = Pt(8)
    np.font.color.rgb = ORACLE_COLORS["muted"]
    np.alignment = PP_ALIGN.LEFT


def add_memory_compare_chart(slide, total_host_mem_gb: float, instance_used_mem_gb: float, x: float, y: float, w: float, h: float):
    server_mem = max(0.0, float(total_host_mem_gb))
    inst_mem = max(0.0, float(instance_used_mem_gb))

    # Base native chart (column): instance dedicated memory.
    col_data = CategoryChartData()
    col_data.categories = ["Server", "Instance"]
    col_data.add_series("Instance Dedicated Memory", [0.0, inst_mem])
    col_chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        col_data,
    ).chart
    col_chart.has_legend = False
    col_chart.value_axis.minimum_scale = 0
    col_chart.value_axis.tick_labels.font.size = Pt(8)
    col_chart.category_axis.tick_labels.font.size = Pt(8)
    try:
        s0 = col_chart.series[0]
        s0.points[0].format.fill.solid()
        s0.points[0].format.fill.fore_color.rgb = RGBColor(0xE6, 0xEA, 0xEF)
        s0.points[1].format.fill.solid()
        s0.points[1].format.fill.fore_color.rgb = ORACLE_COLORS["accent3"]
    except Exception:
        pass

    # Overlay native chart (line): server total memory.
    line_data = CategoryChartData()
    line_data.categories = ["Server", "Instance"]
    line_data.add_series("Server Total Memory", [server_mem, server_mem])
    line_chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        line_data,
    ).chart
    line_chart.has_legend = False
    line_chart.has_title = False
    line_chart.value_axis.minimum_scale = 0
    line_chart.value_axis.maximum_scale = max(server_mem, inst_mem) * 1.1 if max(server_mem, inst_mem) > 0 else 1.0
    line_chart.value_axis.tick_labels.font.size = Pt(8)
    line_chart.category_axis.tick_labels.font.size = Pt(8)
    line_chart.category_axis.visible = False
    try:
        line_chart.value_axis.visible = False
    except Exception:
        pass
    try:
        ls = line_chart.series[0]
        ls.format.line.color.rgb = ORACLE_COLORS["accent1"]
        ls.marker.style = 2  # square
    except Exception:
        pass


def add_hbar_chart(slide, items: list[tuple[str, float]], x: float, y: float, w: float, h: float, series_name: str = "Value"):
    vals = [(short_label(k, 20), float(v)) for k, v in items if float(v) > 0][:8]
    if not vals:
        return
    data = CategoryChartData()
    data.categories = [k for k, _ in vals]
    data.add_series(series_name, [v for _, v in vals])
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        data,
    ).chart
    chart.has_legend = False
    chart.value_axis.minimum_scale = 0
    chart.category_axis.reverse_order = True
    plot = chart.plots[0]
    plot.vary_by_categories = True


def add_pie_chart(slide, items: list[tuple[str, float]], x: float, y: float, w: float, h: float, series_name: str = "Distribution"):
    vals = [(short_label(k, 22), float(v)) for k, v in items if float(v) > 0][:8]
    if not vals:
        return
    data = CategoryChartData()
    data.categories = [k for k, _ in vals]
    data.add_series(series_name, [v for _, v in vals])
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        data,
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT


def add_freshness_timeline_chart(
    slide,
    ranges: list[tuple[str, datetime, datetime]],
    x: float,
    y: float,
    w: float,
    h: float,
    legend_font_size: int = 10,
):
    if not ranges:
        return

    def to_excel_serial(dt: datetime) -> float:
        epoch = datetime(1899, 12, 30)
        delta = dt - epoch
        return delta.days + (delta.seconds + delta.microseconds / 1_000_000.0) / 86400.0

    global_start = min(r[1] for r in ranges)
    global_end = max(r[2] for r in ranges)
    if global_end <= global_start:
        global_end = global_start + timedelta(hours=1)
    min_serial = to_excel_serial(global_start)
    max_serial = to_excel_serial(global_end)
    if max_serial <= min_serial:
        max_serial = min_serial + (1.0 / 24.0)
    xy = XyChartData()
    for idx, (cohort, c_start, c_end) in enumerate(ranges, start=1):
        s = xy.add_series(short_label(cohort, 18))
        x0 = max(to_excel_serial(c_start), min_serial)
        x1 = max(to_excel_serial(c_end), x0 + (1.0 / 1440.0))
        yv = float(idx)
        s.add_data_point(x0, yv)
        s.add_data_point(x1, yv)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        xy,
    ).chart
    chart.has_legend = False
    chart.value_axis.minimum_scale = 0.0
    chart.value_axis.maximum_scale = float(len(ranges) + 1)
    chart.value_axis.major_unit = 1.0
    chart.value_axis.has_major_gridlines = False
    chart.category_axis.minimum_scale = min_serial
    chart.category_axis.maximum_scale = max_serial
    chart.category_axis.tick_labels.number_format = "mmm/dd"
    chart.category_axis.tick_labels.font.size = Pt(8)
    try:
        chart.value_axis.visible = False
    except Exception:
        chart.value_axis.tick_labels.font.size = Pt(1)
        chart.value_axis.format.line.fill.background()
    for i, s in enumerate(chart.series):
        s.format.line.color.rgb = PALETTE[i % len(PALETTE)]
        s.smooth = False
        # Label only the last point with the series (cohort/instance) name.
        try:
            last_point = s.points[len(s.points) - 1]
            dl = last_point.data_label
            dl.has_text_frame = True
            dl.font.name = "Oracle Sans Tab"
            dl.font.size = Pt(4)
            dl.font.color.rgb = ORACLE_COLORS["muted"]
            tf = dl.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = short_label(ranges[i][0], 18)
            run.font.name = "Oracle Sans Tab"
            run.font.size = Pt(4)
            run.font.color.rgb = ORACLE_COLORS["muted"]
        except Exception:
            pass

    # Axis label helper so time context is explicit.
    add_footer_note(
        slide,
        f"Timeline range: {global_start.strftime('%Y-%m-%d %H:%M')} to {global_end.strftime('%Y-%m-%d %H:%M')}",
        x=x,
        y=y + h - 0.02,
        w=w,
    )


def add_freshness_unavailable_note(slide, x: float, y: float, w: float):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.26))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = "AWR Data Freshness Timeline: no CPU timeline data available in current scope."
    p.font.name = "Oracle Sans Tab"
    p.font.size = Pt(11)
    p.font.color.rgb = ORACLE_COLORS["muted"]
    p.alignment = PP_ALIGN.LEFT


def add_cohort_cpu_boxplot_chart(slide, cohort_stats: list[tuple[str, dict[str, float]]], x: float, y: float, w: float, h: float):
    if not cohort_stats:
        add_freshness_unavailable_note(slide, x, y + 0.04, w)
        return
    try:
        os.environ.setdefault("MPLCONFIGDIR", "/tmp/mplconfig")
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        items = sorted(cohort_stats, key=lambda t: t[0])
        labels = [short_label(c, 10) for c, _ in items]
        bxpstats = []
        for c, st in items:
            bxpstats.append(
                {
                    "label": short_label(c, 10),
                    "whislo": to_num(st.get("min")),
                    "q1": to_num(st.get("p30")),
                    "med": to_num(st.get("p50")),
                    "q3": to_num(st.get("p70")),
                    "whishi": to_num(st.get("max")),
                    "fliers": [],
                }
            )

        fig_w = max(2.4, w * 1.8)
        fig_h = max(1.0, h * 1.9)
        fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=180)
        b = ax.bxp(bxpstats, patch_artist=True, showfliers=False, widths=0.55)
        for i, box in enumerate(b.get("boxes", [])):
            box.set_facecolor("#DCECF0")
            box.set_edgecolor("#04536F")
            box.set_linewidth(1.1)
        for med in b.get("medians", []):
            med.set_color("#C74634")
            med.set_linewidth(1.3)
        for wk in b.get("whiskers", []) + b.get("caps", []):
            wk.set_color("#04536F")
            wk.set_linewidth(1.0)

        ax.set_ylabel("")
        ax.set_xticks(range(1, len(labels) + 1))
        ax.set_xticklabels(labels, fontsize=7, fontname="Oracle Sans Tab", color="#6B747A", rotation=0)
        ax.tick_params(axis="y", labelsize=7, colors="#6B747A")
        for tl in ax.get_yticklabels():
            tl.set_fontname("Oracle Sans Tab")
        ax.grid(axis="y", linestyle="-", linewidth=0.5, color="#D9DEDE", alpha=0.85)
        for sp in ("top", "right"):
            ax.spines[sp].set_visible(False)
        ax.spines["left"].set_color("#D9DEDE")
        ax.spines["bottom"].set_color("#D9DEDE")
        ax.set_facecolor((1, 1, 1, 0))
        fig.patch.set_alpha(0.0)
        fig.tight_layout(pad=0.25)

        png = io.BytesIO()
        fig.savefig(png, format="png", dpi=180, transparent=True)
        plt.close(fig)
        png.seek(0)
        slide.shapes.add_picture(png, Inches(x), Inches(y), Inches(w), Inches(h))
    except Exception:
        add_freshness_unavailable_note(slide, x, y + 0.04, w)


def build_plan(model: Model) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    out.append({"id": "title", "type": "title", "title": "AWR Analysis", "subtitle": "Generated report overview"})
    out.append(
        {
            "id": "summary-global",
            "type": "summary",
            "title": "Global Summary",
            "subtitle": "All selected databases and cohorts",
        }
    )
    out.append({"id": "sep-cohorts", "type": "separator", "title": "Cohort Analysis", "subtitle": "Cohort-level drill down"})
    for c in model.cohorts():
        out.append({"id": f"cohort:{c}", "type": "cohort", "cohort": c, "title": f"Cohort: {c}", "subtitle": f"Summary for {c}"})
    out.append({"id": "sep-instances", "type": "separator", "title": "Instance Analysis", "subtitle": "Instance-level drill down"})
    for c in model.cohorts():
        out.append({"id": f"sep-instance-cohort:{c}", "type": "separator", "title": f"Instances - {c}", "subtitle": f"Instances in {c}"})
        rs = sorted(model.rows_for_cohort(c), key=lambda r: r.instance)
        for r in rs:
            db_name = model.display_db_name(r.db)
            inst_label = model.instance_label(r.instance)
            out.append(
                {
                    "id": f"instance:{c}:{r.instance}",
                    "type": "instance",
                    "cohort": c,
                    "instance": r.instance,
                    "db": r.db,
                    "title": f"Instance: {inst_label}",
                    "subtitle": f"{db_name} on {r.host}",
                }
            )
    out.append({"id": "closing", "type": "closing", "title": "Thank You", "subtitle": "AWR analysis completed"})
    selected = set(model.ui_state.get("pptSlidesSelected") or []) if isinstance(model.ui_state, dict) else set()
    if selected:
        out = [x for x in out if x.get("id") in selected]
    return out


def export_ppt(input_path: Path, template_path: Path, output_path: Path):
    payload = json.loads(input_path.read_text(encoding="utf-8"))
    if isinstance(payload, dict) and "raw_data" in payload:
        raw = payload.get("raw_data") or {}
        ui = payload.get("ui_state") or {}
    else:
        raw = payload
        ui = {}
    model = Model(raw, ui)

    prs = presentation_from_template(template_path)
    # Prefer explicit layout mapping from template sample slides:
    # slide 1 = title, 2 = divider, 3 = content, 4 = closing.
    template_slides = list(prs.slides)
    layout_title = template_slides[0].slide_layout if len(template_slides) >= 1 else find_layout(prs, ["title"])
    layout_section = template_slides[1].slide_layout if len(template_slides) >= 2 else find_layout(prs, ["divider", "section"])
    layout_content = template_slides[2].slide_layout if len(template_slides) >= 3 else find_layout(prs, ["content", "1 column", "title/subtitle"])
    layout_closing = template_slides[3].slide_layout if len(template_slides) >= 4 else find_layout(prs, ["closing", "thank"])
    layout_blank_light = find_layout(prs, ["light - blank", "blank", "title 1 column"])
    clear_existing_slides(prs)

    plan = build_plan(model)
    for item in plan:
        t = item["type"]
        if t == "title":
            slide = prs.slides.add_slide(layout_title)
            set_title_and_subtitle(slide, item["title"], f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            bx = slide.shapes.add_textbox(Inches(1.1), Inches(2.5), Inches(10.9), Inches(1.1))
            tf = bx.text_frame
            tf.text = "Executive report with global, cohort and instance drill downs."
            tf.paragraphs[0].font.name = "Oracle Sans Tab"
            tf.paragraphs[0].font.size = Pt(16)
            tf.paragraphs[0].font.color.rgb = ORACLE_COLORS["text"]
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            continue

        if t == "separator":
            slide = prs.slides.add_slide(layout_section)
            set_title_and_subtitle(slide, item["title"], item.get("subtitle", ""))
            continue

        if t == "summary":
            slide = prs.slides.add_slide(layout_content)
            s = model.summary()
            versions = [(k, v) for k, v in model.version_counts()]
            top_ver = versions[0] if versions else ("N/A", 0)
            total_ver = sum(v for _, v in versions) or 1
            cpu_series = model.cpu_series_global()
            all_cpu_points = [v for _, _, ys in cpu_series for v in ys]
            peak_cpu = max(all_cpu_points) if all_cpu_points else None
            p95_cpu = None
            if all_cpu_points:
                sorted_cpu = sorted(all_cpu_points)
                p95_cpu = quantile(sorted_cpu, 0.95)
            total_db = int(s["db_count"])
            total_hosts = int(s["host_count"])
            total_instances = int(s["instance_count"])
            total_cohorts = len(model.cohorts())
            freshness = model.freshness_ranges()
            cpu_cohort_stats = model.cpu_stats_by_cohort()
            mem_items = model.memory_by_cohort()
            iops_items = model.iops_by_cohort()

            title_msg = "Summary of AWR Capture."
            subtitle_msg = f"Instances: {total_instances} (Databases: {total_db}, Hosts: {total_hosts})"
            set_title_and_subtitle(slide, title_msg, subtitle_msg)

            cpu_value = f"{peak_cpu:.1f} vCPU" if peak_cpu is not None else "N/A"
            cpu_sub = f"P95: {p95_cpu:.1f} vCPU" if p95_cpu is not None else "P95: N/A"
            add_exec_cards(
                slide,
                [
                    {
                        "value": fmt_num(total_instances, 0),
                        "label": "Total Instances",
                        "sub": f"{fmt_num(total_db,0)} databases",
                        "color": ORACLE_COLORS["accent1"],
                    },
                    {
                        "value": fmt_num(total_hosts, 0),
                        "label": "Total Hosts",
                        "sub": f"{fmt_num(total_cohorts,0)} cohorts",
                        "color": ORACLE_COLORS["accent1"],
                    },
                    {
                        "value": f"{s['memory_gb_total']/1024.0:.1f} TB",
                        "label": "Total Memory",
                        "sub": f"{fmt_num(s['memory_gb_total'],1)} GB",
                        "color": ORACLE_COLORS["accent5"],
                    },
                    {
                        "value": fmt_num(s["vcpu_total"], 1),
                        "label": "Total vCPU",
                        "sub": "Configured capacity",
                        "color": ORACLE_COLORS["accent3"],
                    },
                    {
                        "value": f"{s['allocated_storage_gb']/1024.0:.1f} TB",
                        "label": "Allocated Storage",
                        "sub": f"{fmt_num(s['allocated_storage_gb'],1)} GB",
                        "color": ORACLE_COLORS["accent2"],
                    },
                ],
            )
            versions_lines = []
            top_versions = versions[:4]
            for ver, count in top_versions:
                pct = 100.0 * count / max(total_db, 1)
                versions_lines.append(f"{ver}: {count}/{total_db} DBs ({pct:.1f}%)")
            if len(versions) > 4:
                others = sum(c for _, c in versions[4:])
                pct = 100.0 * others / max(total_db, 1)
                versions_lines.append(f"Others: {others}/{total_db} DBs ({pct:.1f}%)")
            add_big_versions_card(
                slide,
                "Database Versions",
                versions_lines or ["No version data available"],
                x=8.6,
                y=3.15,
                w=4.2,
                h=1.58,
            )

            cohorts_sorted = sorted(model.by_cohort_table(), key=lambda r: r["mem"], reverse=True)
            top2_mem = sum(r["mem"] for r in cohorts_sorted[:2])
            total_mem = sum(r["mem"] for r in cohorts_sorted) or 1.0
            learned = [
                f"Selected scope: {total_db} DBs, {total_instances} instances, {total_hosts} hosts, {total_cohorts} cohorts.",
                f"Top DB version: {top_ver[0]} with {int(top_ver[1])} databases ({100.0 * top_ver[1] / max(total_db,1):.1f}%).",
                f"Aggregate capacity: {fmt_num(s['memory_gb_total'],1)} GB memory, {fmt_num(s['vcpu_total'],1)} vCPU, {fmt_num(s['allocated_storage_gb'],1)} GB allocated storage.",
                f"Top 2 cohorts account for {100.0 * top2_mem / total_mem:.1f}% of selected memory.",
            ]
            add_panel(slide, "What We Learned", learned, x=0.79, y=3.15, w=7.6, h=1.58, dark=False, body_font_size=14)
            c_y = 5.05
            c_h = 1.65
            c_gap = 0.12
            c_w = (12.0 - (3 * c_gap)) / 4.0
            c1_x = 0.79
            c2_x = c1_x + c_w + c_gap
            c3_x = c2_x + c_w + c_gap
            c4_x = c3_x + c_w + c_gap

            add_section_label(slide, "AWR capture timeline", c1_x, 4.88)
            add_section_label(slide, "Memory", c2_x, 4.88)
            add_section_label(slide, "IOPS", c3_x, 4.88)
            add_section_label(slide, "CPU Box Plot", c4_x, 4.88)
            if freshness:
                add_freshness_timeline_chart(slide, freshness, x=c1_x, y=c_y, w=c_w, h=c_h, legend_font_size=4)
            else:
                add_freshness_unavailable_note(slide, x=c1_x, y=5.62, w=c_w)
            add_bar_chart(slide, mem_items, x=c2_x, y=c_y, w=c_w, h=c_h, series_name="Memory (GB)", x_tick_font_size=4)
            add_bar_chart(slide, iops_items, x=c3_x, y=c_y, w=c_w, h=c_h, series_name="DB IOPS", x_tick_font_size=4)
            add_cohort_cpu_boxplot_chart(slide, cpu_cohort_stats, x=c4_x, y=c_y, w=c_w, h=c_h)
            add_footer_note(
                slide,
                "Source: instances, database_statistics, plots_html | Metrics shown for selected PPT scope",
                x=0.8,
                y=6.92,
                w=11.9,
            )
            continue

        if t == "cohort":
            c = item["cohort"]
            slide = prs.slides.add_slide(layout_content)
            rs = model.rows_for_cohort(c)
            sm = model.summary(rs)
            set_title_and_subtitle(slide, item["title"], f"Databases: {int(sm['db_count'])} | Instances: {int(sm['instance_count'])}")
            version_dbs: dict[str, set[str]] = defaultdict(set)
            for r in rs:
                version_dbs[model.db_versions.get(r.db, "Unknown")].add(r.db)
            versions = sorted(((ver, len(dbs)) for ver, dbs in version_dbs.items()), key=lambda kv: (-kv[1], kv[0]))
            top_ver = versions[0] if versions else ("Unknown", 0)
            total_db = int(sm["db_count"])
            total_hosts = int(sm["host_count"])
            total_instances = int(sm["instance_count"])

            add_exec_cards(
                slide,
                [
                    {
                        "value": f"{total_instances:d}",
                        "label": "Total Instances",
                        "sub": f"Across {total_db:d} DBs",
                        "color": ORACLE_COLORS["accent1"],
                    },
                    {
                        "value": f"{total_hosts:d}",
                        "label": "Total Hosts",
                        "sub": f"{total_instances:d} instances distributed",
                        "color": ORACLE_COLORS["accent3"],
                    },
                    {
                        "value": f"{sm['memory_gb_total']/1024.0:.1f} TB",
                        "label": "Total Memory",
                        "sub": f"{fmt_num(sm['memory_gb_total'],1)} GB",
                        "color": ORACLE_COLORS["accent2"],
                    },
                    {
                        "value": f"{fmt_num(sm['vcpu_total'],1)}",
                        "label": "Total vCPU",
                        "sub": "Configured capacity",
                        "color": ORACLE_COLORS["accent1"],
                    },
                    {
                        "value": f"{sm['allocated_storage_gb']/1024.0:.1f} TB",
                        "label": "Allocated Storage",
                        "sub": f"{fmt_num(sm['allocated_storage_gb'],1)} GB",
                        "color": ORACLE_COLORS["accent2"],
                    },
                ],
            )
            versions_lines = []
            top_versions = versions[:4]
            for ver, count in top_versions:
                pct = 100.0 * count / max(total_db, 1)
                versions_lines.append(f"{ver}: {count}/{total_db} DBs ({pct:.1f}%)")
            if len(versions) > 4:
                others = sum(cn for _, cn in versions[4:])
                pct = 100.0 * others / max(total_db, 1)
                versions_lines.append(f"Others: {others}/{total_db} DBs ({pct:.1f}%)")
            add_big_versions_card(
                slide,
                "Database Versions",
                versions_lines or ["No version data available"],
                x=8.6,
                y=3.15,
                w=4.2,
                h=1.58,
            )

            db_cnt = defaultdict(int)
            for r in rs:
                db_cnt[r.db] += 1
            top2_db = sorted(db_cnt.items(), key=lambda kv: (-kv[1], kv[0]))[:2]
            top2_instances = sum(n for _, n in top2_db)
            learned = [
                f"Cohort scope: {total_db} DBs, {total_instances} instances, {total_hosts} hosts.",
                f"Top DB version: {top_ver[0]} with {int(top_ver[1])} databases ({100.0 * top_ver[1] / max(total_db,1):.1f}%).",
                f"Aggregate capacity: {fmt_num(sm['memory_gb_total'],1)} GB memory, {fmt_num(sm['vcpu_total'],1)} vCPU, {fmt_num(sm['allocated_storage_gb'],1)} GB allocated storage.",
                f"Top 2 DBs in cohort account for {100.0 * top2_instances / max(total_instances,1):.1f}% of instances.",
            ]
            add_panel(slide, "What We Learned", learned, x=0.79, y=3.15, w=7.6, h=1.58, dark=False, body_font_size=14)

            cohort_freshness = model.freshness_ranges_by_instance_in_cohort(c)
            mem_items_c = model.metric_by_db(c, "DB Memory (MB)", scale=(1.0 / 1024.0))
            iops_items_c = model.metric_by_db(c, "DB IOPS", scale=1.0)
            inst_stats = []
            for inst_name, _xs, ys in model.cpu_series_instances_in_cohort(c):
                st = stats_from_values(ys)
                if st:
                    inst_stats.append((inst_name, st))

            c_y = 5.05
            c_h = 1.65
            c_gap = 0.12
            c_w = (12.0 - (3 * c_gap)) / 4.0
            c1_x = 0.79
            c2_x = c1_x + c_w + c_gap
            c3_x = c2_x + c_w + c_gap
            c4_x = c3_x + c_w + c_gap

            add_section_label(slide, "AWR capture timeline", c1_x, 4.88)
            add_section_label(slide, "Memory", c2_x, 4.88)
            add_section_label(slide, "IOPS", c3_x, 4.88)
            add_section_label(slide, "CPU Box Plot", c4_x, 4.88)
            if cohort_freshness:
                add_freshness_timeline_chart(slide, cohort_freshness, x=c1_x, y=c_y, w=c_w, h=c_h, legend_font_size=4)
            else:
                add_freshness_unavailable_note(slide, x=c1_x, y=5.62, w=c_w)
            add_bar_chart(slide, mem_items_c, x=c2_x, y=c_y, w=c_w, h=c_h, series_name="Memory (GB)", x_tick_font_size=4)
            add_bar_chart(slide, iops_items_c, x=c3_x, y=c_y, w=c_w, h=c_h, series_name="DB IOPS", x_tick_font_size=4)
            add_cohort_cpu_boxplot_chart(slide, inst_stats, x=c4_x, y=c_y, w=c_w, h=c_h)
            continue

        if t == "instance":
            slide = prs.slides.add_slide(layout_content)
            c = item["cohort"]
            inst = item["instance"]
            db = item["db"]
            rs = [r for r in model.rows_for_cohort(c) if r.instance == inst]
            row = rs[0] if rs else None
            db_name = model.display_db_name(db)
            db_type, pdb_count = model.db_type_and_pdb_count(db)
            db_version = model.db_version(db)
            is_rac = "Yes" if model.db_is_rac(db) else "No"
            set_title_and_subtitle(slide, item["title"], item["subtitle"])
            add_kpi_boxes(
                slide,
                [
                    ("Cohort", c),
                    ("Database", db_name),
                    ("DB Version", db_version),
                    ("DB Type", db_type),
                    ("PDBs", pdb_count),
                    ("RAC", is_rac),
                    ("Host", row.host if row else "N/A"),
                    ("vCPU", fmt_num((row.init_cpu_count if row else 0), 1)),
                    ("Memory (GB)", fmt_num((row.mem_gb if row else 0), 1)),
                    ("SGA (GB)", fmt_num((row.sga_gb if row else 0), 1)),
                    ("PGA (GB)", fmt_num((row.pga_gb if row else 0), 1)),
                ],
            )

            mem_title = slide.shapes.add_textbox(Inches(0.79), Inches(3.40), Inches(6.2), Inches(0.24))
            mem_tf = mem_title.text_frame
            mem_tf.clear()
            mem_p = mem_tf.paragraphs[0]
            mem_p.text = "Memory Consumption"
            mem_p.font.name = "Oracle Sans Tab"
            mem_p.font.bold = True
            mem_p.font.size = Pt(14)
            mem_p.font.color.rgb = ORACLE_COLORS["muted"]
            mem_p.alignment = PP_ALIGN.LEFT

            host_mem = to_num(row.mem_gb if row else 0.0)
            inst_mem = to_num((row.sga_gb if row else 0.0) + (row.pga_gb if row else 0.0))
            cpu_inst_series = [x for x in model.cpu_series_instances_in_cohort(c) if x[0] == inst]
            xs_mem = cpu_inst_series[0][1] if cpu_inst_series else (model.cpu_by_cohort.get(c)[0] if model.cpu_by_cohort.get(c) else [])
            mem_series = []
            if xs_mem:
                mem_series = [
                    ("Server Memory (GB)", xs_mem, [host_mem] * len(xs_mem)),
                    ("Instance Memory (GB)", xs_mem, [inst_mem] * len(xs_mem)),
                ]
            add_line_chart(
                slide,
                mem_series,
                x=0.79,
                y=3.58,
                w=8.05,
                h=1.65,
                show_legend=True,
                y_min=None,
                y_max=None,
                x_tick_font_size=4,
                y_tick_font_size=8,
            )
            mem_stats = model.db_metric_stats(c, db, "DB Memory (MB)")
            add_vcpu_boxplot_chart(
                slide,
                mem_stats,
                x=8.95,
                y=3.40,
                w=3.84,
                h=1.83,
                title_text="Memory Consumption (Box Plot)",
                value_scale=(1.0 / 1024.0),
            )

            cpu_title = slide.shapes.add_textbox(Inches(0.79), Inches(5.42), Inches(4.6), Inches(0.24))
            cpu_tf = cpu_title.text_frame
            cpu_tf.clear()
            cpu_p = cpu_tf.paragraphs[0]
            cpu_p.text = "CPU Consumption"
            cpu_p.font.name = "Oracle Sans Tab"
            cpu_p.font.bold = True
            cpu_p.font.size = Pt(14)
            cpu_p.font.color.rgb = ORACLE_COLORS["muted"]
            cpu_p.alignment = PP_ALIGN.LEFT
            add_line_chart(
                slide,
                cpu_inst_series,
                x=0.79,
                y=5.60,
                w=8.05,
                h=1.65,
                show_legend=False,
                y_min=None,
                y_max=None,
                x_tick_font_size=4,
                y_tick_font_size=8,
            )
            # Use the same sampled instance series as the line chart to keep boxplot and line fully consistent.
            vcpu_stats = None
            if cpu_inst_series and cpu_inst_series[0][1] and cpu_inst_series[0][2]:
                _sx, _sy = downsample(cpu_inst_series[0][1], cpu_inst_series[0][2], 36)
                vcpu_stats = stats_from_values(_sy)
            add_vcpu_boxplot_chart(
                slide,
                vcpu_stats,
                x=8.95,
                y=5.42,
                w=3.84,
                h=1.83,
                title_text="CPU Consumption (Box Plot)",
                value_scale=1.0,
            )
            continue

        if t == "closing":
            slide = prs.slides.add_slide(layout_closing)
            set_title_and_subtitle(slide, item["title"], item.get("subtitle", ""))
            continue

        # Safety fallback
        slide = prs.slides.add_slide(layout_blank_light)
        set_title_and_subtitle(slide, item.get("title", "Slide"), item.get("subtitle", ""))

    prs.save(str(output_path))


def main():
    ap = argparse.ArgumentParser(description="Export AWR analysis PPT using Oracle template white layouts.")
    ap.add_argument("--input", required=True, help="Path to raw AWR JSON or saved report JSON.")
    ap.add_argument("--template", required=True, help="Path to Oracle .potx template.")
    ap.add_argument("--output", required=True, help="Output .pptx path.")
    args = ap.parse_args()
    export_ppt(Path(args.input), Path(args.template), Path(args.output))
    print(args.output)


if __name__ == "__main__":
    main()
