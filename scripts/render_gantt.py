#!/usr/bin/env python3
"""
CPM Schedule Gantt Chart Renderer

Reads a schedule CSV, computes a forward-pass CPM schedule, and renders
a Gantt chart with WBS color coding and critical path highlighting.

Usage:
  python3 scripts/render_gantt.py --schedule data/example/summit-tower-activities.csv
  python3 scripts/render_gantt.py --schedule data/example/summit-tower-activities.csv --output outputs/gantt.png
"""

from __future__ import annotations

import argparse
import csv
from collections import deque
from dataclasses import dataclass, field
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import Patch
import numpy as np

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs"

# ── Color palette — auto-assigned to WBS codes in order encountered ──────
PALETTE = [
    "#264653", "#2A9D8F", "#457B9D", "#8D99AE", "#E76F51",
    "#F4A261", "#6A994E", "#BC6C25", "#7B2CBF", "#3B82F6",
    "#06B6D4", "#DC2626", "#D97706", "#16A34A", "#A855F7",
]


@dataclass
class Activity:
    wbs: str
    wbs_name: str
    activity_id: str
    activity_name: str
    duration_days: int
    predecessors: list[str]
    successors: list[str] = field(default_factory=list)
    start_day: int = 0
    finish_day: int = 0


def load_activities(path: Path) -> tuple[dict[str, Activity], list[str]]:
    activities: dict[str, Activity] = {}
    order: list[str] = []
    with path.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            preds = [p.strip() for p in row["predecessors"].split(",") if p.strip()]
            act = Activity(
                wbs=row["wbs"], wbs_name=row["wbs_name"],
                activity_id=row["activity_id"], activity_name=row["activity_name"],
                duration_days=int(row["duration_days"]), predecessors=preds,
            )
            activities[act.activity_id] = act
            order.append(act.activity_id)
    for act in activities.values():
        for pred_id in act.predecessors:
            activities[pred_id].successors.append(act.activity_id)
    return activities, order


def topological_order(activities: dict[str, Activity], csv_order: list[str]) -> list[str]:
    idx = {aid: i for i, aid in enumerate(csv_order)}
    indegree = {aid: len(a.predecessors) for aid, a in activities.items()}
    ready = deque(sorted([aid for aid, d in indegree.items() if d == 0], key=idx.get))
    ordered: list[str] = []
    while ready:
        cur = ready.popleft()
        ordered.append(cur)
        for succ in sorted(activities[cur].successors, key=idx.get):
            indegree[succ] -= 1
            if indegree[succ] == 0:
                ready.append(succ)
    return ordered


def compute_schedule(activities: dict[str, Activity], ordered_ids: list[str]) -> None:
    for aid in ordered_ids:
        act = activities[aid]
        if act.predecessors:
            act.start_day = max(activities[p].finish_day for p in act.predecessors)
        else:
            act.start_day = 0
        act.finish_day = act.start_day + act.duration_days


def assign_wbs_colors(activities: dict[str, Activity], display_ids: list[str]) -> dict[str, str]:
    """Assign colors to WBS codes in the order they first appear."""
    wbs_colors: dict[str, str] = {}
    color_idx = 0
    for aid in display_ids:
        wbs = activities[aid].wbs
        if wbs not in wbs_colors:
            wbs_colors[wbs] = PALETTE[color_idx % len(PALETTE)]
            color_idx += 1
    return wbs_colors


def render_chart(
    activities: dict[str, Activity], ordered_ids: list[str],
    display_ids: list[str], out_path: Path,
) -> None:
    wbs_colors = assign_wbs_colors(activities, display_ids)

    labels = [f"{aid}  {activities[aid].activity_name}" for aid in display_ids]
    starts = [activities[aid].start_day for aid in display_ids]
    durations = [activities[aid].duration_days for aid in display_ids]
    colors = [wbs_colors[activities[aid].wbs] for aid in display_ids]

    project_finish = max(activities[aid].finish_day for aid in ordered_ids)

    fig_height = max(11, len(display_ids) * 0.34)
    fig, ax = plt.subplots(figsize=(18, fig_height))

    y_positions = list(range(len(display_ids)))
    ax.barh(y_positions, durations, left=starts, color=colors, edgecolor="#1f1f1f", linewidth=0.5)

    for i, aid in enumerate(display_ids):
        act = activities[aid]
        ax.text(act.finish_day + 4, i, f"{act.start_day}-{act.finish_day}",
                va="center", ha="left", fontsize=8, color="#222222")

    ax.set_yticks(y_positions)
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("Working Day")
    ax.set_ylabel("Activity")
    ax.set_title("CPM Schedule — Finish-to-Start, Zero-Lag Logic")
    ax.set_xlim(0, project_finish + 80)
    ax.grid(axis="x", linestyle="--", linewidth=0.5, alpha=0.45)

    # Legend
    seen: set[str] = set()
    legend_items = []
    for aid in display_ids:
        wbs = activities[aid].wbs
        if wbs not in seen:
            seen.add(wbs)
            legend_items.append(Patch(
                color=wbs_colors[wbs],
                label=f"WBS {wbs} - {activities[aid].wbs_name}",
            ))
    ax.legend(handles=legend_items, loc="upper right", fontsize=8, frameon=True)

    # Summary
    open_ends = sorted([a.activity_id for a in activities.values()
                        if not a.successors and a.finish_day < project_finish])
    summary = f"Computed finish: WD {project_finish}"
    summary += f" | Open ends: {', '.join(open_ends) if open_ends else 'none'}"
    fig.text(0.01, 0.01, summary, fontsize=9)

    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(out_path, dpi=200)
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser(description="CPM Schedule Gantt Chart Renderer")
    parser.add_argument("--schedule", type=Path, required=True,
                        help="Path to schedule CSV")
    parser.add_argument("--output", type=Path, default=None,
                        help="Output PNG path (default: outputs/gantt.png)")
    args = parser.parse_args()

    out_path = args.output or (OUT / "gantt.png")
    out_path.parent.mkdir(exist_ok=True, parents=True)

    activities, csv_order = load_activities(args.schedule)
    ordered_ids = topological_order(activities, csv_order)
    compute_schedule(activities, ordered_ids)
    render_chart(activities, ordered_ids, csv_order, out_path)

    project_finish = max(activities[aid].finish_day for aid in ordered_ids)
    print(f"Gantt chart saved: {out_path}")
    print(f"  {len(activities)} activities | Project finish: WD {project_finish}")


if __name__ == "__main__":
    main()
