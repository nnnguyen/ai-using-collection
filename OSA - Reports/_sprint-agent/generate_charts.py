#!/usr/bin/env python3
"""
generate_charts.py — renders the Sprint Report's chart PNGs (matplotlib, headless) for embedding
into the .docx via ImageRun. docx-js can't draw live charts, so these are pre-rendered images.

Usage:
    python3 generate_charts.py <sprint_data.json> <trend_data.json|-> <output_dir> [--developers Name1,Name2]

--developers restricts workload_bar.png to a chosen subset of developers (comma-separated names,
case-insensitive match against activity[].name). Pass the same list to build_sprint_report_docx.js's
--developers flag so the chart and the table agree. Omitted, or no match, means everyone with
activity in the sprint.

Writes, into <output_dir>:
    byType_donut.png       Work items by type (current sprint)
    open_closed_donut.png  Open vs Closed (current sprint)
    workload_bar.png       Active/Resolved/Closed stacked bar, by developer (current sprint)
    trend_chart.png        Two-panel trend across all sprints in trend_data.json (skipped if "-")
"""
import json
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

NAVY = "#1E2761"
BLUE = "#3B82C4"
AMBER = "#C48A1E"
GREEN = "#1E6B3C"
RED = "#A32D2D"
GREY = "#8B949E"
PALETTE = ["#5B4FA3", "#3B82C4", "#C48A1E", "#C4507A", "#1E6B3C", "#A32D2D", "#5F5E5A"]

plt.rcParams.update({
    "font.size": 10,
    "axes.edgecolor": "#D0D7DE",
    "axes.labelcolor": "#333333",
    "text.color": "#222222",
    "xtick.color": "#333333",
    "ytick.color": "#333333",
})


def donut(ax, labels, values, title):
    colors = [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
    wedges, _ = ax.pie(values, colors=colors, startangle=90, wedgeprops={"width": 0.42, "edgecolor": "white", "linewidth": 2})
    ax.set_title(title, fontsize=12, fontweight="bold", color=NAVY, pad=10)
    total = sum(values) or 1
    legend_labels = [f"{l} ({v}, {round(v/total*100)}%)" for l, v in zip(labels, values)]
    ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False, fontsize=9)


def make_byType_donut(stats, out_path):
    by_type = stats["byType"]
    labels = list(by_type.keys())
    values = [by_type[k] for k in labels]
    fig, ax = plt.subplots(figsize=(5.4, 3.2), dpi=150)
    donut(ax, labels, values, "Work Items by Type")
    fig.tight_layout()
    fig.savefig(out_path, transparent=True)
    plt.close(fig)


def make_open_closed_donut(stats, out_path):
    labels = ["Open", "Closed"]
    values = [stats["open"], stats["closed"]]
    fig, ax = plt.subplots(figsize=(4.6, 3.2), dpi=150)
    colors = [BLUE, GREEN]
    wedges, _ = ax.pie(values, colors=colors, startangle=90, wedgeprops={"width": 0.42, "edgecolor": "white", "linewidth": 2})
    ax.set_title("Open vs Closed", fontsize=12, fontweight="bold", color=NAVY, pad=10)
    total = sum(values) or 1
    legend_labels = [f"{l} ({v}, {round(v/total*100)}%)" for l, v in zip(labels, values)]
    ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False, fontsize=9)
    fig.tight_layout()
    fig.savefig(out_path, transparent=True)
    plt.close(fig)


def make_workload_bar(activity, out_path):
    if not activity:
        return False
    names = [a["name"] for a in activity]
    active = [a["active"] for a in activity]
    resolved = [a["resolved"] for a in activity]
    closed = [a["closed"] for a in activity]

    fig, ax = plt.subplots(figsize=(9.2, 3.6), dpi=150)
    x = range(len(names))
    ax.bar(x, active, color=BLUE, label="Active")
    ax.bar(x, resolved, bottom=active, color=AMBER, label="Resolved")
    bottom2 = [a + r for a, r in zip(active, resolved)]
    ax.bar(x, closed, bottom=bottom2, color=GREEN, label="Closed")
    ax.set_xticks(list(x))
    ax.set_xticklabels(names, rotation=20, ha="right")
    ax.set_title("Workload by Developer (Active / Resolved / Closed)", fontsize=12, fontweight="bold", color=NAVY, pad=10)
    ax.legend(frameon=False, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.28))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(out_path, transparent=True)
    plt.close(fig)
    return True


def make_trend_chart(series, out_path):
    if not series or len(series) < 2:
        return False
    labels = [s["folder"].replace("Sprint-", "").replace("_", " – ") for s in series]
    completed = [s["totalCompleted"] for s in series]
    # "deployedRate" (Closed-only / total) replaced the old blended "completionRate" (Closed+Resolved).
    deployed_rate = [s.get("deployedRate", s.get("completionRate", 0)) for s in series]
    bug_ratio = [s["bugRatio"] for s in series]
    x = range(len(labels))

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(9.6, 6.4), dpi=150)

    # --- Panel 1: Velocity & deployed rate ---
    ax1.bar(x, completed, color=BLUE, alpha=0.85, label="Work items completed")
    ax1.set_ylabel("Items completed", color=BLUE)
    ax1.tick_params(axis="y", colors=BLUE)
    ax1b = ax1.twinx()
    ax1b.plot(x, deployed_rate, color=NAVY, marker="o", linewidth=2, label="Deployed rate (%)")
    ax1b.set_ylabel("Deployed rate (%)", color=NAVY)
    ax1b.set_ylim(0, 100)
    ax1b.tick_params(axis="y", colors=NAVY)
    ax1.set_title("Velocity & Deployed Rate by Sprint", fontsize=12, fontweight="bold", color=NAVY, pad=10)
    ax1.set_xticks(list(x))
    ax1.set_xticklabels(labels, rotation=15, ha="right", fontsize=8)
    lines1, labs1 = ax1.get_legend_handles_labels()
    lines1b, labs1b = ax1b.get_legend_handles_labels()
    ax1.legend(lines1 + lines1b, labs1 + labs1b, frameon=False, loc="upper left", fontsize=8)
    ax1.spines["top"].set_visible(False)

    # --- Panel 2: Quality (bug ratio can exceed 100%, so it gets its own free-scaling axis) ---
    ax2.plot(x, bug_ratio, color=RED, marker="o", linewidth=2, label="Bug ratio (%)")
    ax2.fill_between(x, bug_ratio, color=RED, alpha=0.08)
    ax2.set_ylabel("Bug ratio (%)", color=RED)
    ax2.tick_params(axis="y", colors=RED)
    ax2.set_title("Bug Ratio by Sprint", fontsize=12, fontweight="bold", color=NAVY, pad=10)
    ax2.set_xticks(list(x))
    ax2.set_xticklabels(labels, rotation=15, ha="right", fontsize=8)
    ax2.legend(frameon=False, loc="upper right", fontsize=8)
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_visible(False)

    fig.tight_layout()
    fig.savefig(out_path, transparent=True)
    plt.close(fig)
    return True


def main():
    if len(sys.argv) < 4:
        print("Usage: python3 generate_charts.py <sprint_data.json> <trend_data.json|-> <output_dir> [--developers Name1,Name2]", file=sys.stderr)
        sys.exit(1)
    sprint_data_path, trend_path, out_dir = sys.argv[1], sys.argv[2], sys.argv[3]
    os.makedirs(out_dir, exist_ok=True)

    # --developers restricts the workload-by-developer bar chart to a chosen subset (comma-separated
    # names, matched case-insensitively against activity[].name) — pass the SAME list used for
    # build_sprint_report_docx.js's --developers flag so the chart matches the table. Omitted, or if
    # none of the names match, falls back to everyone with activity in the sprint.
    developers_filter = None
    if "--developers" in sys.argv:
        idx = sys.argv.index("--developers")
        if idx + 1 < len(sys.argv):
            developers_filter = [n.strip() for n in sys.argv[idx + 1].split(",") if n.strip()]

    with open(sprint_data_path) as f:
        data = json.load(f)
    stats = data["stats"]
    activity = data.get("activity", [])
    if developers_filter:
        wanted_lower = {n.lower() for n in developers_filter}
        filtered = [a for a in activity if a["name"].lower() in wanted_lower]
        if filtered:
            activity = filtered
        else:
            print("Warning: none of --developers " + ",".join(developers_filter) + " matched a known developer — showing all developers instead.", file=sys.stderr)

    written = {}
    make_byType_donut(stats, os.path.join(out_dir, "byType_donut.png"))
    written["byType_donut"] = os.path.join(out_dir, "byType_donut.png")
    make_open_closed_donut(stats, os.path.join(out_dir, "open_closed_donut.png"))
    written["open_closed_donut"] = os.path.join(out_dir, "open_closed_donut.png")
    if make_workload_bar(activity, os.path.join(out_dir, "workload_bar.png")):
        written["workload_bar"] = os.path.join(out_dir, "workload_bar.png")

    if trend_path != "-":
        with open(trend_path) as f:
            trend = json.load(f)
        if make_trend_chart(trend["series"], os.path.join(out_dir, "trend_chart.png")):
            written["trend_chart"] = os.path.join(out_dir, "trend_chart.png")

    print(json.dumps(written))


if __name__ == "__main__":
    main()
