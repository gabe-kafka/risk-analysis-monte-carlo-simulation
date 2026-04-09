#!/usr/bin/env python3
"""Combine all PNG outputs into a single composite image."""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.image as mpimg

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs"


def main() -> None:
    parser = argparse.ArgumentParser(description="Combine output PNGs into one image")
    parser.add_argument("--output", type=Path, default=OUT / "summary.png")
    args = parser.parse_args()

    panels = [
        OUT / "gantt.png",
        OUT / "monte-carlo-histogram.png",
        OUT / "monte-carlo-scurve.png",
        OUT / "monte-carlo-tornado.png",
    ]

    for p in panels:
        if not p.exists():
            print(f"Missing: {p.name} — run 'make' first")
            return

    fig, axes = plt.subplots(2, 2, figsize=(28, 20), facecolor="#FFFFFF")

    for ax, path in zip(axes.flat, panels):
        img = mpimg.imread(str(path))
        ax.imshow(img)
        ax.set_axis_off()

    fig.subplots_adjust(wspace=0.02, hspace=0.02)
    fig.savefig(args.output, dpi=200, facecolor="#FFFFFF", bbox_inches="tight", pad_inches=0.3)
    plt.close(fig)
    print(f"Combined output saved: {args.output}")


if __name__ == "__main__":
    main()
