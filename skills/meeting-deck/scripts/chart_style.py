"""The chart style for lab meeting decks. Copy into the deck's report folder.

Draw every chart at its size on the slide, so its text keeps its point size:
    fig, ax = plt.subplots(figsize=(CHART_W, 4.5))
    save(fig, "figs/deck/internal.png")
and place it with deck_style.figure(slide, "figs/deck/internal.png").

Read values from stored files only. A plot script starts with no import of this file and runs joined to it,
which works the same locally and on a server:
    cat chart_style.py plot_deck.py | python3 - figs/deck
    cat chart_style.py plot_deck.py | ssh <server> "cd <results dir> && <python> - figs/deck"
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
import numpy as np  # noqa: E402

# One meaning per color across the deck. To change the font, change font.family here and FONT in deck_style.py.
BEFORE, AFTER, TRAIN, VAL = "#7f7e7a", "#2a78d6", "#1baf7a", "#6250d6"
INK, MUTED, LIGHT, GRID = "#0b0b0b", "#52514e", "#c9c8c4", "#ecebe7"
PLATE = dict(facecolor="white", edgecolor="none", pad=1.2)
CHART_W, DPI = 12.3, 200
plt.rcParams.update({"font.size": 13, "font.family": ["Times New Roman", "STIXGeneral"], "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": LIGHT, "axes.linewidth": 0.9, "xtick.color": MUTED, "ytick.color": MUTED, "axes.labelcolor": MUTED,
                     "xtick.major.size": 0, "ytick.major.size": 0, "axes.titlesize": 14, "axes.titleweight": "bold", "axes.titlecolor": INK, "legend.frameon": False})


def save(fig, path):
    fig.savefig(path, dpi=DPI, facecolor="white")
    plt.close(fig)


def dot(color, label):
    """A legend entry for a dot series."""
    return Line2D([], [], color=color, marker="o", ls="", ms=9, label=label)


def clean_rows(ax, n):
    ax.spines["left"].set_visible(False)
    ax.set_ylim(-0.7, n - 0.3)
    ax.grid(axis="x", color=GRID, lw=0.9)
    ax.set_axisbelow(True)


def dumbbell(ax, labels, before, after, xlim, ci=None):
    """One row per setting: before and after dots with both values printed. `ci` is one (low, high) per row for the after dot."""
    y = np.arange(len(labels))[::-1]
    for i, (b, a) in enumerate(zip(before, after)):
        if ci is not None:
            ax.plot(ci[i], [y[i] - 0.27] * 2, color=AFTER, lw=1.2, alpha=0.45, solid_capstyle="butt")
        ax.plot([b, a], [y[i]] * 2, color=LIGHT, lw=2.2, zorder=1)
        ax.scatter([b], [y[i]], color=BEFORE, s=75, zorder=3, edgecolor="white", linewidth=1.2)
        ax.scatter([a], [y[i]], color=AFTER, s=75, zorder=3, edgecolor="white", linewidth=1.2)
        left, right = (b, a) if b <= a else (a, b)
        ax.annotate(f"{left:.1f}", (left, y[i]), xytext=(-8, 0), textcoords="offset points", ha="right", va="center", fontsize=12, color=MUTED, bbox=PLATE, zorder=2)
        ax.annotate(f"{right:.1f}", (right, y[i]), xytext=(8, 0), textcoords="offset points", ha="left", va="center", fontsize=12, color=MUTED, bbox=PLATE, zorder=2)
    ax.set_yticks(y, labels, color=INK)
    ax.set_xlim(*xlim)
    clean_rows(ax, len(labels))
