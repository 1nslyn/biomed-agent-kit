"""Example chart. Run joined to the style file: cat chart_style.py plot_deck.py | python3 - figs/deck"""

import json
import sys

results = json.load(open("data/results.json"))
fig, ax = plt.subplots(figsize=(CHART_W, 3.2))
dumbbell(ax, results["labels"], results["before"], results["after"], (66, 78))
ax.set_xlabel("Test AUC (%), 80 samples")
ax.legend(handles=[dot(BEFORE, "Before"), dot(AFTER, "After")], loc="lower right")
fig.tight_layout()
save(fig, f"{sys.argv[1]}/before_after.png")
