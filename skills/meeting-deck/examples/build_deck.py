"""Example deck with placeholder values. Run from a folder holding deck_style.py, chart_style.py and this example."""

import json

import deck_style as ds

results = json.load(open("data/results.json"))
prs = ds.new_deck()
ds.title_slide(prs, "Weekly Meeting November 13th", ["A. Presenter", "Nov 6 – 13, 2026"], "Example with placeholder values.")

s = ds.page(prs, "Project A: summary", "Model C, trained on cohort A, tested on the same 80 samples: 69.2% to **72.8%**.", "Example notes: what changed, and where the numbers come from.")
rows = [["Experiment", "Trained on", "Tested on", "Result"]]
for label, before, after in zip(results["labels"], results["before"], results["after"]):
    rows.append([f"{label}, new learning rate", "Cohort A, 320 samples", "Cohort A test, 80 samples", f"{before}% to **{after}%**"])
ds.table(s, rows, ds.M, ds.CONTENT_Y, [4.4, 2.7, 3.0, 2.2], strong=(3,))
ds.caption(s, ["Test AUC, mean of 5 seeds."], y=3.3)

s = ds.page(prs, "Project A: test AUC, before and after", "All three models rise or hold on the same 80 samples. The largest gain is Model C, 69.2% to **72.8%**.", "Example notes.")
ds.figure(s, "figs/deck/before_after.png")
ds.caption(s, ["Test AUC, mean of 5 seeds. Before and after differ only in learning rate."], y=5.1)

s = ds.page(prs, "Project A: training plan", "Sample 1,000 items per step, then score all items at validation and test.", "Example notes. Native shapes stay editable in PowerPoint.")
ds.band(s, ds.M, 1.8, ds.BODY_W, 1.5)
ds.node(s, 0.8, 2.0, 3, 1.1, "Sample", "1,000 per step")
ds.node(s, 5.1, 2.0, 3, 1.1, "Train", "one model per fold")
ds.node(s, 9.4, 2.0, 3, 1.1, "Score", "all items", dashed=True)
ds.arrow(s, 3.8, 2.55, 5.1, 2.55)
ds.arrow(s, 8.1, 2.55, 9.4, 2.55)

ds.todos(prs, {"Project A": ["Finish the 5-seed rerun of Model B.", "Compare the new learning rate on cohort B."]}, "Example notes.")

prs.save("example.pptx")
