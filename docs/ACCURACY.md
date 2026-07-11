# Regression corpus results

Constructed-fixture verification on the [public corpus](../corpus/), regenerated
by `python scripts/make_accuracy_table.py`. Every number is **publicly
reproducible** (the decks and manifests ship in this repository and CI fails on
drift), which is deliberately weaker language than third-party verified: no
independent party has validated these rules yet.

Read this table for what it is: proof that the gates fire on deliberately seeded
defects and stay silent on clean decks. It is NOT an accuracy benchmark, for two
reasons stated plainly. First, the same author wrote the rules and the fixtures,
so this is regression verification, not independent ground truth. Second, the
per-gate samples are tiny; the exact one-sided 95% lower bound below shows how
little a small perfect score proves (1/1 -> 5%, 3/3 -> 37%).

Corpus: 20 decks, 4 generator families, 6 clean negatives, 24 slides, all
synthetic (no field decks yet). Observed false positives across the whole
corpus: 0 (0.00 per 10 slides on this corpus).

| Gate | TP | FP | FN | Deck-level TN | Recall (95% LB) |
|:----:|---:|---:|---:|---:|:---|
| E1 | 4 | 0 | 0 | 16 | 4/4 (>= 47%) |
| E2 | 2 | 0 | 0 | 18 | 2/2 (>= 22%) |
| E3 | 1 | 0 | 0 | 19 | 1/1 (>= 5%) |
| E4 | 2 | 0 | 0 | 18 | 2/2 (>= 22%) |
| W1 | 1 | 0 | 0 | 19 | 1/1 (>= 5%) |
| W6 | 1 | 0 | 0 | 19 | 1/1 (>= 5%) |
| W8 | 1 | 0 | 0 | 19 | 1/1 (>= 5%) |
| W15 | 1 | 0 | 0 | 19 | 1/1 (>= 5%) |
| W16 | 1 | 0 | 0 | 19 | 1/1 (>= 5%) |

Gates absent from the table have no positive fixture in the corpus yet (their
behavior is covered by the unit/property suites, not by corpus ground truth).
W18 is excluded by design: it is an abstention signal, not a finding.

What would upgrade this into an accuracy record: unseen decks from generators
the author did not write fixtures for (Google Slides / Canva / LibreOffice
exports, real corporate templates), independent labeling, and contributed
false-positive reproductions. That corpus growth is the highest-value
contribution this project takes; see [corpus/README.md](../corpus/README.md).
