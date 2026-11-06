# biomed-agent-kit

A collection of Claude Code skills for biomedical and healthcare ML research. Each skill provides reusable patterns, code snippets, and workflows that Claude can apply when working on your projects.

## Skills

| Skill | Purpose |
|-------|---------|
| [project-scaffold](.claude/skills/project-scaffold/) | Bootstrap new research repos with uv, CLAUDE.md, pre-commit, and standard layout |
| [data-preprocessing](.claude/skills/data-preprocessing/) | Load and preprocess DICOM, NIfTI, WSI, and tabular clinical data |
| [training](.claude/skills/training/) | Training loops, transfer learning, mixed precision, checkpointing |
| [experiment](.claude/skills/experiment/) | Experiment grids, GPU management, data splitting, hyperparameter tuning |
| [evaluation](.claude/skills/evaluation/) | Classification, segmentation, survival, calibration, and regression metrics |
| [reporting](.claude/skills/reporting/) | Publication-quality figures, ROC/PR curves, LaTeX tables, and markdown reports |
| [deployment](.claude/skills/deployment/) | Model export (ONNX), inference pipelines, sliding window, TTA, serving |
| [repo-integration](.claude/skills/repo-integration/) | Integrate GitHub repos, HuggingFace models, and MONAI into your project |
| [clinical-nlp](.claude/skills/clinical-nlp/) | Biomedical NLP with clinical notes, radiology reports, and PubMed text |
| [paper-research](.claude/skills/paper-research/) | Research papers across PubMed, Semantic Scholar, arXiv, focusing on top venues |
| [meeting-deck](skills/meeting-deck/) | Build a group meeting .pptx: new results since the last meeting, charts from stored files, every number recomputed, file checks |

## Setup

Install skills into your project by copying the `.claude/skills/` directory, or install individual skills as needed.

### Requirements

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) CLI
- Python 3.10+
- [uv](https://docs.astral.sh/uv/) for package management

## Usage

Skills trigger automatically based on what you ask Claude.

### project-scaffold

```
"Create a new project for my tumor classification study"
"Set up a new repo for lung CT analysis with uv and wandb"
"Bootstrap a research project called cell-segmentation"
"Initialize a project with proper CLAUDE.md and test structure"
```

### experiment

```
"Set up a 5-fold cross-validation grid over 3 encoders"
"Run experiments across 4 GPUs with OOM protection"
"Design patient-level splits so no patient appears in both train and test"
"Create a leave-one-site-out evaluation for our multi-center dataset"
"Set up an experiment grid over feature sets and classifiers with wandb logging"
"Add resumable experiment tracking so I can restart interrupted runs"
```

### evaluation

```
"Compute AUC-ROC with 95% confidence intervals across folds"
"Compute Dice scores and HD95 for each organ class"
"Run a Wilcoxon test comparing model A vs model B across folds"
"Calculate C-index with bootstrap confidence intervals for our survival model"
"Evaluate the segmentation model with surface Dice at 2mm tolerance"
"Compare all models with Bonferroni-corrected p-values"
```

### reporting

```
"Generate a forest plot comparing all models"
"Create a heatmap of AUC scores: feature sets vs encoders"
"Write a markdown report summarizing the experiment results"
"Make publication-quality figures at 300 DPI for the paper"
"Generate a radar chart comparing top 3 configurations across all metrics"
```

### repo-integration

```
"Integrate the pretrained model from this paper's GitHub repo"
"Add this ViT encoder from GitHub as a feature extractor in our pipeline"
"Vendor the preprocessing code from REPO and adapt it to our data format"
"Wrap this external segmentation model so it accepts our DICOM inputs"
"This repo requires torch 1.x but we use torch 2 — help me integrate it"
"Set up the pretrained weights from this paper's Google Drive link"
```

### data-preprocessing

```
"Load DICOM files from this directory and convert to numpy arrays"
"Set up CT windowing for soft tissue and normalize to [0,1]"
"Build a PyTorch Dataset class for our whole-slide images with tile extraction"
"Apply stain normalization to our histopathology patches"
"Create a cached HDF5-backed dataset for this large imaging study"
```

### training

```
"Write a training loop with early stopping and cosine LR scheduling"
"Fine-tune this pretrained ResNet on our pathology dataset"
"Set up mixed precision training with gradient accumulation"
"Add checkpoint saving that keeps the best model by validation AUC"
"Implement progressive unfreezing for transfer learning"
```

### deployment

```
"Export our trained model to ONNX format"
"Build a sliding window inference pipeline for our 3D CT volumes"
"Add test-time augmentation to our prediction pipeline"
"Create a FastAPI endpoint for model inference"
"Package our model in a Docker container for deployment"
```

### clinical-nlp

```
"Extract medical entities from these radiology reports using scispaCy"
"Fine-tune PubMedBERT for classifying clinical notes"
"Set up a text classification pipeline for ICD code prediction"
"Extract embeddings from clinical notes using ClinicalBERT"
"Build a few-shot classifier for pathology report categorization"
```

### paper-research

```
"Find state-of-the-art methods for cell segmentation in histopathology"
"Search for papers on survival prediction from MICCAI and Nature Medicine"
"Create a literature comparison table for federated learning in medical imaging"
"Find which papers have public code for chest X-ray classification"
"Summarize recent advances in foundation models for pathology"
```

### meeting-deck

```
"Make my deck for Friday's group meeting"
"Put the new external test results on the deck"
"Refresh the status slides and check every number on this deck"
"组会 PPT 帮我做一下"
```

To install it for all your projects:

```bash
cp -R skills/meeting-deck ~/.claude/skills/
```

```bash
pip install -r ~/.claude/skills/meeting-deck/requirements.txt
```

On first use, Claude asks for your name, projects, where results live and who will be in the room, then saves them to `settings.md`. The style is a default: tell Claude what you want changed and it records that in `settings.md`, which stays yours when you copy a newer version over it.

## License

MIT
