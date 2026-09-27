# Case Study Pipeline

Turns a practitioner's raw clinical notes into publishable narrative case studies, and extracts
structured metadata from them at the same time - in one pass, against a **locally hosted LLM**.

Built to process a backlog of several hundred `.docx` case files that were only useful as prose,
and were never going to be read as data.

## What it does

For each `.docx` of case notes, the model returns a single JSON object containing both the
rewritten narrative and the extracted fields:

```json
{
  "story_title":   "Chronic Fatigue",
  "main_symptom":  "Extreme fatigue and chronic pain",
  "essences_used": ["Amethyst", "Bluebell", "Brown Kelp"],
  "story_body":    "### Chronic Fatigue\n\n*This story is based on..."
}
```

The narrative is written to a formatted `.docx`; the metadata is appended to `story_log.csv`.
That CSV turns a pile of prose into a queryable dataset - which is what makes the corpus usable
downstream (it became the evaluation set for a recommendation system).

## Why local inference

Clinical notes are exactly the kind of material you don't hand to a third-party API. Running
against [Ollama](https://ollama.com) on `localhost` means the source documents never leave the
machine.

## Privacy

The pipeline is built around the assumption that its input is sensitive:

- The prompt **forbids using the client's name** in the title or story
- Every generated story carries an anonymisation notice
- Narratives are written in the third person with identifying details removed
- This repository ships **no real data** - `.docx` files, `story_log.csv` and the output
  directories are all gitignored, and the included sample is fictional

## Running it

```bash
pip install -r requirements.txt
ollama pull deepseek-v3.1:671b-cloud     # or set OLLAMA_MODEL to a local model
python story_generator.py
```

With no configuration it reads `sample_input/` and writes to `output/`. Point it at real
directories with environment variables:

```bash
CASE_STUDY_INPUT_DIR=/path/to/notes \
CASE_STUDY_OUTPUT_DIR=/path/to/stories \
CASE_STUDY_BASE_DIR=/path/to/project \
python story_generator.py
```

Already-processed files are skipped, so the batch is resumable - useful when a run over hundreds
of documents fails partway.

## Notes on the implementation

- **Structured output by contract** - the prompt specifies an exact JSON schema and the response
  is parsed and validated before anything is written. A malformed response fails that document
  rather than corrupting the log.
- **Filename sanitisation** - titles come from the model, so they're slugified before touching
  the filesystem.
- **Formatting is applied after generation** - the model returns Markdown-ish text, which is
  converted to real `.docx` styling (headings, bold, italics) rather than written literally.

## Stack

Python · Ollama · python-docx
