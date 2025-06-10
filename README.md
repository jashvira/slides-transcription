# Slides Transcription

A simple tool that extracts audio from Microsoft PowerPoint presentations and transcribes each slide using the [OpenAI Speech to Text API](https://platform.openai.com/docs/guides/speech-to-text). A minimal GUI can be launched via [Gooey](https://github.com/chriskiehl/Gooey).

## Installation

Install the package using `pip` (preferably in a virtual environment):

```bash
pip install -e .
```

The command will install the required dependencies such as `python-pptx`, `openai` and `gooey`.

You will also need an OpenAI API key which can be provided via the `OPENAI_API_KEY` environment variable or the `--api-key` command line option.

## Usage

Run the command line interface:

```bash
transcribe-slides path/to/presentation.pptx output_dir \
    --language en --task transcribe --prefix slide --api-key YOUR_API_KEY
```

Add `--gui` to launch a GUI instead of the CLI.

The available options are:

* `--model` – OpenAI STT model name (currently only whisper-1 is available)
* `--language` – language spoken in the audio
* `--task` – transcription task (`transcribe` or `translate`)
* `--prefix` – prefix used when naming transcript files
* `--api-key` – OpenAI API key (or set `OPENAI_API_KEY`)

### Models and Pricing

Currently, OpenAI only provides the Whisper-1 model for speech-to-text tasks. Here are the details:

#### Whisper-1
- Cost: $0.006 per minute of audio
- Features:
  - Automatic language detection
  - Supports multiple languages
  - Handles different accents and dialects
  - Good at handling background noise
  - Supports both transcription and translation
- Limitations:
  - Maximum file size: 25MB
  - Maximum audio length: 25MB per file
  - No real-time streaming (batch processing only)
  - No custom model training available

For more details about the model and its capabilities, visit the [OpenAI Speech-to-Text documentation](https://platform.openai.com/docs/guides/speech-to-text).

### Output Format

The tool extracts audio from each slide and writes a text file per slide inside the `output_dir` folder (`slide1.txt`, `slide2.txt`, etc.). If a slide has multiple audio tracks, they are combined into a single file with clear track labeling:

```
Track 1 of 3:
[transcription text]

Track 2 of 3:
[transcription text]

Track 3 of 3:
[transcription text]
```

### Building the Windows installer

The repository contains basic scripts to build a standalone executable and NSIS installer. You need [PyInstaller](https://pyinstaller.org/) and [NSIS](https://nsis.sourceforge.io/) available on your system.

Run the build script from the project root:

```bash
bash build/build.sh
```

This produces `dist/transcribe-slides.exe` and `dist/SlidesTranscriberSetup.exe`.

## License

This project is released under the MIT License.
