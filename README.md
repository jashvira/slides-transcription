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

* `--model` – OpenAI STT model name.
* `--language` – language spoken in the audio.
* `--task` – transcription task (`transcribe` or `translate`).
* `--prefix` – prefix used when naming transcript files.
* `--api-key` – OpenAI API key (or set `OPENAI_API_KEY`).

The tool extracts audio from each slide and writes a text file per slide inside the `output_dir` folder (`slide1.txt`, `slide2.txt`, ...).

### Building the Windows installer

The repository contains basic scripts to build a standalone executable and NSIS installer. You need [PyInstaller](https://pyinstaller.org/) and [NSIS](https://nsis.sourceforge.io/) available on your system.

Run the build script from the project root:

```bash
bash build/build.sh
```

This produces `dist/transcribe-slides.exe` and `dist/SlidesTranscriberSetup.exe`.

## License

This project is released under the MIT License.
