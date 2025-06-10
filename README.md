# Slides Transcription

A simple tool that extracts audio from Microsoft PowerPoint presentations and transcribes each slide using [OpenAI Whisper](https://github.com/openai/whisper). A minimal GUI can be launched via [Gooey](https://github.com/chriskiehl/Gooey).

## Installation

Install the package using `pip` (preferably in a virtual environment):

```bash
pip install -e .
```

The command will install the required dependencies such as `python-pptx`, `openai-whisper` and `gooey`.

## Usage

Run the command line interface:

```bash
transcribe-slides path/to/presentation.pptx output_dir \
    --model base --language en --task transcribe --prefix slide
```

Add `--gui` to launch a GUI instead of the CLI.

The available options are:

* `--model` – Whisper model name or path.
* `--language` – language spoken in the audio.
* `--task` – Whisper task (`transcribe` or `translate`).
* `--prefix` – prefix used when naming transcript files.

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
