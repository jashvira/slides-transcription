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
transcribe-slides path/to/presentation.pptx output_dir --model base
```

Add `--gui` to launch a GUI instead of the CLI.

The tool extracts audio from each slide and writes a text file per slide inside the `output_dir` folder (`slide1.txt`, `slide2.txt`, ...).

## License

This project is released under the MIT License.
