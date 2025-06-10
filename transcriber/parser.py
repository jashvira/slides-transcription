import argparse
import os

try:
    from gooey import GooeyParser
except Exception:  # pragma: no cover - Gooey is optional
    GooeyParser = None  # type: ignore


def build_parser(gui: bool = False) -> argparse.ArgumentParser:
    """Return an argument parser for the CLI and GUI."""
    if gui and GooeyParser:
        parser = GooeyParser(description="Transcribe audio from PowerPoint slides")
        io = parser.add_argument_group("Input/Output", gooey_options={"columns": 1})
        io.add_argument(
            "pptx",
            help="Path to the PowerPoint file",
            widget="FileChooser",
            gooey_options={"wildcard": "PowerPoint files (*.pptx)|*.pptx"},
        )
        io.add_argument("output", help="Directory to save transcriptions", widget="DirChooser")

        trans = parser.add_argument_group("Transcription Settings", gooey_options={"columns": 1})
        trans.add_argument(
            "--model",
            default="whisper-1",
            choices=["whisper-1"],
            help="OpenAI STT model to use",
            widget="Dropdown",
        )
        trans.add_argument(
            "--language",
            default="en",
            help="Language spoken in the audio (e.g., 'en' for English)",
        )
        trans.add_argument(
            "--task",
            choices=["transcribe", "translate"],
            default="transcribe",
            help="Transcription task",
            widget="Dropdown",
        )
        trans.add_argument("--prefix", default="slide", help="Prefix for output files")

        api = parser.add_argument_group("API Settings", gooey_options={"columns": 1})
        api.add_argument(
            "--api-key",
            default=os.environ.get("OPENAI_API_KEY", ""),
            help="OpenAI API key (or set OPENAI_API_KEY environment variable)",
            widget="PasswordField",
        )
    else:
        parser = argparse.ArgumentParser(
            description="Transcribe slide audio using the OpenAI Speech to Text API"
        )
        parser.add_argument("pptx", help="Path to PPTX file")
        parser.add_argument("output", help="Output directory for transcripts")
        parser.add_argument(
            "--model",
            default="whisper-1",
            choices=["whisper-1"],
            help="OpenAI STT model to use (currently only whisper-1 is available)",
        )
        parser.add_argument(
            "--language",
            default="en",
            help="Language spoken in the audio (e.g., 'en' for English)",
        )
        parser.add_argument(
            "--task",
            choices=["transcribe", "translate"],
            default="transcribe",
            help="Transcription task",
        )
        parser.add_argument("--prefix", default="slide", help="Prefix for output files")
        parser.add_argument(
            "--api-key",
            help="OpenAI API key (or set OPENAI_API_KEY environment variable)",
        )
        parser.add_argument(
            "--gui",
            action="store_true",
            help="Launch a GUI instead of the CLI",
        )
    return parser
