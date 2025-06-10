import os
import sys
from argparse import Namespace

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from transcriber.transcribe_slides import run

TEST_PPTX = os.path.join(os.path.dirname(__file__), 'assets', 'multimedia.pptx')


def test_run_no_audio(capsys, tmp_path):
    args = Namespace(
        pptx=TEST_PPTX,
        output=str(tmp_path),
        model='whisper-1',
        language='en',
        task='transcribe',
        prefix='slide',
        api_key='dummy',
    )
    run(args)
    captured = capsys.readouterr()
    assert 'No audio found in PPTX' in captured.out


def test_run_requires_api_key(capsys, tmp_path):
    args = Namespace(
        pptx=TEST_PPTX,
        output=str(tmp_path),
        model='whisper-1',
        language='en',
        task='transcribe',
        prefix='slide',
        api_key=None,
    )
    run(args)
    captured = capsys.readouterr()
    assert 'OpenAI API key required' in captured.out
