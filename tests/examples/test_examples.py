from pathlib import Path
import subprocess
import sys


def test_example_modules_import_without_running_cli():
    __import__("examples.extract_tables_from_msg")
    __import__("examples.extract_live_tables")
    __import__("examples.save_and_extract_message")


def test_offline_example_runs_against_fixture_without_outlook():
    fixture = Path(__file__).parents[1] / "fixtures" / "table_message.msg"
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "examples.extract_tables_from_msg",
            str(fixture),
            "--column",
            "Account",
            "--column",
            "Amount",
            "--occurrence",
            "0",
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    assert "Account | Amount" in result.stdout
    assert "A-100 | 125.50" in result.stdout
