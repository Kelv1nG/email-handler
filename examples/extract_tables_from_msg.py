"""Extract selected HTML tables from a saved Outlook MSG file."""

from argparse import ArgumentParser
from pathlib import Path

from extractors import extract_tables_from_msg
from schemas import TableSelector

from ._display import print_tables


def main() -> None:
    parser = ArgumentParser(
        description="Extract tables from a saved MSG without Outlook."
    )
    parser.add_argument("msg_path", type=Path)
    parser.add_argument("--column", action="append", default=[])
    position = parser.add_mutually_exclusive_group()
    position.add_argument("--source-index", type=int)
    position.add_argument("--occurrence", type=int)
    args = parser.parse_args()

    selector = TableSelector(
        source_index=args.source_index,
        required_columns=args.column,
        occurrence=args.occurrence,
    )
    print_tables(extract_tables_from_msg(args.msg_path, selector))


if __name__ == "__main__":
    main()
