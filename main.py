from collections import Counter
from datetime import datetime, date
from os.path import isfile
from sys import stderr, argv, exit
from typing import TypeAlias, Literal

import pyexcel as pyxl

# Types
"""(name, date): num [assigned|closed]"""
ExtractedInfo: TypeAlias = Counter[tuple[str, date]]
BothExtractedInfo: TypeAlias = tuple[ExtractedInfo, ExtractedInfo]
Row: TypeAlias = tuple[str, str | int, str | int]


def get_info(path: str, key: Literal["assigned", "aps_rec"]) -> ExtractedInfo | None:
    """assumes that `path` exists.
    `key` can only be the string 'assigned' or 'aps_rec'"""
    records = pyxl.iget_records(file_name=path)
    cntr: ExtractedInfo = Counter()
    for i in records:
        date_: datetime | date = i[key]
        if isinstance(date_, datetime):
            cntr[(i["retriever"], date_.date())] += 1
        else:
            cntr[(i["retriever"], date_)] += 1
    pyxl.free_resources()
    print(f"successfully read {'assigned' if key == 'assigned' else 'closed'} file: {path}")
    return cntr


def get_assigned_and_closed(assigned_path: str, closed_path: str) -> BothExtractedInfo | None:
    """gets both the assigned and the closed info"""
    if not isfile(assigned_path):
        print(f"[ERROR]: [display_assigned_or_closed]: '{assigned_path}' does not exist", file=stderr)
        return None
    if not isfile(closed_path):
        print(f"[ERROR]: [display_assigned_or_closed]: '{closed_path}' does not exist", file=stderr)
        return None

    assigned = get_info(assigned_path, "assigned")
    if assigned is None:
        return None

    closed = get_info(closed_path, "aps_rec")
    if closed is None:
        return None
    return assigned, closed


def generate_assigned_to_closed(assigned_path: str, closed_path: str, output_path: str) -> bool:
    """generates the assigned to closed info. returns False on error and True on success"""
    both_optional = get_assigned_and_closed(assigned_path, closed_path)
    if both_optional is None:
        return False
    assigned, closed = both_optional

    # flatten
    data: list[Row] = [("name", "assigned", "closed")] + [(k[0], v, closed[k]) for k, v in assigned.items()]
    pyxl.save_as(array=data, sheet_name="parameds consolidated", dest_file_name=output_path)
    print(f"successfully created {output_path}")
    return True


if __name__ == '__main__':
    if len(argv) != 4:
        print("usage: report-gen /path/to/assigned.xls, /path/to/closed.xls /path/to/output.xlsx")
        exit(1)

    assigned, closed, output = argv[1:]
    if not output.endswith(".xlsx"):
        print("[ERROR]: [main]: output path must end with xlsx", file=stderr)
        exit(1)

    if not generate_assigned_to_closed(assigned, closed, output):
        print("[ERROR]: [main]: unable to create consolidated output", file=stderr)
        exit(1)
