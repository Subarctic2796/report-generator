from sys import stderr, argv, exit
from collections import Counter
from datetime import datetime, date
from os.path import isfile
from typing import TypeAlias

import pyexcel as pyxl

# Types
"""(name, date): num [assigned|closed]"""
AssignedClosedInfo: TypeAlias = Counter[tuple[str, date]]
BothAssignedClosedInfo: TypeAlias = tuple[AssignedClosedInfo, AssignedClosedInfo]


def get_assigned(path: str) -> AssignedClosedInfo:
    """assumes that `path` exists. You are responsible for making sure that `path` exists"""
    records = pyxl.iget_records(file_name=path)
    cntr: Counter[tuple[str, date]] = Counter()
    for i in records:
        date_: datetime | date = i["assigned"]
        if isinstance(date_, datetime):
            cntr[(i["retriever"], date_.date())] += 1
        else:
            cntr[(i["retriever"], date_)] += 1
    pyxl.free_resources()
    print(f"successfully read assigned file: {path}")
    return cntr


def get_closed(path: str) -> AssignedClosedInfo:
    """assumes that `path` exists.You are responsible for making sure that `path` exists"""
    records = pyxl.iget_records(file_name=path)
    cntr: Counter[tuple[str, date]] = Counter()
    for i in records:
        date_: datetime | date = i["aps_rec"]
        if isinstance(date_, datetime):
            cntr[(i["retriever"], date_.date())] += 1
        else:
            cntr[(i["retriever"], date_)] += 1
    pyxl.free_resources()
    print(f"successfully read closed file: {path}")
    return cntr


def get_assigned_and_closed(assigned_path: str, closed_path: str) -> BothAssignedClosedInfo | None:
    """gets both the assigned and the closed info"""
    if not isfile(assigned_path):
        print(f"[ERROR]: [display_assigned_or_closed]: '{assigned_path}' does not exist", file=stderr)
        return None
    if not isfile(closed_path):
        print(f"[ERROR]: [display_assigned_or_closed]: '{closed_path}' does not exist", file=stderr)
        return None
    return get_assigned(assigned_path), get_closed(closed_path)


def generate_assigned_to_closed(assigned_path: str, closed_path: str, output_path: str) -> bool:
    """generates the assigned to closed info. returns False on error and True on success"""
    both_optional = get_assigned_and_closed(assigned_path, closed_path)
    if both_optional is None:
        return False
    assigned, closed = both_optional

    # flatten
    data: list[tuple[str, str | int, str | int]] = [("name", "assigned", "closed")]
    for k, v in assigned.items():
        data.append((k[0], v, closed[k]))
    pyxl.save_as(array=data, sheet_name="parameds consolidated", dest_file_name=output_path)
    print(f"successfully created {output_path}")
    return True


if __name__ == '__main__':
    if len(argv) != 4:
        print("usage: report-gen /path/to/assigned.xls, /path/to/closed.xls /path/to/output.xlsx")
        exit(1)

    assigned_path, closed_path, output_path = argv[1:]
    if not output_path.endswith(".xlsx"):
        print("[ERROR]: [main]: output path must end with xlsx", file=stderr)
        exit(1)

    if not generate_assigned_to_closed(assigned_path, closed_path, output_path):
        print("[ERROR]: [main]: unable to create consolidated output", file=stderr)
        exit(1)
