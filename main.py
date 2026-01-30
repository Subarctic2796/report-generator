from pathlib import Path
from collections import Counter
from datetime import datetime, date
from os.path import isfile
from sys import stderr, argv, exit
from typing import TypeAlias, Literal
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QPushButton, QMessageBox

import pyexcel as pyxl

# Types
"""(name, date): num [assigned|closed]"""
ExtractedInfo: TypeAlias = Counter[tuple[str, date]]
BothExtractedInfo: TypeAlias = tuple[ExtractedInfo, ExtractedInfo]
Row: TypeAlias = tuple[str, str | int, str | int]


def get_info(path: str, key: Literal["assigned", "aps_rec"]) -> ExtractedInfo | None:
    """assumes that `path` exists.
    `key` can only be the string 'assigned' or 'aps_rec'"""
    print(f"[INFO]: [get_info]: attempting to read '{path}'")
    records = pyxl.iget_records(file_name=path)
    cntr: ExtractedInfo = Counter()
    for ix, i in enumerate(records):
        if ix == 0:
            print(i)
        try:
            name: str = i["retriever"]
        except KeyError:
            name: str = i["ret"]
        except Exception as err:
            print(f"[ERROR]: [get_info]: '{path=}' {key=}: {err=}, {type(err)=}", file=stderr)
            return None

        date_: datetime | date = i[key].date() if isinstance(i[key], datetime) else i[key]
        cntr[(name, date_)] += 1
    pyxl.free_resources()
    print(f"[INFO]: [get_info]: successfully read {'assigned' if key == 'assigned' else 'closed'} file: {path}")
    return cntr


def get_assigned_and_closed(assigned_path: str, closed_path: str) -> BothExtractedInfo | None:
    """gets both the assigned and the closed info"""
    if not isfile(assigned_path):
        print(f"[ERROR]: [get_assigned_and_closed]: '{assigned_path}' does not exist", file=stderr)
        return None
    if not isfile(closed_path):
        print(f"[ERROR]: [get_assigned_and_closed]: '{closed_path}' does not exist", file=stderr)
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

class MainWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("report generator")
        self.resize(720, 480)
        self.setAcceptDrops(True)

        self.assign_path: str | None = None
        self.close_path: str | None = None

        self.get_assigned()
        self.get_closed()
        self.save_btn()

    def get_assigned(self):
        self.button = QPushButton("Choose assigned file", self)
        self.button.clicked.connect(self.openFileDialog_assigned)
        self.button.setGeometry(150, 150, 150, 30)

    def get_closed(self):
        self.button = QPushButton("Choose closed file", self)
        self.button.clicked.connect(self.openFileDialog_closed)
        self.button.setGeometry(150, 180, 150, 30)

    def save_btn(self):
        self.button = QPushButton("Create assigned to closed", self)
        self.button.clicked.connect(self.createClosed)
        self.button.setGeometry(150, 210, 150, 30)

    def createClosed(self):
        msg_box = QMessageBox()
        can_save = self.close_path is not None and self.assign_path is not None
        if can_save:
            save_path = str(Path.home()) + f"\\Desktop\\output_{date.today()}.xlsx"
            status = generate_assigned_to_closed(self.assign_path, self.close_path, save_path)
            if status:
                msg_box.setText("Successfully created output")
            else:
                msg_box.setText("Failed to create consolidated file, check logs")
        else:
            msg_box.setText("must choose a closed and assigned file")
        msg_box.exec()

    def openFileDialog_closed(self):
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("Choose a closed file")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setViewMode(QFileDialog.ViewMode.Detail)

        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            self.close_path = selected_files[0]
            print("Selected File:", selected_files[0])

    def openFileDialog_assigned(self):
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("Choose an assigned file")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setViewMode(QFileDialog.ViewMode.Detail)

        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            self.assign_path = selected_files[0]
            print("Selected File:", selected_files[0])


if __name__ == '__main__':
    app = QApplication(argv)
    ui = MainWidget()
    ui.show()
    exit(app.exec())
