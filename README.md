# IntelliBPO Paramed Assigned vs Closed Report Generator
This creates the assigned vs closed report for parameds at IntelliBPO.

## Usage
```shell
report-gen /path/to/assigned.xls /path/to/closed.xls /path/to/output.xlsx
```

## Building
Dependencies for building
- pyinstaller
```bash
git clone --depth=1 url/to/repo/report-generator.git
cd report-generator
# create and setup venv
pip install -r requirements.txt
pyinstaller -F --clean -n report-gen main.py
```