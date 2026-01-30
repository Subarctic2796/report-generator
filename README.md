# IntelliBPO Paramed Assigned vs Closed Report Generator
This creates the assigned vs closed report for parameds at IntelliBPO.

## Usage
```shell
report-gen /path/to/assigned.xls /path/to/closed.xls /path/to/output.xlsx
```

## Building
Dependencies for building
- pyinstaller
```shell
git clone --depth=1 url/to/repo/report-generator.git
cd report-generator
# create and setup venv
pip install -r requirements.txt
pip install pyinstaller
```
Then either run
```shell
pyinstaller report-gen.spec
```
or
```shell
pyinstaller --clean -Fn report-gen main.py \
         --hiddenimport=pyexcel_io.writers \
         --hiddenimport=pyexcel_io.readers \
         --hiddenimport=pyexcel_xls        \
         --hiddenimport=pyexcel_xlsx 
```