# excel2csv
Convert the inaccessible and proprietary office spreadsheets to something useful. 

## Usage
Place the repository in a globally visible location (ie within `$PATH`)

``` python -m excel2csv /path/myawkwardfile.xlsx -o /outputpath/OUTfolder ```


## args 
```
usage: excel2csv [-h] [-o [OUT]] fin

Convert Excel files to CSV files.

positional arguments:
  fin                   Input xlsx file

optional arguments:
  -h, --help            show this help message and exit
  -o [OUT], --out [OUT]
                        Output Directory

```