# xcat

Expanded cat tool especially for Excel

# How to use

``` bash
xcat 0.1.0
Toshiya. K. <kawasakitoshiya@gmail.com>
Expanded cat tool especially for Excel

USAGE:
    xcat [OPTIONS] <file>

FLAGS:
    -h, --help       Prints help information
    -V, --version    Prints version information

OPTIONS:
    -d, --delimiter <DELIMITER>    Delimiter for csv file

ARGS:
    <file>...    Input files to use

```

# Description
`xcat` is tool to show excel and csv file simultaneously on command line interface.

``` bash
$ xcat data.xlsx data.csv
1,excel_row_1
2,excel_row_1
1,csv_row_2
2,csv_row_2
```

*Change delimiter*
``` bash
$ xcat -d"|" data.xlsx data.csv
1|excel_row_1
2|excel_row_1
1|csv_row_2
2|csv_row_2
```
