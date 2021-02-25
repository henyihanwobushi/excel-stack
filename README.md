# Excel stack
Copy from input excel files and stack the sheets together.

### Input files
work book | work sheet
--|--
wb1 | 
| | ws1
| | ws2
wb2 | 
| | ws1
| | ws2

### Output  files
work book | work sheet
--|--
ws1 |
| | wb1
| | wb2
ws2 |
| | wb1
| | wb2

## Features
* Copy worksheet between workbooks
* Copy worksheet with format, cell merge
* Compile to exe for windows with [gooey](https://github.com/chriskiehl/Gooey)

## Attention
* The script assumes that the input files have same format
