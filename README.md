# .XLSX to Wekan

Convert an xlsx file to a json file that can be imported as a wekan board.

## Usage

```
usage: xlsx_to_wekan.py [-h] [--card_title_col CARD_TITLE_COL]
                         [--card_description_col CARD_DESCRIPTION_COL]
                         [--sheet_index SHEET_INDEX]
                         [--first_data_row FIRST_DATA_ROW]
                         [--last_data_row LAST_DATA_ROW]
                         sheet_path board_title

positional arguments:
  sheet_path            path to excel sheet
  board_title           wekan board title

optional arguments:
  -h, --help            show this help message and exit
  --card_title_col CARD_TITLE_COL
                        col index for card title
  --card_description_col CARD_DESCRIPTION_COL
                        col index for card description
  --sheet_index SHEET_INDEX
                        index of sheet inside the workbook
  --first_data_row FIRST_DATA_ROW
                        first row that contains task data
  --last_data_row LAST_DATA_ROW
                        last row that contains task data
```

## Example

```bash
python3 xlsx_to_wekan.py example.xlsx "example board" --card_title_col=1 --card_description_col=2 --sheet_index=0
```
