import time
import jinja2
import argparse
import collections
import itertools
import xlrd


BOARD_TEMPLATE_FILE_NAME = "board_template.json.jinja2"
WEKAN_TIME_FORMAT = "%Y-%m-%dT%H:%M:%S.000Z'"


class Card:
    _counter = itertools.count()

    def __init__(self, title, description):
        self._id = next(self._counter)
        self._title = title
        self._description = description

    @property
    def title(self):
        return self._title

    @property
    def description(self):
        return self._description

    @property
    def id(self):
        return self._id


Board = collections.namedtuple('Board', ('title', 'date'))


def extract_cards_from_excel(sheet_path, sheet_index, title_col, description_col, first_data_row, last_data_row) -> list:
    workbook = xlrd.open_workbook(sheet_path)
    sheet = workbook.sheet_by_index(sheet_index)

    if last_data_row is None:
        last_data_row = sheet.nrows

    cards = [
        Card(
            title=sheet.cell_value(row, title_col), description=sheet.cell_value(row, description_col)
        )
        for row in range(first_data_row, last_data_row)
    ]

    return cards


def generate_board(board_name: str, cards: list) -> str:
    board_creation_date =\
        time.strftime(WEKAN_TIME_FORMAT, time.localtime())

    template_data = open(BOARD_TEMPLATE_FILE_NAME).read()

    board = \
        Board(board_name, board_creation_date)

    json_data = jinja2.Template(template_data).render(
        board=board,
        cards=cards,
    )

    return json_data


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "sheet_path", help="path to excel sheet", type=str
    )

    parser.add_argument(
        "board_title", help="wekan board title", type=str
    )

    parser.add_argument(
        "--card_title_col", help="col index for card title", type=int, default=1
    )

    parser.add_argument(
        "--card_description_col", help="col index for card description", type=int, default=2
    )

    parser.add_argument(
        "--sheet_index", help="index of sheet inside the workbook", type=int, default=0
    )

    parser.add_argument(
        "--first_data_row", help="first row that contains task data", type=int, default=1
    )

    parser.add_argument(
        "--last_data_row", help="last row that contains task data", type=int
    )

    return parser.parse_args()


if __name__ == "__main__":
    args = parse_arguments()

    extracted_cards = extract_cards_from_excel(
        args.sheet_path,
        args.sheet_index,
        args.card_title_col,
        args.card_description_col,
        args.first_data_row,
        args.last_data_row,
    )

    board_json_data = generate_board(args.board_title, extracted_cards)

    print(board_json_data)
