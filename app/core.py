"""Program-specific classes and functions."""

from __future__ import annotations

from app import card, constants, sheet
from datetime import datetime
from PIL import Image

import pathlib
import subprocess
import tkinter as tk


class App(tk.Tk):
    """Main Application."""

    def __init__(self, title: str, size: tuple):
        super().__init__()

        self.title(title)
        self.geometry(f"{size[0]}x{size[1]}")


def add_cards_to_sheets(
    sheet_template: sheet.SheetTemplate, cards: list[card.Card]
) -> list[sheet.Sheet]:
    for c in cards:
        c.build_image()

    initial_x = 0
    initial_y = (sheet_template.size[1] // 3 - cards[0].front_image.size[1] // 2) // 4
    offset_x = sheet_template.size[0] // 2
    offset_y = sheet_template.size[1] // 3

    back_x = offset_x + 10
    front_x = initial_x
    y = initial_y

    current_sheet = sheet.Sheet(sheet_template)
    sheets = []

    for index, current_card in enumerate(cards):
        if index != 0 and not index % 6:
            sheets.append(current_sheet)
            current_sheet = sheet.Sheet(sheet_template)
            back_x = offset_x + 10
            front_x = initial_x
            y = initial_y

        current_card.build_image()

        if index == 0 or not index % 2:
            front_x = initial_x
            back_x = offset_x + 10

            current_sheet.back.paste(current_card.back_image, (back_x, y))
            current_sheet.front.paste(current_card.front_image, (front_x, y))

        else:
            front_x += offset_x
            back_x -= offset_x - 10

            current_sheet.back.paste(current_card.back_image, (back_x, y))
            current_sheet.front.paste(current_card.front_image, (front_x, y))

            y += offset_y

    if not current_sheet in sheets:
        sheets.append(current_sheet)

    # page margins
    margin_params = (
        "RGB",
        tuple(int(i * 1.1) for i in sheet_template.parameters[1]),
        "white",
    )

    # add margins to sheets
    for s in sheets:
        front_margins = Image.new(*margin_params)
        back_margins = Image.new(*margin_params)

        front_margins.paste(
            s.front,
            (
                (front_margins.size[0] - sheet_template.size[0]) // 2,
                (front_margins.size[1] - sheet_template.size[1]) // 2,
            ),
        )
        s.front = front_margins

        back_margins.paste(
            s.back,
            (
                (back_margins.size[0] - sheet_template.size[0]) // 2,
                (back_margins.size[1] - sheet_template.size[1]) // 2,
            ),
        )
        s.back = back_margins

    return sheets


def create_cards(card_data: card.CardData) -> list[card.Card]:
    """Takes information from a CardData object and returns a list of Cards."""
    card_list = []

    for card_index in range(0, card_data.card_count):
        current_card = card.Card(
            assembly=card_data.assembly,
            bkt_hrs=card_data.bkt_hrs,
            bkt_qty=card_data.remainder_parts
            if card_data.remainder_parts and card_index + 1 == card_data.card_count
            else card_data.bkt_qty,
            card_count=card_data.card_count,
            card_num=card_index + 1,
            exp_vel=card_data.exp_vel,
            job_num=card_data.job_num,
            job_qty=card_data.job_qty,
            ops=card_data.ops,
            part_name=card_data.part_name,
            part_num=card_data.part_num,
            pro_date=card_data.pro_date,
        )

        card_list.append(current_card)
        card_index += 1

    return card_list


def create_card_data(
    quantities: dict[str, str], details: dict[str, str], ops: str
) -> card.CardData:
    """Create a new card.CardData object from form values."""
    qtys: dict = translate_dict_keys(quantities)
    dtls: dict = translate_dict_keys(details)
    ops: list = [i.strip().upper() for i in ops.split(",")]

    return card.CardData(**qtys, **dtls, ops=ops)


def export_cards(quantities: dict[str, str], details: dict[str, str], ops: str):
    """Run full card-creation process."""
    sheet_template = sheet.SheetTemplate(constants.SHEET_SIZE)

    card_data: card.CardData = create_card_data(quantities, details, ops)
    card_list: list[card.Card] = create_cards(card_data)
    sheets = add_cards_to_sheets(sheet_template, card_list)

    job_num: str = card_data.job_num

    output_path = generate_page_images(job_num, sheets)
    output_file = generate_pdf(job_num, output_path)

    open_in_explorer(output_file)


def generate_page_images(
    output_dir_name: str, sheets: list[sheet.Sheet]
) -> pathlib.Path:
    """Create image files from list of Sheets and return the output path."""
    export_time = datetime.now().strftime("%m.%d.%Y-%H.%M")
    path = pathlib.Path(f"./output/{output_dir_name}-{export_time}")
    path.mkdir(parents=True, exist_ok=True)

    for page_num, s in enumerate(sheets):
        s.front.save(path / f"{page_num}-1.jpg")
        s.back.save(path / f"{page_num}-2.jpg")

    return path


def generate_pdf(filename: str, image_directory: pathlib.Path):
    """Generate the final PDF for printing and return the output file path."""
    image_paths = image_directory.glob("*.jpg")

    images = [Image.open(path) for path in image_paths]
    first_page = images.pop(0)

    output_filename = f"{image_directory / filename}.pdf"

    first_page.save(
        output_filename, "PDF", resolution=100.0, save_all=True, append_images=images
    )

    return output_filename


def get_form_translation(label: str) -> str:
    """Get variable name from constants.FORM_TRANSLATIONS constant, given label text."""
    return constants.FORM_TRANSLATIONS[label]


def open_in_explorer(path: pathlib.Path):
    subprocess.Popen(f'explorer "{path}"')


def translate_dict_keys(d: dict[str, str]) -> dict:
    """Return a dict with translated keys per constants.FORM_TRANSLATIONS constant."""
    return {get_form_translation(lbl): v for lbl, v in d.items()}
