import PySimpleGUI as sg
import os
from pathlib import Path
import openpyxl
from typing import NamedTuple
import pandas as pd
from pandas import DataFrame


class ShadowList(NamedTuple):
    checkboxes: list
    source_items: list
    target_items: list
    ok_marks: list
    source_files: list
    target_files: list


class Defaults:
    ROUTES_SHEET_NAME = '0.Свод'
    ECONOMY_SHEET_NAME = 'Экономика'
    ROUTES_AGGREGATE_RANGE = 'A4:AK9'
    ECONOMY_TARGET_CELL = 'B46'
    SKIP_ROWS_IN_TOTAL = 2
    TARGET_ROWS_OFFSET = 22


class Errors:
    PFILE_NOT_FOUND: str = 'Файл проекта не найден, используем значения по умолчанию'
    ROUTES_DIR_MISSING: str = 'Каталог параметров маршрутов отсутствует'
    ECONOMY_DIR_MISSING: str = 'Каталог расчётов экономики отсутствует'


class Success:
    DIR_PRESENT: str = 'Присутствует'
    FILE_PRESENT: str = 'Присутствует'


COL1_WIDTH = 30
COL2_WIDTH = 30
PATH_ADD_WIDTH = 30

ROUTES_DIR_ID = '-ROUTES-DIR-'
ECONOMY_DIR_ID = '-ECONOMY-DIR-'
ROUTES_INPUT_KEY = '-ROUTES-PREFIX-'
ECONOMY_INPUT_KEY = '-ECONOMY-PREFIX-'
CHECKBOX_KEY = 'CHECK-'
ROUTES_COMBO_KEY = 'ROUTES-COMBO-'
ECONOMY_COMBO_KEY = 'ECONOMY-COMBO-'
OK_TEXT_KEY = 'OK-'
TRANSFER_BUTTON_KEY = '-TRANSFER-'
MARK_ALL_BUTTON_KEY = '-MARK-ALL-'
EXIT_BUTTON_KEY = '-EXIT-'

NO_MATCHING_FILE = 'Нет сопоставления'

ADD_ID = '-ADD-'
FRAME_ID = '-FRAME-'

ROUTES_DIR = 'routes'
ECONOMY_DIR = 'economy'
PFILE = 'project.json'

ROUTES_PREFIX = 'МаршБуд'
ECONOMY_PREFIX = 'ЭкПоказатели'

BASE_DIR = Path(__file__).resolve(strict=True).parent
ROUTES_FULL_PATH = os.path.join(BASE_DIR, ROUTES_DIR)
ECONOMY_FULL_PATH = os.path.join(BASE_DIR, ECONOMY_DIR)

WINOW_FONT = ('Times', 14, 'normal')

"""
def copy_range(from_sheet, from_range: tuple[str, str], to_sheet, to_range: str):

    tcell = to_sheet[to_range]
    for rindex, row in enumerate(from_sheet[from_range]):
        for cindex, cell in enumerate(row):
            tcell.offset(row=rindex, column=cindex).value = cell.value
    return True

"""


def copy_range(dataframe: DataFrame, to_sheet, to_range: str, row_offset=0, header=''):
    tcell = to_sheet[to_range]
    rindex = 0
    if header != '':
        tcell.offset(row=rindex+row_offset,
                     column=0).value = header
        rindex += 1
    for row in dataframe.itertuples():
        for cindex, value in enumerate(row):
            tcell.offset(row=rindex+row_offset,
                         column=cindex).value = value
        rindex += 1
    return True


def transfer_data(from_filename, to_filename,
                  from_sheet_name=Defaults.ROUTES_SHEET_NAME,
                  to_sheet_name=Defaults.ECONOMY_SHEET_NAME,
                  from_range=Defaults.ROUTES_AGGREGATE_RANGE,
                  to_range=Defaults.ECONOMY_TARGET_CELL,
                  row_offset=Defaults.TARGET_ROWS_OFFSET):
    # from_workbook = openpyxl.load_workbook(
    #    from_filename, read_only=True, data_only=True)
    to_workbook = openpyxl.load_workbook(to_filename)
    # from_sheet = from_workbook[from_sheet_name]
    to_sheet = to_workbook[to_sheet_name]
    df = pd.read_excel(from_filename, from_sheet_name,
                       skiprows=Defaults.SKIP_ROWS_IN_TOTAL, index_col=0)
    print(df)
    print(f'Перенос данных из файла {from_filename} в файл {to_filename}')
    if copy_range(df, to_sheet, to_range):
        to_workbook.save(to_filename)
        print('Успешно.')
        return True
    else:
        print('Неудача!')
        return False


def check_projfile(file=Path(os.path.join(BASE_DIR, PFILE))):
    result = Success.FILE_PRESENT if file.exists(
    ) and file.is_file else Errors.PFILE_NOT_FOUND
    return result


sg.theme('Material2')   # Add a touch of color
# All the stuff inside your window.


def l_header():
    rpath = Path(ROUTES_FULL_PATH)
    epath = Path(ECONOMY_FULL_PATH)
    header = [
        # [sg.Titlebar('Красивый заголовок окна')],
        [sg.Text('Текущий каталог: ', size=(COL1_WIDTH, 1)), sg.Text(
            BASE_DIR, size=(COL2_WIDTH+PATH_ADD_WIDTH, 2))],
        [sg.Text('Файл проекта: ', size=(COL1_WIDTH, 1)), sg.Text(
            Errors.PFILE_NOT_FOUND, size=(COL2_WIDTH, 2))],
        [sg.Text('Каталог маршрутов: ', size=(COL1_WIDTH, 1)),
         sg.Text(
            key=ROUTES_DIR_ID, text='/'+ROUTES_DIR, size=(COL2_WIDTH, 1)),
         sg.Text(Errors.ROUTES_DIR_MISSING if not rpath.exists(
         ) or not rpath.is_dir() else Success.DIR_PRESENT)],
        [sg.Text('Каталог экономики: ', size=(COL1_WIDTH, 1)),
         sg.Text(key=ECONOMY_DIR_ID,
                 text='/'+ECONOMY_DIR, size=(COL2_WIDTH, 1)),
         sg.Text(Errors.ECONOMY_DIR_MISSING if not epath.exists(
         ) or not epath.is_dir() else Success.DIR_PRESENT)],
        [sg.Text('Префикс файлов маршрутов: ', size=(COL1_WIDTH, 1)), sg.Text(
            key=ROUTES_INPUT_KEY, text=ROUTES_PREFIX, size=(COL2_WIDTH, 1))],
        [sg.Text('Префикс файлов экономики: ', size=(COL1_WIDTH, 1)), sg.Text(
            key=ECONOMY_INPUT_KEY, text=ECONOMY_PREFIX, size=(COL2_WIDTH, 1))],
        [sg.HSeparator()],]
    return header


def l_footer():
    footer = [[sg.HSeparator()],
              [sg.Button('Перенести', disabled=True, disabled_button_color='grey', key=TRANSFER_BUTTON_KEY),
               sg.Button('Выход', button_color='red', key=EXIT_BUTTON_KEY),
               ]]
    return footer


def read_files_list(rpath: Path, epath: Path):
    rlist = [Path(file).name.replace(ROUTES_PREFIX, '')
             for file in sorted(rpath.glob('*.xlsx')) if not Path(file).name.startswith('~')]
    elist = [Path(file).name.replace(ECONOMY_PREFIX, '')
             for file in sorted(epath.glob('*.xlsx')) if not Path(file).name.startswith('~')]
    return (rlist, elist)


def make_files_panel_shadow_list(lists: tuple):
    sources = [*lists[0]]
    targets = [*lists[1]]
    checkboxes = []
    source_items = []
    target_items = []
    ok_marks = []
    for file in sources:
        source_items.append(file)
        if file in targets:
            target_items.append(file)
            targets.remove(file)
            checkboxes.append(True)
            ok_marks.append(True)
        else:
            target_items.append(NO_MATCHING_FILE)
            checkboxes.append(False)
            ok_marks.append(False)
    for file in targets:
        source_items.append(NO_MATCHING_FILE)
        target_items.append(file)
        checkboxes.append(False)
        ok_marks.append(False)
    return ShadowList(checkboxes, source_items, target_items, ok_marks,
                      lists[0]+[NO_MATCHING_FILE], lists[1]+[NO_MATCHING_FILE])


def files_panel(shadow_list: ShadowList):
    checkbox_col = [[sg.Text('')]]
    sources_col = [[sg.Text('Файлы маршрутов')]]
    targets_col = [[sg.Text('Файлы экономики')]]
    ok_col = [[sg.Text('')]]
    for index, checkbox in enumerate(shadow_list.checkboxes):
        checkbox_col.append([sg.Checkbox(
            text='Перенести', key=(CHECKBOX_KEY, index), default=checkbox, disabled=not checkbox, enable_events=True)])
        sources_col.append([sg.Combo(
            values=shadow_list.source_files, default_value=shadow_list.source_items[index],
            key=(ROUTES_COMBO_KEY, index), enable_events=True, readonly=True)])
        targets_col.append([sg.Combo(
            values=shadow_list.target_files, default_value=shadow_list.target_items[index],
            key=(ECONOMY_COMBO_KEY, index), enable_events=True, readonly=True)])
        ok_col.append([sg.Text('OK', text_color='green', font=(WINOW_FONT[0], WINOW_FONT[1], 'bold'), key=(
            OK_TEXT_KEY, index)) if checkbox else sg.Text('X', text_color='red', key=(OK_TEXT_KEY, index))])
    result = [
        [sg.Column(checkbox_col),
         sg.Column(sources_col),
         sg.Column(targets_col),
         sg.Column(ok_col)]
    ]
    return result


def find_duplicates(values_list):
    seen = set()
    res = {}
    for index, value in enumerate(values_list):
        if value in seen:
            res[value] = res[value]+1 if value in res else 2
        else:
            seen.add(value)
            # res[value] = [index]
    if res:
        del res[NO_MATCHING_FILE]
    return res


def make_window():
    shadow_list = make_files_panel_shadow_list(
        read_files_list(Path(ROUTES_FULL_PATH), Path(ECONOMY_FULL_PATH)))
    layout = l_header() + [[sg.Frame('Перенос параметров сети:',
                                     files_panel(shadow_list),
                                     key=FRAME_ID)]] + l_footer()
    window = sg.Window('Переносим данные', layout, resizable=True,
                       use_ttk_buttons=False, font=WINOW_FONT,
                       # use_custom_titlebar=True,
                       finalize=True)
    return (shadow_list, window)


def main():
    # Create the Window
    shadow_list, window = make_window()
    window[TRANSFER_BUTTON_KEY].update(
        disabled=not (True in shadow_list.checkboxes))
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        # if user closes window or clicks cancel
        if event in [sg.WIN_CLOSED, EXIT_BUTTON_KEY]:
            break
        if isinstance(event, tuple):
            index = event[1]
            not_matching = not (values[
                (ROUTES_COMBO_KEY, index)] != NO_MATCHING_FILE and values[
                    (ECONOMY_COMBO_KEY, index)] != NO_MATCHING_FILE)
            window[(CHECKBOX_KEY, index)].update(disabled=not_matching)
            if not_matching:
                window[(CHECKBOX_KEY, index)].update(value=False)
            # window[TRANSFER_BUTTON_KEY].update(disabled=not_matching)
            window[(OK_TEXT_KEY, index)].update(
                text_color='red' if not_matching else 'green',
                value='X' if not_matching else 'OK',
                font=(WINOW_FONT[0], WINOW_FONT[1], 'normal') if not_matching else (WINOW_FONT[0], WINOW_FONT[1], 'bold'))
            shadow_list.checkboxes[index] = values[(
                CHECKBOX_KEY, index)]
            shadow_list.source_items[index] = values[(
                ROUTES_COMBO_KEY, index)]
            shadow_list.target_items[index] = values[(
                ECONOMY_COMBO_KEY, index)]
            shadow_list.ok_marks[index] = values[(
                ROUTES_COMBO_KEY, index)]
            window[TRANSFER_BUTTON_KEY].update(
                disabled=not (True in shadow_list.checkboxes))
            if duplicates := find_duplicates([item[1] if shadow_list.source_items[item[0]] != NO_MATCHING_FILE
                                              else NO_MATCHING_FILE for item in enumerate(shadow_list.target_items)]):
                err_layout = [[sg.Text('Внимание!')]]
                for duplicate in duplicates:
                    err_layout.append([sg.Text(
                        f'Файл {duplicate} установлен в качестве цели для переноса файлов {duplicates[duplicate]} раза!')])
                err_layout.append(
                    [sg.Text('При каждом переносе данных предыдущие данные в файле перезаписываются!')])
                err_layout.append([sg.OK()])
                sg.Window('Внимание!', err_layout,
                          font=WINOW_FONT).read(close=True)
        if event == TRANSFER_BUTTON_KEY:
            for index, checkbox in enumerate(shadow_list.checkboxes):
                if checkbox:
                    try:
                        transfer_data(os.path.join(ROUTES_FULL_PATH, ROUTES_PREFIX+shadow_list.source_items[index]), os.path.join(
                            ECONOMY_FULL_PATH, ECONOMY_PREFIX+shadow_list.target_items[index]))
                    except Exception as err:
                        sg.popup(
                            f'При копировании данных из файла {ROUTES_PREFIX+shadow_list.source_items[index]} в {ECONOMY_PREFIX+shadow_list.target_items[index]} возникла ошибка {err}')
            sg.popup('Данные скопированы успешно')
    window.close()


if __name__ == '__main__':
    main()
