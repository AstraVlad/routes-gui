import PySimpleGUI as sg
import os
from pathlib import Path
import openpyxl


class Defaults:
    ROUTES_SHEET_NAME = '0.Свод'
    ECONOMY_SHEET_NAME = 'Экономика'
    ROUTES_AGGREGATE_RANGE = 'A4:AK9'
    ECONOMY_TARGET_CELL = 'B46'


class Errors:
    PFILE_NOT_FOUND: str = 'Файл проекта не найден, используем значения по умолчанию'
    ROUTES_DIR_MISSING: str = 'Каталог параметров маршрутов отсутствует'
    ECONOMY_DIR_MISSING: str = 'Каталог расчётов экономики отсутствует'


class Success:
    DIR_PRESENT: str = 'Присутствует'
    FILE_PRESENT: str = 'Присутствует'


COL1_WIDTH = 30
COL2_WIDTH = 30

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


def copy_range(from_sheet, from_range: tuple[str, str], to_sheet, to_range: str):

    tcell = to_sheet[to_range]
    for rindex, row in enumerate(from_sheet[from_range]):
        for cindex, cell in enumerate(row):
            tcell.offset(row=rindex, column=cindex).value = cell.value
    return True


def transfer_data(from_filename, to_filename,
                  from_sheet_name=Defaults.ROUTES_SHEET_NAME,
                  to_sheet_name=Defaults.ECONOMY_SHEET_NAME,
                  from_range=Defaults.ROUTES_AGGREGATE_RANGE,
                  to_range=Defaults.ECONOMY_TARGET_CELL):
    from_workbook = openpyxl.load_workbook(
        from_filename, read_only=True, data_only=True)
    to_workbook = openpyxl.load_workbook(to_filename)
    from_sheet = from_workbook[from_sheet_name]
    to_sheet = to_workbook[to_sheet_name]
    if copy_range(from_sheet, from_range, to_sheet, to_range):
        to_workbook.save(to_filename)
        return True
    else:
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
            BASE_DIR, size=(COL2_WIDTH, 2))],
        [sg.Text('Файл проекта: ', size=(COL1_WIDTH, 1)), sg.Text(
            Errors.PFILE_NOT_FOUND, size=(COL2_WIDTH, 2))],
        [sg.Text('Каталог маршрутов: ', size=(COL1_WIDTH, 1)),
         sg.InputText(
            key=ROUTES_DIR_ID, default_text=ROUTES_DIR, size=(COL2_WIDTH, 1)),
         sg.Text(Errors.ROUTES_DIR_MISSING if not rpath.exists(
         ) or not rpath.is_dir() else Success.DIR_PRESENT)],
        [sg.Text('Каталог экономики: ', size=(COL1_WIDTH, 1)),
         sg.InputText(key=ECONOMY_DIR_ID,
                      default_text=ECONOMY_DIR, size=(COL2_WIDTH, 1)),
         sg.Text(Errors.ECONOMY_DIR_MISSING if not epath.exists(
         ) or not epath.is_dir() else Success.DIR_PRESENT)],
        [sg.Text('Префикс файлов маршрутов: ', size=(COL1_WIDTH, 1)), sg.InputText(
            key=ROUTES_INPUT_KEY, default_text=ROUTES_PREFIX, size=(COL2_WIDTH, 1))],
        [sg.Text('Префикс файлов экономики: ', size=(COL1_WIDTH, 1)), sg.InputText(
            key=ECONOMY_INPUT_KEY, default_text=ECONOMY_PREFIX, size=(COL2_WIDTH, 1))],
        [sg.HSeparator()],]
    return header


def l_footer():
    footer = [[sg.HSeparator()],
              [sg.Button('Перенести', disabled=True, disabled_button_color='grey', key=TRANSFER_BUTTON_KEY),
               sg.Button('Выход', button_color='red', key=EXIT_BUTTON_KEY),
               ]]
    return footer


def read_files_list(rpath: Path, epath: Path):
    rlist: list = rpath.glob('*.xlsx')
    elist: list = epath.glob('*.xlsx')
    return (rlist, elist)


def files_panel(rlist, elist):
    files1 = [Path(file).name.replace(ROUTES_PREFIX, '')
              for file in sorted(rlist) if not Path(file).name.startswith('~')]
    files2 = [Path(file).name.replace(ECONOMY_PREFIX, '')
              for file in sorted(elist) if not Path(file).name.startswith('~')]
    max_lenght = max(len(files1), len(files2))
    files1.append(NO_MATCHING_FILE)
    files2.append(NO_MATCHING_FILE)
    result = []

    for i in range(max_lenght):
        row = ['', '', '', sg.T('X', text_color='red', key=(OK_TEXT_KEY, i))]
        if i < len(files1):
            row[1] = sg.Combo(
                values=files1, default_value=files1[i], key=(ROUTES_COMBO_KEY, i), enable_events=True, readonly=True)
            if files1[i] in files2:
                row[0] = sg.Checkbox(
                    text='Перенести', key=(CHECKBOX_KEY, i), default=True)
                row[2] = sg.Combo(
                    values=files2, default_value=files2[files2.index(files1[i])], key=(ECONOMY_COMBO_KEY, i), enable_events=True, readonly=True)
                row[3] = sg.Text('OK', text_color='green', font=(
                    WINOW_FONT[0], WINOW_FONT[1], 'bold'), key=(OK_TEXT_KEY, i))
            else:
                row[0] = sg.Checkbox(
                    text='Перенести', key=(CHECKBOX_KEY, i), disabled=True)
                row[2] = sg.Combo(
                    values=files2, default_value=files2[-1], key=(ECONOMY_COMBO_KEY, i), readonly=True, enable_events=True)
        else:
            row[0] = sg.Checkbox(
                text='Перенести', key=(CHECKBOX_KEY, i), disabled=True)
            row[1] = sg.Combo(
                values=files1, default_value=files1[-1], key=(ROUTES_COMBO_KEY, i), readonly=True, enable_events=True)
            row[2] = sg.Combo(
                values=files2, default_value=files2[files2[i]], key=(ECONOMY_COMBO_KEY, i), readonly=True, enable_events=True)
        result.append(row)
    return result


def find_duplicates(values_list):
    seen = set()
    res = {}
    for index, value in enumerate(values_list):
        if value in seen:
            print(value, res[value])
            res[value] = res[value]+[index]
        else:
            seen.add(value)
            res[value] = [index]
            print(res)
    return res


def make_window():
    rlist, elist = read_files_list(Path(ROUTES_FULL_PATH), Path(
        ECONOMY_FULL_PATH))
    layout = l_header() + [[sg.Frame('Перенос параметров сети:',
                                     files_panel(rlist, elist),
                                     key=FRAME_ID)]] + l_footer()
    window = sg.Window('Переносим данные', layout, resizable=True,
                       use_ttk_buttons=False, font=WINOW_FONT,
                       # use_custom_titlebar=True,
                       finalize=True)
    return window


def main():
    # Create the Window
    window = make_window()

    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        # if user closes window or clicks cancel
        if event in [sg.WIN_CLOSED, EXIT_BUTTON_KEY]:
            break
        if event == ADD_ID:
            # window.extend_layout(window[FRAME_ID], [
            #                     [sg.T('New one'), sg.T('New two')]])
            print(values[(ROUTES_COMBO_KEY, 1)])
        if isinstance(event, tuple):
            not_matching = not (values[
                (ROUTES_COMBO_KEY, event[1])] != NO_MATCHING_FILE and values[
                    (ECONOMY_COMBO_KEY, event[1])] != NO_MATCHING_FILE)
            window[(CHECKBOX_KEY, event[1])].update(disabled=not_matching)
            # window[TRANSFER_BUTTON_KEY].update(disabled=not_matching)
            window[(OK_TEXT_KEY, event[1])].update(
                text_color='red' if not_matching else 'green',
                value='X' if not_matching else 'OK',
                font=(WINOW_FONT[0], WINOW_FONT[1], 'normal') if not_matching else (WINOW_FONT[0], WINOW_FONT[1], 'bold'))

    window.close()


if __name__ == '__main__':
    main()
