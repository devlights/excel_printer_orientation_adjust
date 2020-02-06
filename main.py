#################################################################
# 指定されたフォルダ配下のExcelを開いていき特定の条件にマッチするシートの印刷方向を調整します.
#
# 実行には、以下のライブラリが必要です.
#   - win32com
#     - $ python -m pip install pywin32
#
# [参考にした情報]
#   - http://excel.style-mods.net/tips_vba/tips_vba_7_09.htm
#   - https://stackoverflow.com/a/37635373
#   - https://www.sejuku.net/blog/23647
#   - https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlpageorientation
#   - http://excel.style-mods.net/tips_vba/tips_vba_7_03.htm
#################################################################
import argparse

# 縦
xlPortrait = 1
# 横
xlLandscape = 2


# noinspection SpellCheckingInspection
def go(target_dir: str, pattern: str, orientation: str):
    import pathlib
    import win32com.client

    excel_dir = pathlib.Path(target_dir)
    if not excel_dir.exists():
        print(f'target directory not found [{target_dir}]')
        return

    if not orientation:
        print(f'orientation is invalid [{orientation}]')
        return
    else:
        if orientation == 'portrait':
            xl_orientation = xlPortrait
        elif orientation == 'landscape':
            xl_orientation = xlLandscape
        else:
            print(f'orientation is invalid [{orientation}]')
            return

    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True
    except Exception as err:
        print(err)
        return

    try:
        for f in excel_dir.glob('**/*.xlsx'):
            abs_path = str(f)
            try:
                wb = excel.Workbooks.Open(abs_path)
            except Exception as err:
                print(err)
                continue

            try:
                sheets_count = wb.Sheets.Count
                for sheet_index in range(0, sheets_count):
                    ws = wb.Worksheets(sheet_index + 1)
                    ws.Activate()
                    if not pattern:
                        ws.PageSetup.Orientation = xl_orientation
                    else:
                        if pattern in ws.Name:
                            ws.PageSetup.Orientation = xl_orientation
                if sheets_count >= 0:
                    ws = wb.Worksheets(1)
                    ws.Activate()
                wb.Save()
                wb.Saved = True
            finally:
                wb.Close()
    finally:
        excel.Quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        usage='python main.py -d /path/to/excel/dir -p シート名条件 -o [portrait|landscape]',
        description='Excelの特定シートの印刷方向を調整します.',
        add_help=True
    )

    parser.add_argument('-d', '--directory', help='対象ディレクトリ', required=True)
    parser.add_argument('-p', '--pattern', help='シート名の条件 (python の in 演算子で判定しています）指定しない場合は全シートが対象', default='')
    parser.add_argument('-o', '--orientation', help='印刷方向 (portrait(縦) or landscape(横))', default='portrait')

    args = parser.parse_args()

    go(args.directory, args.pattern, args.orientation)
