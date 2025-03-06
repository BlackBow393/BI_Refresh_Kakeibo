import os
import sys
import time

from pywinauto import Desktop, Application, keyboard

def main(workbook):
    exe = 'PBIDesktop.exe'

    # ファイルを開く
    os.system('start "" "{0}"'.format(workbook))
    app = Application(backend='uia').connect(path=exe)
    time.sleep(60)
    try:
        # Windowを指定
        win = app.window(title_re = '.*Power BI Desktop')
        win.set_focus()
        # ホーム＞更新をクリック
        win.ホーム.wait("visible")
        win.ホーム.click_input()
        win.更新.wait("visible")
        win.更新.click_input()
        win.キャンセル.wait_not("visible",timeout=6000)
        # 保存
        keyboard.send_keys("^s")
        time.sleep(120)
    except Exception as e:
        print(e)
    finally:
        app.kill()


if __name__ == '__main__':
    try:
        file_path = sys.argv[1]
    except (IndexError):
        print('ファイルを指定してください。')
        sys.exit()

    main(file_path)
