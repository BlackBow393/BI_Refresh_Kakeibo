import os
import time
from pywinauto import Application, keyboard

def main():
    exe = 'PBIDesktop.exe'

    # Power BIのファイルパスを直接指定
    workbook = r"C:\Users\t9374\OneDrive\デスクトップ\個人家計簿BI\個人家計簿_ver.1.0.0.pbix"  # ここにPower BIファイルのパスを指定

    # ファイルを開く
    os.system('start "" "{0}"'.format(workbook))
    
    try:
        # アプリケーションに接続
        app = Application(backend='uia').connect(path=exe)
        
        # ウィンドウが表示されるまで待つ
        win = app.window(title_re=r'^個人家計簿_ver\.1\.0\.0')
        win.wait('visible', timeout=30)  # 30秒間、ウィンドウが表示されるのを待機

        # ウィンドウにフォーカスを当てる
        win.set_focus()

        # ホーム＞更新をクリック
        win.ホーム.wait('visible')
        win.ホーム.click_input()
        
        win.更新.wait('visible')
        win.更新.click_input()

        # キャンセルボタンが非表示になるのを待つ（最大15秒）
        win.キャンセル.wait_not('visible', timeout=15000)

        # 保存
        keyboard.send_keys("^s")
        time.sleep(15)

        # 更新完了メッセージ
        print("更新完了")
        
    except Exception as e:
        print("更新できませんでした")
        print(e)
    finally:
        app.kill()

if __name__ == '__main__':
    main()
