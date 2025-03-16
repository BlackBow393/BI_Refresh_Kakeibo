import os
import time
from pywinauto import Application, keyboard

def main():
    exe = 'PBIDesktop.exe'

    # Power BIのファイルパスを直接指定
    workbook = r"C:\Users\t9374\OneDrive\デスクトップ\個人家計簿BI\個人家計簿_ver.1.0.0.pbix"

    # ファイルを開く
    os.system('start "" "{0}"'.format(workbook))
    
    try:
        # アプリケーションに接続
        app = Application(backend='uia').connect(path=exe)
        
        # メインウィンドウが表示されるまで待つ
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
        try:
            # キャンセルボタンが別ウィンドウに表示されている場合、他のウィンドウを探す
            cancel_win = app.window(title_re='最新の情報に更新')  # 例: キャンセルウィンドウを探す
            cancel_win.wait('visible')
            print("更新画面を発見")
            
            
        except Exception as e:
            print("キャンセルボタンが表示されませんでした", e)

        print("更新完了")

        # 「保存」オプションが表示されるまで待つ
        win.保存.wait('visible')  # 「保存」オプションが表示されるまで待機

        # 「保存」をクリック
        win.保存.click_input()
        time.sleep(15)  # 15秒間待機して保存が完了するのを待つ

        # 保存完了メッセージ
        print("保存完了")
        
    except Exception as e:
        print("保存できませんでした")
        print(e)
    finally:
        app.kill()

if __name__ == '__main__':
    main()
