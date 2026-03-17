import xlwings as xw             # 1. xlwingsをインポート
wb = xw.Book('msgtest1.xlsm') # 2. ブックを開く
macro = wb.macro('VBA_TEST')    # 3. マクロを取得
macro()                          # 4. マクロを実行