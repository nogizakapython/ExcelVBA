Sub Clear_click()
    ' --- 1. 変数の定義 ---
    Dim wb As Workbook             ' 対象のワークブック
    Dim ws1 As Worksheet           ' 「データ1」シート用変数
    Dim ws2 As Worksheet           ' 「データ2」シート用変数
    Dim sheet_name1 As String      ' シート名格納用（データ1）
    Dim sheet_name2 As String      ' シート名格納用（データ2）
    Dim clear_range As Range       ' クリア対象となるセルの範囲オブジェクト
    Dim cls1 As Class1             ' 独自のクラス（Class1）のインスタンス用変数
    
    Dim start_row_num As Long      ' データの開始行番号
    Dim last_row_num As Long       ' データの最終行番号
    
    Dim clear_data_check As Long   ' データクリア処理の結果を受け取る変数（0:成功 / 1:失敗）
    
    Dim not_clear_msg As String    ' クリア対象データがない場合のメッセージボックスのタイトル
    
    
    ' --- 2. 設定値の初期化 ---
    sheet_name1 = "データ1"
    sheet_name2 = "データ2"
    not_clear_msg = "データクリア済だよ～"
    
    
    ' --- 3. オブジェクトのセット ---
    Set wb = ThisWorkbook                ' マクロが記述されているこのブックをセット
    Set ws1 = wb.Worksheets(sheet_name1) ' "データ1"シートをセット
    Set ws2 = wb.Worksheets(sheet_name2) ' "データ2"シートをセット
    Set cls1 = New Class1                ' Class1（ソースコード②）を呼び出すためにインスタンス化してセット
    
    
    ' --- 4. 行番号の取得とデータ存在チェック ---
    start_row_num = 3 ' データの開始行を3行目に指定
    
    ' ws2（データ2シート）のB列（2列目）を下から上へ検索し、データが入っている最終行を取得
    last_row_num = ws2.Cells(Rows.Count, 2).End(xlUp).Row
    
    ' 取得した最終行が開始行(3行目)未満の場合、すでにクリア済み（またはデータ無し）と判断
    If last_row_num < start_row_num Then
        ' メッセージを表示してマクロ処理を終了する
        MsgBox "クリア対象のデータがありません", vbInformation, not_clear_msg
        Exit Sub
    End If
    
    
    ' --- 5. クリア処理の実行（クラスの呼び出し） ---
    ' クリア対象のセル範囲をセット（データ2シートの B列3行目 ～ C列の最終行）
    Set clear_range = ws2.Range("B" & start_row_num & ":C" & last_row_num)
    
    ' Class1の「data_clear」関数にセル範囲を渡して実行し、戻り値（成功=0, 失敗=1）を受け取る
    clear_data_check = cls1.data_clear(clear_range)
    
    
    ' --- 6. 結果の判定と通知 ---
    ' Class1から返ってきた戻り値が0なら成功、それ以外（1など）なら失敗としてメッセージを表示
    If clear_data_check = 0 Then
        MsgBox "データクリア成功"
    Else
        MsgBox "データクリア失敗"
    End If
    
End Sub