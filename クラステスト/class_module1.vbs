' 外部から呼び出し可能な関数。引数としてRange（セル範囲）を受け取り、結果をLong型（数値）で返す
Public Function data_clear(ByVal clear_range As Range) As Long
    
    ' --- 1. エラーチェック ---
    ' もし渡されたセル範囲が空（セットされていない状態）だった場合の処理
    If clear_range Is Nothing Then
        ' エラーメッセージを表示（vbCriticalは「×」アイコンの警告）
        MsgBox "有効なRangeを引数に指定してください", vbCritical
        
        ' 失敗を表す戻り値として「1」をセットして関数を終了する
        data_clear = 1
        Exit Function
    End If
        
    ' --- 2. クリア処理の実行 ---
    ' 指定されたセル範囲の「値と数式」のみを消去する（背景色や罫線などの書式は残る）
    clear_range.ClearContents
    
    ' --- 3. 正常終了の戻り値設定 ---
    ' 成功を表す戻り値として「0」をセットする
    data_clear = 0
    
End Function