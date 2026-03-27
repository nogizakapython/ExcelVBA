VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function input_check(ByVal check_value As String) As Long
    '変数の定義
    Dim ws As Worksheet
    Dim sheet_name As String
    Dim input_value As String
    Dim input_value_count As Long
    Dim error_msg As String
    'シート名を代入する
    sheet_name = "ボタンフォーム"
    'ws変数にボタンフォームシートを設定する。
    Set ws = ThisWorkbook.Worksheets(sheet_name)
    
    '処理で値を変える
    If check_value = "name_check" Then
        '作業者を取得する
        input_value = ws.Range("C4").Value
        error_msg = name_check_error_msg
    ElseIf check_value = "health_check" Then
        '作業者を取得する
        input_value = ws.Range("C7").Value
        error_msg = health_check_error_msg
    End If
    '作業者名の文字列を定義する
    input_value_count = Len(input_value)
    
    
    
    '作業者名が選択されていない場合、エラーメッセージを表示させて、リターンコード1で返す
    If input_value_count = 0 Then
        MsgBox error_msg, vbCritical
        input_check = 1
        Exit Function
    End If
    
    input_check = 0
End Function


