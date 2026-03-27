Attribute VB_Name = "Module1"
Option Explicit
'グローバル変数の定義
Public Const End_msg As String = "登録が完了しました！"
Public Const name_check_error_msg As String = "報告するメンバーの名前を選択してください"
Public Const health_check_error_msg As String = "本日の体調評価を選択してください"
Public Const output_check_error_msg As String = "データシートへの登録に失敗しました"


'メイン処理
'2026/3/27 新規作成　takao.hattori
Sub main()
    Dim name_check_code As Long
    Dim health_check_code As Long
    Dim output_check_code As Long
    
    name_check_code = name_check()
    
       
    If name_check_code > 0 Then
        Exit Sub
    End If
    
    health_check_code = health_check()
    
       
    If health_check_code > 0 Then
        Exit Sub
    End If
    
    output_check_code = output_data()
    
    If output_check_code > 0 Then
        Exit Sub
    End If
    
    MsgBox End_msg
    
    
    
End Sub

'報告者の名前選択チェック処理
'新規作成 2026/3/27 takao.hattori
Function name_check() As Long
    '変数の定義
    Dim cls1 As Class1
    Set cls1 = New Class1
    
    name_check = cls1.input_check("name_check")
    
    
End Function

'報告者の体調選択チェック処理
'新規作成 2026/3/27 takao.hattori
Function health_check() As Long
    '変数の定義
    Dim cls1 As Class1
    Set cls1 = New Class1
    
    health_check = cls1.input_check("health_check")
    
End Function
'データの入力処理
'新規作成  2026/3/27 takao.hattori
Function output_data() As Long
    Dim base_ws As Worksheet
    Dim target_ws As Worksheet
    Dim base_sheet_name As String
    Dim target_sheet_name As String
    
    
    Dim input_ymd As Date
    Dim member_name As String
    Dim health_status As String
    Dim detail As String
    
    Dim last_row_num As Long
    Dim output_row_num As Long
    
    On Error GoTo Errorproc
    
    base_sheet_name = "ボタンフォーム"
    target_sheet_name = "データ"
    Set base_ws = ThisWorkbook.Worksheets(base_sheet_name)
    Set target_ws = ThisWorkbook.Worksheets(target_sheet_name)
    
    input_ymd = Date
    member_name = base_ws.Range("C4").Value
    health_status = base_ws.Range("C7").Value
    detail = base_ws.Range("C10").Value
    
    last_row_num = target_ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    output_row_num = last_row_num + 1
    
    target_ws.Cells(output_row_num, 1) = input_ymd
    target_ws.Cells(output_row_num, 2) = member_name
    target_ws.Cells(output_row_num, 3) = health_status
    target_ws.Cells(output_row_num, 4) = detail
    
    ThisWorkbook.Save
    
    output_data = 0
    
    Exit Function
    
Errorproc:
    
    MsgBox output_check_error_msg, vbCritical
    output_data = 1
    

End Function


