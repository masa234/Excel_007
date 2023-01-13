Option Explicit
 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err

    Dim arrData() As Variant
    
    'データを配列で取得
    arrData = GetAllData(DATA_SHEET_NAME)
    
    '画面の更新をオフにする
    Application.ScreenUpdating = False
    
    '配列を出力
    If ArrToExcelSheet(arrData, DATA_SHEET_NAME) = False Then
        Call MsgBox(DATA_OUTPUT_FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
     
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
    '画面の更新をオンにする
    Application.ScreenUpdating = True
End Sub
