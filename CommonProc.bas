Option Explicit

'定数
Public Const DATA_SHEET_NAME = "データ"
Public Const DATA_OUTPUT_FAILED = "データの出力に失敗しました。"
Public Const CONFIRM = "確認"


'【概要】配列をExcelファイルとして出力する
'【作成日】2023/01/10
Public Function ArrToExcelSheet(ByVal arrOutPut As Variant, _
                            ByVal strSheetName As String) As Boolean
On Error GoTo ArrToExcelSheet_Err
    
    Dim lngArrIdx As Long
    Dim lngCurrentRow As Long
    Dim objWb As Excel.Workbook
    
    ArrToExcelSheet = False
    
    'Excelのブック作成
    Set objWb = Workbooks.Add
    
    'シート名を設定
    ActiveSheet.Name = strSheetName
    
    '行を初期化
    lngCurrentRow = 1
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrOutPut)
        '出力
        objWb.Worksheets(strSheetName).Cells(lngCurrentRow, 1).Value = arrOutPut(lngArrIdx)
        '行をカウントアップ
        lngCurrentRow = lngCurrentRow + 1
    Next lngArrIdx
    
    ArrToExcelSheet = True
    
ArrToExcelSheet_Err:

ArrToExcelSheet_Exit:
    Set objWb = Nothing
End Function
