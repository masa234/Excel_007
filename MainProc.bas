Option Explicit


'【概要】全データを配列にセットする
'【作成日】2023/01/10
Public Function GetAllData(ByVal strSheetName As String) As Variant
On Error GoTo GetAllData_Err
    
    Dim lngLastRow As Long
    Dim lngCurrentRow As Long
    Dim lngLastCol As Long
    Dim lngCurrentCol As Long
    Dim lngArrIdx As Long
    Dim strValue As String
    Dim arrRet() As Variant

    '最終行を取得
    lngLastRow = ThisWorkbook.Worksheets(strSheetName).Cells(1, 1).End(xlDown).Row
    
    '配列の要素番号初期化
    lngArrIdx = 0
    
    '最終行まで繰り返す
    For lngCurrentRow = 1 To lngLastRow
        '最終列を取得
        lngLastCol = ThisWorkbook.Worksheets(strSheetName).Cells(lngCurrentRow, 1).End(xlToRight).Column
        '最終列まで繰り返す
        For lngCurrentCol = 1 To lngLastCol
            '値
            strValue = ThisWorkbook.Worksheets(strSheetName).Cells(lngCurrentRow, lngCurrentCol).Value
            '配列再宣言
            ReDim Preserve arrRet(lngArrIdx)
            '配列に格納
            arrRet(lngArrIdx) = strValue
            '配列の要素番号を1つ進める
            lngArrIdx = lngArrIdx + 1
        Next lngCurrentCol
    Next lngCurrentRow
    
    GetAllData = arrRet
    
GetAllData_Err:

GetAllData_Exit:

End Function

