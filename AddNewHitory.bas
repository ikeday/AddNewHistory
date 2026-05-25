Attribute VB_Name = "AddNewHitory"
Option Explicit

' ==================================================
' モジュール名: AddNewHitory
' 目的:
' 1) FrontPage.xlsm の Sheets(2) A列(1行目開始)を製品名として読み取る
' 2) N列が "New" の製品は除外
' 3) products 配下の各製品ファイル Sheets(2) の最終行+1行に履歴を追記
'    - B,C列: 最終行をコピー
'    - D列: 06/01/2026
'    - E列: Adopted a new format of the Front Page
'    - 追加行の行高: 18.00
'    - 次行の行高: 6.00
'    - 追加行(B:E)の罫線: 格子
' ==================================================

Public Sub AddNewHitory()
    Dim masterSheet As Worksheet
    Dim productsFolder As String
    Dim lastMasterRow As Long
    Dim i As Long
    Dim productName As String
    Dim statusText As String
    Dim productFilePath As String

    Dim targetCount As Long
    Dim successCount As Long
    Dim skipCount As Long
    Dim errorCount As Long
    Dim errorDetails As String

    On Error GoTo FatalError

    Set masterSheet = ThisWorkbook.Sheets(2)
    productsFolder = "D:\Ｄ_Project\BSIフロントページ\products\"

    lastMasterRow = GetLastRow(masterSheet, 1)

    targetCount = 0
    successCount = 0
    skipCount = 0
    errorCount = 0
    errorDetails = ""

    Application.ScreenUpdating = False

    Debug.Print "=== AddNewHitory 開始 ==="
    Debug.Print "マスター: " & ThisWorkbook.Name
    Debug.Print "対象フォルダー: " & productsFolder

    For i = 1 To lastMasterRow
        productName = Trim(CStr(masterSheet.Cells(i, 1).Value))
        statusText = Trim(CStr(masterSheet.Cells(i, 14).Value)) ' N列

        If productName = "" Then
            skipCount = skipCount + 1
            Debug.Print "[行" & i & "] 製品名空白のためスキップ"

        ElseIf UCase(statusText) = "NEW" Then
            skipCount = skipCount + 1
            Debug.Print "[行" & i & "] " & productName & " は N列=New のため除外"

        Else
            targetCount = targetCount + 1
            productFilePath = productsFolder & productName & ".xlsm"

            If Not FileExists(productFilePath) Then
                errorCount = errorCount + 1
                errorDetails = errorDetails & vbCrLf & "  - " & productName & ".xlsm: ファイルが見つかりません"
                Debug.Print "  ✗ ファイルなし: " & productFilePath
            Else
                If AppendHistoryRow(productFilePath, productName, errorDetails) Then
                    successCount = successCount + 1
                Else
                    errorCount = errorCount + 1
                End If
            End If
        End If
    Next i

Cleanup:
    Application.ScreenUpdating = True

    ShowResult targetCount, successCount, skipCount, errorCount, errorDetails
    Exit Sub

FatalError:
    errorDetails = errorDetails & vbCrLf & "  - 致命的エラー: " & Err.Description
    Debug.Print "致命的エラー: " & Err.Description
    Resume Cleanup
End Sub

Private Function AppendHistoryRow(ByVal filePath As String, _
                                  ByVal productName As String, _
                                  ByRef errorDetails As String) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nextRow As Long

    On Error GoTo HandleError

    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets(2)

    lastRow = GetLastRow(ws, 2) ' B列基準
    If lastRow < 1 Then lastRow = 1
    nextRow = lastRow + 1

    ws.Range("B" & lastRow & ":C" & lastRow).Copy _
        Destination:=ws.Range("B" & nextRow & ":C" & nextRow)

    ws.Range("D" & nextRow).Value = DateSerial(2026, 6, 1)
    ws.Range("D" & nextRow).NumberFormat = "mm/dd/yyyy"
    ws.Range("D" & nextRow).HorizontalAlignment = xlCenter
    ws.Range("E" & nextRow).Value = "Adopted a new format of the Front Page"

    ws.Rows(nextRow).RowHeight = 18
    If nextRow < ws.Rows.Count Then
        ws.Rows(nextRow + 1).RowHeight = 6
    End If

    ApplyGridBorders ws.Range("B" & nextRow & ":E" & nextRow)

    wb.Save
    wb.Close SaveChanges:=False

    Debug.Print "  ✓ " & productName & ": 行" & nextRow & " に追記"
    AppendHistoryRow = True
    Exit Function

HandleError:
    errorDetails = errorDetails & vbCrLf & "  - " & productName & ".xlsm: " & Err.Description
    Debug.Print "  ✗ " & productName & ": " & Err.Description

    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    AppendHistoryRow = False
End Function

Private Sub ApplyGridBorders(ByVal targetRange As Range)
    With targetRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

Private Function GetLastRow(ByVal ws As Worksheet, ByVal columnNum As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, columnNum).End(xlUp).Row
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

Private Sub ShowResult(ByVal targetCount As Long, _
                       ByVal successCount As Long, _
                       ByVal skipCount As Long, _
                       ByVal errorCount As Long, _
                       ByVal errorDetails As String)
    Dim msg As String

    Debug.Print "=== AddNewHitory 完了 ==="
    Debug.Print "処理対象: " & targetCount
    Debug.Print "成功: " & successCount
    Debug.Print "除外/スキップ: " & skipCount
    Debug.Print "エラー: " & errorCount

    msg = "改訂履歴の追記が完了しました。" & vbCrLf & vbCrLf
    msg = msg & "処理対象: " & targetCount & " 製品" & vbCrLf
    msg = msg & "成功: " & successCount & " ファイル" & vbCrLf
    msg = msg & "除外/スキップ: " & skipCount & " 行" & vbCrLf
    msg = msg & "エラー: " & errorCount & " ファイル"

    If errorDetails <> "" Then
        msg = msg & vbCrLf & vbCrLf & "【エラー詳細】" & errorDetails
    End If

    MsgBox msg, vbInformation, "AddNewHitory"
End Sub
