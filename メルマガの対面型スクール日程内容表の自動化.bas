Option Explicit
' ここを環境に合わせて変更してください
Private Const BASE_PATH As String = "C:\Users\intern-yo-onodera\OneDrive - DMG MORI\テクニウム 教育サービスGP - 0_対面型スクール日程表作成"
'Private Const BASE_PATH As String = "C:\Users\intern-yk-shimizu\OneDrive - DMG MORI\メルマガ\田中李空＆遥馨さん＆清水さん共有\01_メルマガ資料\02_メルマガ詳細資料\0_対面型スクール日程表作成"

Private Const DEBUG_MODE As Boolean = False ' Trueにするとデバッグ用MsgBoxを表示

Private Sub DebugMessage(ByVal message As String)
    If DEBUG_MODE Then
        MsgBox message, vbInformation, "GenerateCategorySchedule - Debug"
    End If
End Sub

' 指定の形式でスケジュールシートをHTMLに変換して保存/貼付する
Public Sub ExportScheduleToHTML()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long, rr As Long, j As Long
    Dim html As String
    Dim courseName As String
    Dim sFile As String
    Dim fnum As Integer
    Dim timeStamp As String
    Dim monthsDict As Object
    Dim mList() As String
    Dim mKey As Variant
    Dim monthsLabel As String
    Dim swapped As Boolean
    Dim t As String
    Dim cellA As String
    Dim cellB As String
    Dim cellC As String
    Dim cellD As String
    Dim lookRow As Long
    Dim hasDataForCategory As Boolean
    Dim openTable As Boolean
    Dim isFirstTable As Boolean
    Dim courseStart As Long
    Dim courseEnd As Long
    Dim rowCount As Long
    Dim idx As Long
    Dim spanCount As Long
    Dim lastPlace As String
    Dim lastLink As String
    Dim thisPlace As String
    Dim escapedCourseName As String
    Dim escapedPlace As String
    Dim escapedDate As String
    Dim escapedSeat As String
    Dim rawSeat As Variant
    Dim nextA As String
    Dim nextB As String
    Dim nextC As String
    Dim nextD As String
    Dim placeArr() As String
    Dim linkArr() As String
    Dim dateArr() As String
    Dim seatArr() As String
    Dim placeSpan() As Long
    Dim advance As Long
    Dim dateParam As String
    Dim startAlpha As String
    Dim currentAlpha As String
    Dim urlAlphaMap As Object

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("スケジュール")

    ' URLパラメータの初期化
    dateParam = Format(ws.Range("G14").value, "yymmdd")
    startAlpha = LCase(Trim(ws.Range("G15").value))
    currentAlpha = startAlpha
    Set urlAlphaMap = CreateObject("Scripting.Dictionary")

    Dim lastA As Long, lastB As Long, lastC As Long, lastD As Long
    lastA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastC = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    lastD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRow = lastA
    If lastB > lastRow Then lastRow = lastB
    If lastC > lastRow Then lastRow = lastC
    If lastD > lastRow Then lastRow = lastD
    If lastRow < 3 Then
        MsgBox "スケジュール表が空です。", vbInformation
        Exit Sub
    End If

    Set monthsDict = CreateObject("Scripting.Dictionary")
    For r = 3 To lastRow
        cellC = CStr(ws.Cells(r, "C").value)
        If Len(Trim$(cellC)) > 0 Then
            j = InStr(1, cellC, "/")
            If j > 1 Then
                On Error Resume Next
                Dim m As Long
                m = CLng(Trim$(Left$(cellC, j - 1)))
                If Err.Number = 0 Then
                    If Not monthsDict.Exists(CStr(m)) Then monthsDict.Add CStr(m), True
                End If
                On Error GoTo ErrHandler
            End If
        End If
    Next r

    If monthsDict.Count > 0 Then
        ReDim mList(0 To monthsDict.Count - 1)
        j = 0
        For Each mKey In monthsDict.Keys
            mList(j) = mKey
            j = j + 1
        Next mKey

        Do
            swapped = False
            For j = 0 To UBound(mList) - 1
                If CLng(mList(j)) > CLng(mList(j + 1)) Then
                    t = mList(j)
                    mList(j) = mList(j + 1)
                    mList(j + 1) = t
                    swapped = True
                End If
            Next j
        Loop While swapped
    End If

    monthsLabel = ""
    If monthsDict.Count > 0 Then
        For j = 0 To UBound(mList)
            If monthsLabel <> "" Then monthsLabel = monthsLabel & "、"
            monthsLabel = monthsLabel & mList(j)
        Next j
        monthsLabel = monthsLabel & "月講座一覧 表"
    End If

    html = "<!-- " & monthsLabel & " -->" & vbCrLf & vbCrLf & "<table>" & vbCrLf

    ' 最初のカテゴリヘッダを先頭に出す場合、スケジュールシートの先頭行にカテゴリがあることを期待して
    ' ループ前に先頭のカテゴリ行を探して出力しておく（ユーザの要求: 一番上に最初のカテゴリの表示）
    Dim firstCatFound As Boolean
    firstCatFound = False
    Dim firstCatName As String
    For r = 1 To lastRow
        If Trim$(CStr(ws.Cells(r, "A").value)) <> "" Then
            If Left$(Trim$(CStr(ws.Cells(r, "A").value)), 1) = "〇" Then
                firstCatName = Trim$(CStr(ws.Cells(r, "A").value))
                html = html & "<tr>" & vbCrLf & _
                            "<td style=""height:12px!important;font-size:12px!important;line-height:0.5!important;padding:6px 8px!important;margin-bottom:6px;"" width=""100%"" class=""responsive-td"" valign=""top"" align=""left"">" & vbCrLf & _
                            "<font size=""4"" color=""#000000"" style=""font-size:14px;color:#000000;line-height:1;"">" & Replace$(firstCatName, "&", "&amp;") & "</font>" & vbCrLf & _
                            "</td></tr>" & vbCrLf & vbCrLf
                firstCatFound = True
            End If
            Exit For
        End If
    Next r

    openTable = False
    isFirstTable = True
    rr = 3
    Do While rr <= lastRow
        advance = 1

        cellA = Trim$(CStr(ws.Cells(rr, "A").value))
        cellB = Trim$(CStr(ws.Cells(rr, "B").value))
        cellC = Trim$(CStr(ws.Cells(rr, "C").value))
        cellD = Trim$(CStr(ws.Cells(rr, "D").value))

        If Len(cellA) > 0 And Left$(cellA, 1) = "〇" And cellB = "" And cellC = "" And cellD = "" Then
            If openTable Then
                html = html & "        </tbody>" & vbCrLf & _
                              "    </table>" & vbCrLf & _
                              "</td>" & vbCrLf & vbCrLf
                openTable = False
            End If

            hasDataForCategory = False
            lookRow = rr + 1
            Do While lookRow <= lastRow
                nextA = Trim$(CStr(ws.Cells(lookRow, "A").value))
                nextB = Trim$(CStr(ws.Cells(lookRow, "B").value))
                nextC = Trim$(CStr(ws.Cells(lookRow, "C").value))
                nextD = Trim$(CStr(ws.Cells(lookRow, "D").value))
                If Len(nextA) > 0 And Left$(nextA, 1) = "〇" And nextB = "" And nextC = "" And nextD = "" Then
                    Exit Do
                End If
                If nextA <> "" Or nextB <> "" Or nextC <> "" Or nextD <> "" Then
                    hasDataForCategory = True
                    Exit Do
                End If
                lookRow = lookRow + 1
            Loop

            ' カテゴリ間の余白を確保（trの前に余白用trを挿入）
            html = html & "<tr>" & vbCrLf & _
                          "<td style=""height:12px!important;font-size:12px!important;line-height:0.5!important;padding:6px 8px!important;"" width=""100%"" class=""responsive-td"" valign=""top"" align=""left"">" & vbCrLf & _
                          "<font size=""4"" color=""#000000"" style=""font-size:14px;color:#000000;line-height:1;"">" & Replace$(cellA, "&", "&amp;") & "</font>" & vbCrLf & _
                          "</td></tr>" & vbCrLf & vbCrLf

            If hasDataForCategory Then
                ' テーブルの前に少し余白を挿入してカテゴリとテーブルの間隔を確保
                html = html & "<td width=""100%"" class=""responsive-td"" valign=""top"" style=""padding: 10px;"">" & vbCrLf
                html = html & "<table border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse; font-size: 14px; line-height: 1.2; width: 100%; vertical-align: middle;"">" & vbCrLf
                html = html & "<colgroup><col width=""50%;""><col width=""10%;""><col width=""30%;""><col width=""10%;""></colgroup>" & vbCrLf
                html = html & "<tbody>" & vbCrLf
                If isFirstTable Then
                    html = html & "<tr style=""background-color: #555555; color: #ffffff; line-height: 1.4;"">" & vbCrLf
                    html = html & "<th width=""50%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;text-align: center!important;border: 1px solid #000!important;"">コース名</th>" & vbCrLf
                    html = html & "<th width=""10%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;text-align: center!important;border: 1px solid #000!important;"">場所</th>" & vbCrLf
                    html = html & "<th width=""30%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;text-align: center!important;border: 1px solid #000!important;"">日程</th>" & vbCrLf
                    html = html & "<th width=""10%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;text-align: center!important;border: 1px solid #000!important;"">残席数</th>" & vbCrLf
                    html = html & "</tr>" & vbCrLf
                End If
                html = html & "</tbody>" & vbCrLf
                html = html & "<tbody>" & vbCrLf
                openTable = True
                isFirstTable = False
            End If

            GoTo AdvanceLoop
        End If

        If cellA = "" And cellB = "" And cellC = "" And cellD = "" Then
            GoTo AdvanceLoop
        End If

        If cellA = "" Then
            GoTo AdvanceLoop
        End If

        If Not openTable Then
            html = html & "<td width=""100%"" class=""responsive-td"" valign=""top"" style=""padding: 10px;"">" & vbCrLf
            html = html & "<table border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse; font-size: 14px; line-height: 1.2; width: 100%; vertical-align: middle;"">" & vbCrLf
            html = html & "<colgroup><col width=""50%;""><col width=""10%;""><col width=""30%;""><col width=""10%;""></colgroup>" & vbCrLf
            html = html & "<tbody>" & vbCrLf
            If isFirstTable Then
                html = html & "<tr style=""background-color: #555555; color: #ffffff; line-height: 1.4;"">" & vbCrLf
                html = html & "<th width=""50%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;text-align: center!important;border: 1px solid #000!important;"">コース名</th>" & vbCrLf
                html = html & "<th width=""10%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;text-align: center!important;border: 1px solid #000!important;"">場所</th>" & vbCrLf
                html = html & "<th width=""30%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;text-align: center!important;border: 1px solid #000!important;"">日程</th>" & vbCrLf
                html = html & "<th width=""10%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;text-align: center!important;border: 1px solid #000!important;"">残席数</th>" & vbCrLf
                html = html & "</tr>" & vbCrLf
            End If
            html = html & "</tbody>" & vbCrLf
            html = html & "<tbody>" & vbCrLf
            openTable = True
            isFirstTable = False
        End If

        courseStart = rr
        courseEnd = rr
        Do While courseEnd + 1 <= lastRow
            nextA = Trim$(CStr(ws.Cells(courseEnd + 1, "A").value))
            nextB = Trim$(CStr(ws.Cells(courseEnd + 1, "B").value))
            nextC = Trim$(CStr(ws.Cells(courseEnd + 1, "C").value))
            nextD = Trim$(CStr(ws.Cells(courseEnd + 1, "D").value))
            If nextA <> "" Then Exit Do
            If nextB = "" And nextC = "" And nextD = "" Then Exit Do
            courseEnd = courseEnd + 1
        Loop

        rowCount = courseEnd - courseStart + 1
        If rowCount <= 0 Then GoTo AdvanceLoop

        ReDim placeArr(1 To rowCount)
        ReDim linkArr(1 To rowCount)
        ReDim dateArr(1 To rowCount)
        ReDim seatArr(1 To rowCount)
        ReDim placeSpan(1 To rowCount)

        lastPlace = ""
        lastLink = ""
        For idx = 1 To rowCount
            thisPlace = Trim$(CStr(ws.Cells(courseStart + idx - 1, "B").value))
            If thisPlace <> "" Then
                lastPlace = thisPlace
                If ws.Cells(courseStart + idx - 1, "B").Hyperlinks.Count > 0 Then
                    lastLink = ws.Cells(courseStart + idx - 1, "B").Hyperlinks(1).Address
                Else
                    lastLink = ""
                End If
            End If
            placeArr(idx) = lastPlace
            linkArr(idx) = lastLink

            dateArr(idx) = Trim$(CStr(ws.Cells(courseStart + idx - 1, "C").value))
            rawSeat = ws.Cells(courseStart + idx - 1, "D").value
            If IsNumeric(rawSeat) Then
                seatArr(idx) = CStr(rawSeat)
            Else
                seatArr(idx) = Trim$(CStr(rawSeat))
            End If
        Next idx

        idx = 1
        Do While idx <= rowCount
            spanCount = 1
            Do While idx + spanCount <= rowCount
                If StrComp(placeArr(idx + spanCount), placeArr(idx), vbTextCompare) = 0 Then
                    spanCount = spanCount + 1
                Else
                    Exit Do
                End If
            Loop
            placeSpan(idx) = spanCount
            idx = idx + spanCount
        Loop

        courseName = CStr(ws.Cells(courseStart, "A").value)
        escapedCourseName = Replace$(courseName, "&", "&amp;")

        For idx = 1 To rowCount
            If idx = 1 Then
                html = html & "<tr style=""background-color: #ffffff;"">" & vbCrLf
                html = html & "<td width=""50%"" rowspan=""" & rowCount & """ style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;border:1px solid #000!important;vertical-align: middle; text-align: center;"">" & escapedCourseName & "</td>" & vbCrLf
            Else
                html = html & "<tr>" & vbCrLf
            End If

            If placeSpan(idx) > 0 Then
                escapedPlace = Replace$(placeArr(idx), "&", "&amp;")
                If linkArr(idx) <> "" And escapedPlace <> "" Then
                    Dim fullUrl As String
                    If Not urlAlphaMap.Exists(linkArr(idx)) Then
                        urlAlphaMap.Add linkArr(idx), currentAlpha
                        currentAlpha = NextAlpha(currentAlpha)
                    End If
                    fullUrl = BuildTrackedUrl(linkArr(idx), dateParam, urlAlphaMap(linkArr(idx)), True)
                    html = html & "<td width=""10%"" rowspan=""" & placeSpan(idx) & """ style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;border:1px solid #000!important;vertical-align: middle; text-align: center;""><a href=""" & fullUrl & """ style=""color:#008eef;text-decoration:underline;"" target=""_blank"">" & escapedPlace & "</a></td>" & vbCrLf
                Else
                    html = html & "<td width=""10%"" rowspan=""" & placeSpan(idx) & """ style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;border:1px solid #000!important;vertical-align: middle; text-align: center;"">" & escapedPlace & "</td>" & vbCrLf
                End If
            End If

            escapedDate = Replace$(dateArr(idx), "&", "&amp;")
            escapedSeat = Replace$(seatArr(idx), "&", "&amp;")
            html = html & "<td width=""30%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;border:1px solid #000!important;text-align: center;"">" & escapedDate & "</td>" & vbCrLf
            html = html & "<td width=""10%"" style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;border:1px solid #000!important;text-align: center;"">" & escapedSeat & "</td>" & vbCrLf
            html = html & "</tr>" & vbCrLf
        Next idx

        advance = rowCount

AdvanceLoop:
        rr = rr + advance
    Loop

    If openTable Then
        html = html & "        </tbody>" & vbCrLf & _
                      "    </table>" & vbCrLf & _
                      "</td>" & vbCrLf & vbCrLf
        openTable = False
    End If

    html = html & "</table>" & vbCrLf

    timeStamp = Format$(Now, "yyyymmdd_Hhnnss")
    sFile = BASE_PATH & "\schedule_export_" & timeStamp & ".html"
    fnum = FreeFile
    Open sFile For Output As #fnum
    Print #fnum, html
    Close #fnum

    MsgBox "HTML を生成し、ファイルに保存しました: " & sFile, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "HTML エクスポート中にエラーが発生しました: " & Err.Description, vbExclamation
End Sub

' セル結合時の警告を抑制して安全に結合するヘルパー
Private Sub SafeMerge(ByVal rng As Range)
    On Error Resume Next
    Dim topVal As Variant
    Dim c As Range
    If rng Is Nothing Then Exit Sub
    topVal = rng.Cells(1, 1).value
    ' 右下などのセルに値が入っている場合はクリアして警告を回避
    For Each c In rng.Cells
        If c.Address <> rng.Cells(1, 1).Address Then
            If Len(Trim$(CStr(c.value))) > 0 Then c.ClearContents
        End If
    Next c
    Application.DisplayAlerts = False
    rng.Merge
    Application.DisplayAlerts = True
    rng.Cells(1, 1).value = topVal
End Sub

Private Function NormalizeDateText(ByVal value As Variant) As String
    Dim txt As String

    txt = Trim$(CStr(value))
    If Len(txt) = 0 Then Exit Function

    ' 全角／異体字などを標準化
    txt = Replace$(txt, "年", "/")
    txt = Replace$(txt, "月", "/")
    txt = Replace$(txt, "日", "")
    txt = Replace$(txt, "－", "-")
    txt = Replace$(txt, "―", "-")
    txt = Replace$(txt, "ー", "-")
    txt = Replace$(txt, "／", "/")
    txt = Replace$(txt, "\", "")
    txt = Replace$(txt, ChrW$(65295), "/") ' 全角スラッシュ
    txt = Replace$(txt, ChrW$(8722), "-") ' マイナス記号
    txt = Replace$(txt, vbCr, "")
    txt = Replace$(txt, vbLf, "")
    txt = Replace$(txt, "　", " ")
    txt = Trim$(txt)

    NormalizeDateText = txt
End Function

Private Function TryParseDate(ByVal value As Variant, ByRef parsedDate As Date, Optional ByVal fallbackYear As Variant) As Boolean
    Dim txt As String
    Dim yearNumber As Long
    Dim yearText As String
    Dim combinedTxt As String
    Dim hasYearInValue As Boolean
    Dim hasFallbackYear As Boolean
    Dim parts As Variant
    Dim part As Variant

    On Error GoTo CleanFail

    txt = NormalizeDateText(value)

    If Len(txt) = 0 Then
        If IsDate(value) Then
            parsedDate = CDate(value)
            TryParseDate = True
        End If
        Exit Function
    End If

    ' 数字だけの場合 (日付シリアル) は value の日付を優先
    If InStr(txt, "/") = 0 And InStr(txt, "-") = 0 And InStr(txt, ".") = 0 Then
        If IsDate(value) Then
            txt = Format$(CDate(value), "yyyy/m/d")
        End If
    End If

    hasYearInValue = ContainsYear(txt)

    ' 補助年の指定があれば利用
    If Not IsMissing(fallbackYear) Then
        If IsNumeric(fallbackYear) Then
            yearNumber = CLng(fallbackYear)
        ElseIf IsDate(fallbackYear) Then
            yearNumber = Year(CDate(fallbackYear))
        Else
            yearText = NormalizeDateText(fallbackYear)
            If Len(yearText) >= 4 Then
                parts = Split(Replace(Replace(yearText, "-", "/"), ".", "/"), "/")
                For Each part In parts
                    If Len(part) >= 4 And IsNumeric(part) Then
                        yearNumber = CLng(part)
                        Exit For
                    End If
                Next part
                If yearNumber = 0 And IsDate(yearText) Then
                    yearNumber = Year(CDate(yearText))
                End If
            End If
        End If
        If yearNumber <> 0 Then
            hasFallbackYear = True
        End If
    End If

    ' 元の値やテキスト内に年が含まれている場合はそのまま解釈
    ' フォールバック年が指定されている場合は、たとえ G 列の値に年が含まれていても
    ' まずフォールバック年での合成を試みる（ユーザ指定の年を優先するため）
    If hasFallbackYear Then
        Dim cleanTxt2 As String
        Dim mdParts As String
        Dim p As String
        cleanTxt2 = txt
        Do While Left$(cleanTxt2, 1) = "/" Or Left$(cleanTxt2, 1) = "-" Or Left$(cleanTxt2, 1) = "."
            cleanTxt2 = Mid$(cleanTxt2, 2)
        Loop
        If Len(cleanTxt2) > 0 Then
            ' 区切り文字を統一してパーツ化
            parts = Split(Replace(Replace(cleanTxt2, "-", "/"), ".", "/"), "/")
            mdParts = ""
            For Each part In parts
                p = Trim$(part)
                If Len(p) > 0 Then
                    ' 4桁以上の数値かつ年域(1900-2100)は除外し、残りを月日として扱う
                    If Not (Len(p) >= 4 And IsNumeric(p) And CLng(p) >= 1900 And CLng(p) <= 2100) Then
                        If mdParts = "" Then
                            mdParts = p
                        Else
                            mdParts = mdParts & "/" & p
                        End If
                    End If
                End If
            Next part
            If Len(mdParts) > 0 Then
                combinedTxt = CStr(yearNumber) & "/" & mdParts
                If IsDate(combinedTxt) Then
                    parsedDate = CDate(combinedTxt)
                    TryParseDate = True
                    Exit Function
                End If
                combinedTxt = mdParts & "/" & CStr(yearNumber)
                If IsDate(combinedTxt) Then
                    parsedDate = CDate(combinedTxt)
                    TryParseDate = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' フォールバック年がない場合は、元の値やテキスト内に年が含まれていればそのまま解釈
    If Not hasFallbackYear Then
        If hasYearInValue Then
            If IsDate(txt) Then
                parsedDate = CDate(txt)
                TryParseDate = True
                Exit Function
            ElseIf IsDate(value) Then
                parsedDate = CDate(value)
                TryParseDate = True
                Exit Function
            End If
        End If
    End If

    ' ここまでで確定できない場合は value の日付評価にフォールバック
    If IsDate(value) Then
        parsedDate = CDate(value)
        TryParseDate = True
        Exit Function
    End If

    If IsDate(txt) Then
        parsedDate = CDate(txt)
        TryParseDate = True
        Exit Function
    End If

    Exit Function

CleanFail:
    TryParseDate = False
End Function

Private Function ContainsYear(ByVal s As String) As Boolean
    Dim t As String, parts As Variant, p As Variant, y As Long
    t = NormalizeDateText(s)
    parts = Split(Replace(Replace(t, "-", "/"), ".", "/"), "/")
    For Each p In parts
        If Len(p) >= 4 And IsNumeric(p) Then
            y = CLng(p)
            If y >= 1900 And y <= 2100 Then
                ContainsYear = True
                Exit Function
            End If
        End If
    Next p
End Function


'**********************************************************************
' マクロ名 : GenerateCategorySchedule
' 役割     : 指定したカテゴリと日付範囲に該当するスクール情報を
'            「エクセル1.xlsx」「エクセル2.xlsx」から抽出し、
'            「エクセル3.xlsm」のアクティブシートに一覧を作成する。
'            「コースとURL対応表.xlsx」でURLリンクも付与
' 利用手順 :
'   1) 「エクセル3.xlsm」のG3セルにカテゴリ名を入力する。
'   2) G4セルに抽出対象となる開始日の下限、I4セルに上限を入力する。
'   3) 本マクロを実行する。
'   4) A列～D列にカテゴリ→コース別の一覧が自動生成される。
'**********************************************************************
Public Sub GenerateCategorySchedule()

    '=== ★ 基本設定エリア =================================================
    Const FILE_EXCEL1 As String = "エクセル1.xlsx"          ' コース毎の詳細 (場所・期間・残席など)
    Const FILE_EXCEL2 As String = "エクセル2.xlsx"          ' カテゴリとコースの対応表
    Const FILE_EXCEL3 As String = "エクセル3.xlsm"          ' 本マクロを格納しているブック (参照用の定数。念のため保持)
    Const FILE_URL_MAP As String = "コースとURL対応表.xlsx" ' コースとURLの対応表 (A列=コース名, B列=場所, E列=URL)

    '=== ★ 変数宣言 =======================================================
    Dim targetCategory As String            ' ユーザがG3に入力したカテゴリ名
    Dim filterStartDate As Date             ' 検索対象となる開始日(下限)
    Dim filterEndDate As Date               ' 検索対象となる開始日(上限)

    Dim wbDetail As Workbook                ' エクセル1.xlsx
    Dim wbCategory As Workbook              ' エクセル2.xlsx
    Dim wbUrl As Workbook                   ' コースとURL対応表.xlsx
    Dim wsDetail As Worksheet
    Dim wsCategory As Worksheet
    Dim wsOutput As Worksheet
    Dim wsUrl As Worksheet

    Dim lastRowDetail As Long
    Dim lastRowCategory As Long
    Dim lastRowUrl As Long

    Dim courseList As Collection            ' カテゴリ一致のコース名を格納
    Dim courseSchedules As Object           ' Scripting.Dictionary : Key=コース名, Item=Collection(各スケジュール)
    Dim urlMap As Object                    ' Scripting.Dictionary : Key="コース名|場所", Item=URL
    Dim courseSeen As Object                 ' Scripting.Dictionary : 重複コース名を排除するためのフラグ

    Dim courseName As String
    Dim locationName As String
    Dim scheduleStart As Variant
    Dim scheduleEnd As Variant
    Dim tmpDate As Date
    Dim seatValue As Variant
    Dim cName As String

    Dim rowIdx As Long
    Dim outputRow As Long
    Dim scheduleInfo As Variant
    Dim scheduleList As Collection
    Dim i As Long
    Dim scheduleIdx As Long
    Dim currentRow As Long
    Dim firstRowForCourse As Long
    Dim firstCatKey As String
    Dim mergeRngA As Range
    Dim courseUrl As String
    Dim urlLocation As String
    Dim urlKey As String
    Dim debugBuffer As String
    Dim scheduleAddedCount As Long
    Dim skippedBeforeRange As Long
    Dim skippedAfterRange As Long
    Dim invalidStartDateCount As Long
    Dim nonTargetCourseCount As Long

    Dim detailOpened As Boolean             ' コード内で開いたかどうか (Trueなら終了時にClose)
    Dim categoryOpened As Boolean
    Dim urlOpened As Boolean

    Dim calcState As XlCalculation          ' 計算モード退避用
    Dim screenUpdatingState As Boolean
    Dim enableEventsState As Boolean
    Dim bRow As Long
    Dim segStart As Long
    Dim segEnd As Long
    Dim mergeRngB As Range
    Dim mergeRngB2 As Range
    Dim isAllCategories As Boolean
    Dim dateParam As String
    Dim startAlpha As String
    Dim currentAlpha As String
    Dim urlAlphaMap As Object

    On Error GoTo ErrHandler

    '=== ★ 実行前のアプリケーション状態を退避 =============================
    screenUpdatingState = Application.ScreenUpdating
    Application.ScreenUpdating = False

    calcState = Application.Calculation
    Application.Calculation = xlCalculationManual

    enableEventsState = Application.EnableEvents
    Application.EnableEvents = False

    '=== ★ 入力シートの参照を取得 =========================================
    ' 出力先をシート名で指定
    Set wsOutput = ThisWorkbook.Worksheets("スケジュール") ' 出力先は「エクセル3.xlsm」のシート名 'スケジュール'

    '=== ★ 入力チェック (カテゴリ名・日付範囲) ============================
    ' チェックボックスから選択されたカテゴリを収集
    Dim selectedCategories As Collection
    Set selectedCategories = New Collection
    
    Dim categoryMappings As Object
    Set categoryMappings = CreateObject("Scripting.Dictionary")
    categoryMappings.Add "F4", "ターニングセンタ"
    categoryMappings.Add "F5", "複合加工機"
    categoryMappings.Add "F6", "マシニングセンタ"
    categoryMappings.Add "F7", "5軸加工機"
    categoryMappings.Add "F8", "マクロプログラミング"
    categoryMappings.Add "F9", "メンテナンス"
    categoryMappings.Add "F10", "産業用ロボット"
    
    Dim cell As Range
    For Each cell In wsOutput.Range("F4:F10")
        If cell.value = True Then
            selectedCategories.Add categoryMappings(cell.Address(RowAbsolute:=False, ColumnAbsolute:=False))
        End If
    Next cell
    
    If selectedCategories.Count = 0 Then
        Err.Raise vbObjectError + 100, "GenerateCategorySchedule", "F4～F10のチェックボックスで少なくとも1つのカテゴリを選択してください。"
    End If

    isAllCategories = (selectedCategories.Count > 1)

    ' URLパラメータの初期化
    dateParam = Format(wsOutput.Range("G14").value, "yymmdd")
    startAlpha = LCase(Trim(wsOutput.Range("G15").value))
    currentAlpha = startAlpha
    Set urlAlphaMap = CreateObject("Scripting.Dictionary")

    If Not TryParseDate(wsOutput.Range("G12").value, filterStartDate) Then
        Err.Raise vbObjectError + 101, "GenerateCategorySchedule", "G12セルの開始日が日付として認識できません。"
    End If

    If Not TryParseDate(wsOutput.Range("I12").value, filterEndDate) Then
        Err.Raise vbObjectError + 102, "GenerateCategorySchedule", "I12セルの終了日が日付として認識できません。"
    End If

    If filterEndDate < filterStartDate Then
        Err.Raise vbObjectError + 103, "GenerateCategorySchedule", "終了日(I12)が開始日(G12)よりも前の日付になっています。"
    End If

    '=== ★ 参照ファイルのオープン =========================================
    ' ※ Dir関数で存在確認を行い、なければエラーを通知する。
    If Dir(BASE_PATH & "\" & FILE_EXCEL1) = "" Then
        Err.Raise vbObjectError + 104, "GenerateCategorySchedule", _
                  "ファイルが見つかりません: " & FILE_EXCEL1
    End If
    If Dir(BASE_PATH & "\" & FILE_EXCEL2) = "" Then
        Err.Raise vbObjectError + 105, "GenerateCategorySchedule", _
                  "ファイルが見つかりません: " & FILE_EXCEL2
    End If
    If Dir(BASE_PATH & "\" & FILE_URL_MAP) = "" Then
        Err.Raise vbObjectError + 107, "GenerateCategorySchedule", _
              "ファイルが見つかりません: " & FILE_URL_MAP
    End If

    Set wbDetail = Workbooks.Open(BASE_PATH & "\" & FILE_EXCEL1, ReadOnly:=True)
    detailOpened = True
    Set wsDetail = wbDetail.Worksheets("空席情報2025_26年1Q")

    Set wbCategory = Workbooks.Open(BASE_PATH & "\" & FILE_EXCEL2, ReadOnly:=True)
    categoryOpened = True
    Set wsCategory = wbCategory.Worksheets("カテゴリ")

    Set urlMap = CreateObject("Scripting.Dictionary")
    Set wbUrl = Workbooks.Open(BASE_PATH & "\" & FILE_URL_MAP, ReadOnly:=True)
    urlOpened = True
    Set wsUrl = wbUrl.Worksheets("コースとURL対応表")

    lastRowUrl = wsUrl.Cells(wsUrl.Rows.Count, "E").End(xlUp).Row
    For rowIdx = 1 To lastRowUrl
        courseName = Trim$(wsUrl.Cells(rowIdx, "A").value)
        If Len(courseName) > 0 Then
            urlLocation = Trim$(wsUrl.Cells(rowIdx, "B").value)
            courseUrl = Trim$(wsUrl.Cells(rowIdx, "E").value)
            If Len(urlLocation) > 0 And Len(courseUrl) > 0 Then
                urlKey = LCase$(courseName) & "|" & LCase$(urlLocation)
                If Not urlMap.Exists(urlKey) Then
                    urlMap.Add urlKey, courseUrl
                End If
            End If
        End If
    Next rowIdx

    '=== ★ カテゴリに紐づくコース候補を収集 ===============================
    Set courseList = New Collection
    Set courseSeen = CreateObject("Scripting.Dictionary")

    lastRowCategory = wsCategory.Cells(wsCategory.Rows.Count, "A").End(xlUp).Row
    Dim cat As Variant
    For Each cat In selectedCategories
        For rowIdx = 2 To lastRowCategory
            If StrComp(Trim$(wsCategory.Cells(rowIdx, "A").value), "〇" & cat, vbTextCompare) = 0 Then
                courseName = Trim$(wsCategory.Cells(rowIdx, "B").value)
                If Len(courseName) > 0 Then
                    If Not courseSeen.Exists(courseName) Then
                        courseList.Add courseName
                        courseSeen.Add courseName, True
                    End If
                End If
            End If
        Next rowIdx
    Next cat

    If courseList.Count = 0 Then
        Err.Raise vbObjectError + 106, "GenerateCategorySchedule", _
                  "選択されたカテゴリに紐づくコースがエクセル2.xlsxで見つかりませんでした。"
    End If    '=== ★ コースごとのスケジュール格納用ディクショナリ ==================
    Set courseSchedules = CreateObject("Scripting.Dictionary")
    For i = 1 To courseList.Count
        If Not courseSchedules.Exists(courseList(i)) Then
            Set scheduleList = New Collection
            courseSchedules.Add courseList(i), scheduleList
        End If
    Next i

    ' G3 が全て指定のときはカテゴリごとに出力するためのマップを作成
    Dim categoryMap As Object ' Key=カテゴリ名, Item=Collection of course names
    Set categoryMap = CreateObject("Scripting.Dictionary")
    For Each cat In selectedCategories
        Set categoryMap(cat) = New Collection
        For rowIdx = 2 To lastRowCategory
            If StrComp(Trim$(wsCategory.Cells(rowIdx, "A").value), "〇" & cat, vbTextCompare) = 0 Then
                cName = Trim$(CStr(wsCategory.Cells(rowIdx, "B").value))
                If Len(cName) > 0 Then
                    On Error Resume Next
                    categoryMap(cat).Add cName, CStr(cName)
                    On Error GoTo ErrHandler
                End If
            End If
        Next rowIdx
    Next cat

    '=== ★ エクセル1.xlsxから該当データを抽出 =============================
    lastRowDetail = wsDetail.Cells(wsDetail.Rows.Count, "C").End(xlUp).Row
    For rowIdx = 6 To lastRowDetail
        courseName = Trim$(wsDetail.Cells(rowIdx, "C").value)
        Dim rowStatus As String
        Dim parsedStartText As String
        Dim rawStart As String
        rawStart = Trim$(CStr(wsDetail.Cells(rowIdx, "G").value))
        If courseSchedules.Exists(courseName) Then
            If TryParseDate(wsDetail.Cells(rowIdx, "G").value, tmpDate, wsDetail.Cells(rowIdx, "A").value) Then
                scheduleStart = tmpDate
                parsedStartText = Format$(scheduleStart, "yyyy/mm/dd")
                ' 指定範囲に含まれる開始日のみ対象
                If scheduleStart >= filterStartDate And scheduleStart <= filterEndDate Then
                    If TryParseDate(wsDetail.Cells(rowIdx, "H").value, tmpDate, wsDetail.Cells(rowIdx, "A").value) Then
                        scheduleEnd = tmpDate
                    Else
                        scheduleEnd = vbNullString  ' 終了日が空欄の場合は空文字を保持
                    End If
                    locationName = Trim$(wsDetail.Cells(rowIdx, "B").value)
                    seatValue = wsDetail.Cells(rowIdx, "K").value
                    ' スケジュール情報を配列にまとめて保持
                    scheduleInfo = Array(locationName, scheduleStart, scheduleEnd, seatValue)
                    courseSchedules(courseName).Add scheduleInfo
                    scheduleAddedCount = scheduleAddedCount + 1
                    rowStatus = "IN"
                ElseIf scheduleStart < filterStartDate Then
                    skippedBeforeRange = skippedBeforeRange + 1
                    rowStatus = "OUT_BEFORE"
                Else
                    skippedAfterRange = skippedAfterRange + 1
                    rowStatus = "OUT_AFTER"
                End If
            Else
                invalidStartDateCount = invalidStartDateCount + 1
                rowStatus = "INVALID"
            End If
        Else
            nonTargetCourseCount = nonTargetCourseCount + 1
            rowStatus = "NON_TARGET"
        End If
        ' 各行のデバッグ情報を蓄積
        If DEBUG_MODE Then
            If parsedStartText = "" Then parsedStartText = rawStart
            debugBuffer = debugBuffer & "行" & rowIdx & ": " & courseName & " - " & parsedStartText & " - " & rowStatus & vbCrLf
        End If
    Next rowIdx

    If DEBUG_MODE Then
        debugBuffer = "◆デバッグ概要" & vbCrLf & _
            "カテゴリ: " & targetCategory & vbCrLf & _
            "抽出対象コース数: " & courseList.Count & vbCrLf & _
            "URLマップ件数: " & urlMap.Count & vbCrLf & _
            "抽出されたスケジュール数: " & scheduleAddedCount & vbCrLf & _
            "期間外(開始日が下限より前): " & skippedBeforeRange & vbCrLf & _
            "期間外(開始日が上限より後): " & skippedAfterRange & vbCrLf & _
            "開始日が日付形式でない行: " & invalidStartDateCount & vbCrLf & _
            "カテゴリ対象外の行数: " & nonTargetCourseCount & vbCrLf & vbCrLf & _
            "◆コース別スケジュール件数" & vbCrLf

        For i = 1 To courseList.Count
            debugBuffer = debugBuffer & "  ・" & courseList(i) & " : " & _
                courseSchedules(courseList(i)).Count & "件" & vbCrLf
        Next i

        ' --- ワークシートにログを書き出す ---
        Dim dbgWS As Worksheet
        Dim lines As Variant
        Dim r As Long
        On Error Resume Next
        Set dbgWS = ThisWorkbook.Worksheets("DebugLog")
        On Error GoTo 0
        If dbgWS Is Nothing Then
            Set dbgWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            On Error Resume Next
            dbgWS.Name = "DebugLog"
            On Error GoTo 0
        Else
            dbgWS.Cells.Clear
        End If

        lines = Split(debugBuffer, vbCrLf)
        For r = LBound(lines) To UBound(lines)
            dbgWS.Cells(r + 1, 1).value = lines(r)
        Next r
        dbgWS.Columns(1).EntireColumn.AutoFit

        ' 互換性のため MsgBox も残す（必要なければ削除可）
        ' DebugMessage debugBuffer
    End If

    '=== ★ 出力シートの既存レイアウトをクリア ============================
    With wsOutput.Range("A1:D1000")
        .Clear
        .ClearFormats
    End With

    '=== ★ カテゴリタイトルとヘッダー行を整形 ============================
    ' ヘッダ表示: 選択された最初のカテゴリ名をA1に表示
    If selectedCategories.Count > 0 Then
        wsOutput.Range("A1").value = "〇" & selectedCategories(1)
    Else
        wsOutput.Range("A1").value = "講座一覧（選択カテゴリ）"
    End If
    wsOutput.Range("A1").Font.Bold = True
    wsOutput.Range("A1").Font.Size = 14

    wsOutput.Range("A2").value = "コース名"
    wsOutput.Range("B2").value = "場所"
    wsOutput.Range("C2").value = "日程"
    wsOutput.Range("D2").value = "残席数"
    wsOutput.Range("A2:D2").Font.Bold = True
    wsOutput.Range("A2:D2").Interior.Color = RGB(85, 85, 85)
    wsOutput.Range("A2:D2").Font.Color = RGB(255, 255, 255)
    wsOutput.Range("A2:D2").HorizontalAlignment = xlCenter
    wsOutput.Range("A2:D2").VerticalAlignment = xlCenter

    outputRow = 3

    '=== ★ コースごとのデータを書き出し ===================================
    Dim hasOutput As Boolean
    hasOutput = False
    Dim skipBorderRows As Object
    Set skipBorderRows = CreateObject("Scripting.Dictionary")

    ' 出力: G3 が全てならカテゴリごとの見出しを付けて出力、そうでなければ従来どおり全コースを順に出力
    If isAllCategories Then
        Dim catKey As Variant
        For Each catKey In categoryMap.Keys
            ' カテゴリ間に1行の空白を入れる（先頭カテゴリの前は不要）
            If outputRow > 3 Then
                ' mark this blank row to skip borders and then increment
                skipBorderRows.Add CStr(outputRow), True
                outputRow = outputRow + 1
            End If
            ' カテゴリ見出し（14pt） - ただし最初のカテゴリは A1 に表示済みのため、
            ' A1 と同じカテゴリ名を3行目に重複して書かない（添付画像の要望）
            If Not (StrComp(CStr(catKey), selectedCategories(1), vbTextCompare) = 0) Then
                wsOutput.Cells(outputRow, "A").value = "〇" & catKey
                With wsOutput.Range(wsOutput.Cells(outputRow, "A"), wsOutput.Cells(outputRow, "D"))
                    .Font.Bold = True
                    .Font.Size = 14
                End With
                ' mark this category header row to skip borders
                skipBorderRows.Add CStr(outputRow), True
                outputRow = outputRow + 1
            End If
            ' 各コースを出力
            Dim cCol As Collection
            Set cCol = categoryMap(catKey)
            For i = 1 To cCol.Count
                courseName = cCol(i)
                If courseSchedules.Exists(courseName) Then
                    Set scheduleList = courseSchedules(courseName)
                    If scheduleList.Count > 0 Then
                        firstRowForCourse = outputRow
                        For scheduleIdx = 1 To scheduleList.Count
                            scheduleInfo = scheduleList(scheduleIdx)
                            If scheduleIdx = 1 Then
                                wsOutput.Cells(outputRow, "A").value = courseName
                            Else
                                wsOutput.Cells(outputRow, "A").value = vbNullString
                            End If
                            locationName = Trim$(CStr(scheduleInfo(0)))
                            wsOutput.Cells(outputRow, "B").value = locationName
                            If Len(locationName) > 0 Then
                                urlKey = LCase$(courseName) & "|" & LCase$(locationName)
                                If urlMap.Exists(urlKey) Then
                                    If wsOutput.Cells(outputRow, "B").Hyperlinks.Count > 0 Then
                                        wsOutput.Cells(outputRow, "B").Hyperlinks.Delete
                                    End If
                                    Dim baseUrl As String
                                    baseUrl = urlMap(urlKey)
                                    If Not urlAlphaMap.Exists(baseUrl) Then
                                        urlAlphaMap.Add baseUrl, currentAlpha
                                        currentAlpha = NextAlpha(currentAlpha)
                                    End If
                                    wsOutput.Hyperlinks.Add Anchor:=wsOutput.Cells(outputRow, "B"), _
                                        Address:=BuildTrackedUrl(baseUrl, dateParam, urlAlphaMap(baseUrl)), _
                                        TextToDisplay:=locationName
                                End If
                            End If
                            If IsDate(scheduleInfo(1)) Then
                                If IsDate(scheduleInfo(2)) Then
                                    wsOutput.Cells(outputRow, "C").value = _
                                        Format$(scheduleInfo(1), "m/d") & "～" & Format$(scheduleInfo(2), "m/d")
                                Else
                                    wsOutput.Cells(outputRow, "C").value = Format$(scheduleInfo(1), "m/d")
                                End If
                            Else
                                wsOutput.Cells(outputRow, "C").value = vbNullString
                            End If
                            ' 残席数 0 の場合は "満席" と表示
                            Dim seatValOut As String
                            If IsNumeric(scheduleInfo(3)) Then
                                If CLng(scheduleInfo(3)) = 0 Then
                                    seatValOut = "満席"
                                Else
                                    seatValOut = CStr(scheduleInfo(3))
                                End If
                            Else
                                seatValOut = CStr(scheduleInfo(3))
                            End If
                            wsOutput.Cells(outputRow, "D").value = seatValOut
                            currentRow = outputRow
                            With wsOutput.Range(wsOutput.Cells(currentRow, 1), wsOutput.Cells(currentRow, 4))
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                            End With
                            outputRow = outputRow + 1
                        Next scheduleIdx
                        ' マージ処理（コース名・場所）を従来と同様に行う
            If outputRow - firstRowForCourse > 1 Then
                            Set mergeRngA = wsOutput.Range(wsOutput.Cells(firstRowForCourse, "A"), _
                                    wsOutput.Cells(outputRow - 1, "A"))
                            SafeMerge mergeRngA
                            mergeRngA.VerticalAlignment = xlCenter
                            ' 同一場所の結合
                            segStart = firstRowForCourse
                            For bRow = firstRowForCourse + 1 To outputRow - 1
                                If Trim$(CStr(wsOutput.Cells(bRow, "B").value)) <> Trim$(CStr(wsOutput.Cells(bRow - 1, "B").value)) Then
                                    segEnd = bRow - 1
                                    If segEnd - segStart >= 1 Then
                                        Set mergeRngB = wsOutput.Range(wsOutput.Cells(segStart, "B"), wsOutput.Cells(segEnd, "B"))
                                        SafeMerge mergeRngB
                                        mergeRngB.HorizontalAlignment = xlCenter
                                        mergeRngB.VerticalAlignment = xlCenter
                                    End If
                                    segStart = bRow
                                End If
                            Next bRow
                            segEnd = outputRow - 1
                            If segEnd - segStart >= 1 Then
                                Set mergeRngB2 = wsOutput.Range(wsOutput.Cells(segStart, "B"), wsOutput.Cells(segEnd, "B"))
                                SafeMerge mergeRngB2
                                mergeRngB2.HorizontalAlignment = xlCenter
                                mergeRngB2.VerticalAlignment = xlCenter
                            End If
                        End If
                        ' 出力があったことを示すフラグ
                        hasOutput = True
                    End If
                End If
            Next i
        Next catKey
    Else
        For i = 1 To courseList.Count
            courseName = courseList(i)
            Set scheduleList = courseSchedules(courseName)
            
            If scheduleList.Count > 0 Then
            firstRowForCourse = outputRow

            For scheduleIdx = 1 To scheduleList.Count
                scheduleInfo = scheduleList(scheduleIdx)

                If scheduleIdx = 1 Then
                    wsOutput.Cells(outputRow, "A").value = courseName
                Else
                    wsOutput.Cells(outputRow, "A").value = vbNullString
                End If

                locationName = Trim$(CStr(scheduleInfo(0)))
                wsOutput.Cells(outputRow, "B").value = locationName

                ' URLが定義されている場合は、コース名×場所の組み合わせでハイパーリンクを追加
                If Len(locationName) > 0 Then
                    urlKey = LCase$(courseName) & "|" & LCase$(locationName)
                    If urlMap.Exists(urlKey) Then
                        If wsOutput.Cells(outputRow, "B").Hyperlinks.Count > 0 Then
                            wsOutput.Cells(outputRow, "B").Hyperlinks.Delete
                        End If
                        Dim baseUrl2 As String
                        baseUrl2 = urlMap(urlKey)
                        If Not urlAlphaMap.Exists(baseUrl2) Then
                            urlAlphaMap.Add baseUrl2, currentAlpha
                            currentAlpha = NextAlpha(currentAlpha)
                        End If
                        wsOutput.Hyperlinks.Add Anchor:=wsOutput.Cells(outputRow, "B"), _
                            Address:=BuildTrackedUrl(baseUrl2, dateParam, urlAlphaMap(baseUrl2)), _
                            TextToDisplay:=locationName
                    End If
                End If

                If IsDate(scheduleInfo(1)) Then
                    If IsDate(scheduleInfo(2)) Then
                        wsOutput.Cells(outputRow, "C").value = _
                            Format$(scheduleInfo(1), "m/d") & "～" & Format$(scheduleInfo(2), "m/d")
                    Else
                        wsOutput.Cells(outputRow, "C").value = Format$(scheduleInfo(1), "m/d")
                    End If
                Else
                    wsOutput.Cells(outputRow, "C").value = vbNullString
                End If

                wsOutput.Cells(outputRow, "D").value = scheduleInfo(3)

                currentRow = outputRow
                With wsOutput.Range(wsOutput.Cells(currentRow, 1), wsOutput.Cells(currentRow, 4))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                outputRow = outputRow + 1
            Next scheduleIdx

            ' 複数行に跨るコース名は見やすさのためにセル結合 (可読性向上)
            If outputRow - firstRowForCourse > 1 Then
                Set mergeRngA = wsOutput.Range(wsOutput.Cells(firstRowForCourse, "A"), _
                            wsOutput.Cells(outputRow - 1, "A"))
                SafeMerge mergeRngA
                mergeRngA.VerticalAlignment = xlCenter
            End If

            ' 同一コース内で連続する「場所」(列B)が同じであれば結合して見やすくする
            If outputRow - firstRowForCourse > 1 Then
                segStart = firstRowForCourse
                For bRow = firstRowForCourse + 1 To outputRow - 1
                    If Trim$(CStr(wsOutput.Cells(bRow, "B").value)) <> Trim$(CStr(wsOutput.Cells(bRow - 1, "B").value)) Then
                        segEnd = bRow - 1
                        If segEnd - segStart >= 1 Then
                            Set mergeRngB = wsOutput.Range(wsOutput.Cells(segStart, "B"), wsOutput.Cells(segEnd, "B"))
                            SafeMerge mergeRngB
                            mergeRngB.HorizontalAlignment = xlCenter
                            mergeRngB.VerticalAlignment = xlCenter
                        End If
                        segStart = bRow
                    End If
                Next bRow
                ' 最終セグメント処理
                segEnd = outputRow - 1
                If segEnd - segStart >= 1 Then
                    Set mergeRngB2 = wsOutput.Range(wsOutput.Cells(segStart, "B"), wsOutput.Cells(segEnd, "B"))
                    SafeMerge mergeRngB2
                    mergeRngB2.HorizontalAlignment = xlCenter
                    mergeRngB2.VerticalAlignment = xlCenter
                End If
            End If

            hasOutput = True
        End If
    Next i

    ' Close the If isAllCategories block
    End If

    If Not hasOutput Then
        wsOutput.Range("A3").value = "※指定した期間内のデータはありません。"
        wsOutput.Range("A3").Font.Color = RGB(192, 0, 0)
        Dim tmpMergeRng As Range
        Set tmpMergeRng = wsOutput.Range("A3:D3")
        SafeMerge tmpMergeRng
        tmpMergeRng.HorizontalAlignment = xlCenter
        tmpMergeRng.VerticalAlignment = xlCenter
    Else
        ' 罫線・列幅などの体裁調整
        Dim rIdx As Long
        For rIdx = 2 To outputRow - 1
            If Not skipBorderRows.Exists(CStr(rIdx)) Then
                With wsOutput.Range(wsOutput.Cells(rIdx, "A"), wsOutput.Cells(rIdx, "D"))
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
            Else
                ' クリア: カテゴリ見出しや空行は罫線を消す
                With wsOutput.Range(wsOutput.Cells(rIdx, "A"), wsOutput.Cells(rIdx, "D"))
                    .Borders.LineStyle = xlNone
                End With
            End If
        Next rIdx

        ' 列幅自動調整（最後に実行）
        wsOutput.Columns("A:D").EntireColumn.AutoFit

        ' コース名の列は幅に余裕を持たせる
        wsOutput.Columns("A").ColumnWidth = WorksheetFunction.Max(20, wsOutput.Columns("A").ColumnWidth)
        wsOutput.Columns("B").ColumnWidth = WorksheetFunction.Max(12, wsOutput.Columns("B").ColumnWidth)
        wsOutput.Columns("C").ColumnWidth = WorksheetFunction.Max(14, wsOutput.Columns("C").ColumnWidth)
        wsOutput.Columns("D").ColumnWidth = WorksheetFunction.Max(8, wsOutput.Columns("D").ColumnWidth)
    End If

CleanExit:
    '=== ★ 後片付け (参照の解放とアプリ状態の復元) =========================
    On Error Resume Next
    Set wsDetail = Nothing
    Set wsCategory = Nothing
    Set wsUrl = Nothing
    Set wsOutput = Nothing
    Set wbDetail = Nothing
    Set wbCategory = Nothing
    Set wbUrl = Nothing
    On Error GoTo 0

    Application.ScreenUpdating = screenUpdatingState
    Application.Calculation = calcState
    Application.EnableEvents = enableEventsState

    ' 終了時に出力シートをアクティブにして A1 を選択（保護やエラーを考慮）
    On Error Resume Next
    If Not ThisWorkbook Is Nothing Then
        With ThisWorkbook
            If .Worksheets.Count >= 1 Then
                .Worksheets(1).Activate
                .Worksheets(1).Range("A1").Select
            End If
        End With
    End If
    On Error GoTo 0

    Exit Sub

ErrHandler:
    MsgBox "スケジュール生成中にエラーが発生しました。" & vbCrLf & _
           "詳細: " & Err.Description, vbCritical + vbOKOnly, "GenerateCategorySchedule"
    Resume CleanExit

End Sub

Private Function BuildTrackedUrl(ByVal baseUrl As String, ByVal dateParam As String, ByVal alphaCode As String, Optional ByVal encodeAmpersands As Boolean = False) As String
    Dim separator As String
    Dim query As String
    Dim finalUrl As String

    If Len(baseUrl) = 0 Then
        BuildTrackedUrl = baseUrl
        Exit Function
    End If

    query = "utm_source=mail&utm_medium=email&utm_campaign=mail" & dateParam & "edu_" & alphaCode

    If InStr(baseUrl, "?") > 0 Then
        If Right$(baseUrl, 1) = "?" Or Right$(baseUrl, 1) = "&" Then
            separator = ""
        Else
            separator = "&"
        End If
    Else
        separator = "?"
    End If

    finalUrl = baseUrl & separator & query

    If encodeAmpersands Then
        finalUrl = Replace$(finalUrl, "&", "&amp;")
    End If

    BuildTrackedUrl = finalUrl
End Function

Private Function NextAlpha(ByVal alpha As String) As String
    Dim i As Long
    Dim carry As Boolean
    Dim chars() As String
    ReDim chars(1 To Len(alpha))
    For i = 1 To Len(alpha)
        chars(i) = Mid(alpha, i, 1)
    Next i
    carry = True
    For i = Len(alpha) To 1 Step -1
        If carry Then
            If chars(i) = "z" Then
                chars(i) = "a"
            Else
                chars(i) = Chr(Asc(chars(i)) + 1)
                carry = False
            End If
        End If
    Next i
    If carry Then
        NextAlpha = "a" & Join(chars, "")
    Else
        NextAlpha = Join(chars, "")
    End If
End Function







