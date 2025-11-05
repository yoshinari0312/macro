Option Explicit
'
'=== ファイルパスなどの基本設定を定義 ===
' ここを環境に合わせて変更してください
Private Const BASE_PATH As String = "C:\Users\intern-yo-onodera\OneDrive - DMG MORI\テクニウム 教育サービスGP - 0_対面型スクール日程表作成"
'Private Const BASE_PATH As String = "C:\Users\intern-yk-shimizu\OneDrive - DMG MORI\メルマガ\田中李空＆遥馨さん＆清水さん共有\01_メルマガ資料\02_メルマガ詳細資料\0_対面型スクール日程表作成"

Private Const DEBUG_MODE As Boolean = False ' Trueにするとデバッグ用MsgBoxを表示

Private Sub DebugMessage(ByVal message As String)
    ' デバッグモードが有効な場合のみ情報をメッセージボックスで表示
    If DEBUG_MODE Then
        MsgBox message, vbInformation, "GenerateCategorySchedule - Debug"
    End If
End Sub

' 指定の形式でスケジュールシートをHTMLに変換して保存/貼付する
Public Sub ExportScheduleToHTML()
    ' スケジュールシートを走査してメール配信用のHTMLテーブルを生成・保存するメイン処理
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
    ' 出力対象となる「スケジュール」シートを取得

    ' URLパラメータの初期化
    dateParam = Format(ws.Range("G14").value, "yymmdd")
    ' HTML内リンクのパラメータ生成に使用する日付と連番の初期化
    startAlpha = LCase(Trim(ws.Range("G15").value))
    currentAlpha = startAlpha
    Set urlAlphaMap = CreateObject("Scripting.Dictionary")
    ' URLごとに付与する英字パラメータを重複なく管理

    Dim lastA As Long, lastB As Long, lastC As Long, lastD As Long
    ' A～D列の最終行を取得して実際のデータ範囲を把握
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
    ' 日程に含まれる月を収集してHTML冒頭のコメントに出力する
    For r = 3 To lastRow
        ' 日程列から月情報を抽出
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
            ' 月番号を昇順に並べ替え
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
        ' 収集した月番号をカンマ区切りで連結しコメント文を作成
        For j = 0 To UBound(mList)
            If monthsLabel <> "" Then monthsLabel = monthsLabel & "、"
            monthsLabel = monthsLabel & mList(j)
        Next j
        monthsLabel = monthsLabel & "月講座一覧 表"
    End If

    html = "<!-- " & monthsLabel & " -->" & vbCrLf & vbCrLf & "<table>" & vbCrLf
    ' 月情報のコメントとテーブル開始タグを作成

    ' 最初のカテゴリヘッダを先頭に出す場合、スケジュールシートの先頭行にカテゴリがあることを期待して
    ' ループ前に先頭のカテゴリ行を探して出力しておく（ユーザの要求: 一番上に最初のカテゴリの表示）
    Dim firstCatFound As Boolean
    firstCatFound = False
    Dim firstCatName As String
    For r = 1 To lastRow
        ' 最上段に表示する最初のカテゴリ行を探してHTMLへ出力
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
        ' スケジュール全体を先頭から順に処理
        advance = 1

        cellA = Trim$(CStr(ws.Cells(rr, "A").value))
        cellB = Trim$(CStr(ws.Cells(rr, "B").value))
        cellC = Trim$(CStr(ws.Cells(rr, "C").value))
        cellD = Trim$(CStr(ws.Cells(rr, "D").value))

        If Len(cellA) > 0 And Left$(cellA, 1) = "〇" And cellB = "" And cellC = "" And cellD = "" Then
            ' 新しいカテゴリ見出しが出現した場合の処理
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
                ' カテゴリに紐づくデータが続く場合のみテーブル要素を作成
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
            ' 完全な空行はスキップ
            GoTo AdvanceLoop
        End If

        If cellA = "" Then
            ' コース名が空の行もスキップ
            GoTo AdvanceLoop
        End If

        If Not openTable Then
            ' テーブルが開かれていない場合は新規にテーブルを生成
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
            ' 同一コースに属する連続行を判定して最終行を求める
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
            ' コースに紐づく場所・リンク・日程・残席情報を配列へ格納
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
            ' 同一場所が連続する行数を計算して結合用情報を作成
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
            ' HTMLの1行分を出力（コース名・場所・日程・残席）
            If idx = 1 Then
                html = html & "<tr style=""background-color: #ffffff;"">" & vbCrLf
                html = html & "<td width=""50%"" rowspan=""" & rowCount & """ style=""height:12px!important;font-size:12px!important;line-height:1!important;padding:6px 8px!important;border:1px solid #000!important;vertical-align: middle; text-align: center;"">" & escapedCourseName & "</td>" & vbCrLf
            Else
                html = html & "<tr>" & vbCrLf
            End If

            If placeSpan(idx) > 0 Then
                ' 同一場所が続く場合はrowspanでまとめて表示
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
        ' まとめて処理した行数分だけループカウンタを進める

AdvanceLoop:
        ' 次の行に進むためのラベル
        rr = rr + advance
    Loop

    If openTable Then
        ' ループ終了時に開きっぱなしのテーブルがあれば閉じる
        html = html & "        </tbody>" & vbCrLf & _
                      "    </table>" & vbCrLf & _
                      "</td>" & vbCrLf & vbCrLf
        openTable = False
    End If

    html = html & "</table>" & vbCrLf

    timeStamp = Format$(Now, "yyyymmdd_Hhnnss")
    ' 出力ファイル名にタイムスタンプを付与し保存
    sFile = BASE_PATH & "\schedule_export_" & timeStamp & ".html"
    fnum = FreeFile
    Open sFile For Output As #fnum
    ' HTML文字列をファイルへ書き込み
    Print #fnum, html
    Close #fnum

    MsgBox "HTML を生成し、ファイルに保存しました: " & sFile, vbInformation
    ' 出力完了をユーザへ通知
    Exit Sub

ErrHandler:
    ' エラー発生時はメッセージを表示して処理を終了
    MsgBox "HTML エクスポート中にエラーが発生しました: " & Err.Description, vbExclamation
End Sub

' セル結合時の警告を抑制して安全に結合するヘルパー
Private Sub SafeMerge(ByVal rng As Range)
    ' セル結合時の警告を回避しつつ上段セルの値を保持するユーティリティ
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
    ' 日付文字列に混在する全角記号や改行などを統一的な形式に整形
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
    ' 日付セルと補助情報から正しい日付を推定し、解析に成功したらTrueを返す
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
        ' 区切りがない場合は日付シリアル値の可能性が高いため元のセル値を優先
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
        ' 指定された年を優先して月日情報と結合する
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
        ' フォールバック年がない場合は元の値内の年情報を尊重
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
        ' 解析できなかった場合でもセルが日付ならそれを採用
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
    ' 文字列内に4桁の西暦が含まれているかどうかを判定
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

    ' 指定カテゴリと期間に合致する講座情報を抽出し、スケジュールシートへ整形出力するメイン処理

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
    ' チェックボックスで選択されたカテゴリ名をまとめるコレクション
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
        ' TRUEになっているセルごとにカテゴリ名を収集
        If cell.value = True Then
            selectedCategories.Add categoryMappings(cell.Address(RowAbsolute:=False, ColumnAbsolute:=False))
        End If
    Next cell
    
    If selectedCategories.Count = 0 Then
        ' 一つも選択されていなければエラーで処理を中断
        Err.Raise vbObjectError + 100, "GenerateCategorySchedule", "F4～F10のチェックボックスで少なくとも1つのカテゴリを選択してください。"
    End If

    isAllCategories = (selectedCategories.Count > 1)
    ' 2カテゴリ以上選択されている場合はカテゴリ別出力モード

    ' URLパラメータの初期化
    dateParam = Format(wsOutput.Range("G14").value, "yymmdd")
    startAlpha = LCase(Trim(wsOutput.Range("G15").value))
    currentAlpha = startAlpha
    Set urlAlphaMap = CreateObject("Scripting.Dictionary")

    If Not TryParseDate(wsOutput.Range("G12").value, filterStartDate) Then
        ' 開始日が不正な場合はエラー
        Err.Raise vbObjectError + 101, "GenerateCategorySchedule", "G12セルの開始日が日付として認識できません。"
    End If

    If Not TryParseDate(wsOutput.Range("I12").value, filterEndDate) Then
        ' 終了日が不正な場合はエラー
        Err.Raise vbObjectError + 102, "GenerateCategorySchedule", "I12セルの終了日が日付として認識できません。"
    End If

    If filterEndDate < filterStartDate Then
        ' 期間の前後が逆転していればエラー
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
    ' コース詳細の元データブックを読み取り専用で開く
    detailOpened = True
    Set wsDetail = wbDetail.Worksheets("空席情報2025_26年1Q")

    Set wbCategory = Workbooks.Open(BASE_PATH & "\" & FILE_EXCEL2, ReadOnly:=True)
    ' カテゴリ対応表を開く
    categoryOpened = True
    Set wsCategory = wbCategory.Worksheets("カテゴリ")

    Set urlMap = CreateObject("Scripting.Dictionary")
    Set wbUrl = Workbooks.Open(BASE_PATH & "\" & FILE_URL_MAP, ReadOnly:=True)
    ' コースとURLの対応表を開き、辞書に格納
    urlOpened = True
    Set wsUrl = wbUrl.Worksheets("コースとURL対応表")

    lastRowUrl = wsUrl.Cells(wsUrl.Rows.Count, "E").End(xlUp).Row
    For rowIdx = 1 To lastRowUrl
        ' コース名×場所の組み合わせでURLを辞書に保存
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
        ' 選択されたカテゴリごとに対応するコース名を収集
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
        ' 該当コースが一つも見つからない場合はエラー
        Err.Raise vbObjectError + 106, "GenerateCategorySchedule", _
                  "選択されたカテゴリに紐づくコースがエクセル2.xlsxで見つかりませんでした。"
    End If    '=== ★ コースごとのスケジュール格納用ディクショナリ ==================
    Set courseSchedules = CreateObject("Scripting.Dictionary")
    ' コース名ごとに開催スケジュールを格納する辞書（値はCollection）
    For i = 1 To courseList.Count
        If Not courseSchedules.Exists(courseList(i)) Then
            Set scheduleList = New Collection
            courseSchedules.Add courseList(i), scheduleList
        End If
    Next i

    ' G3 が全て指定のときはカテゴリごとに出力するためのマップを作成
    Dim categoryMap As Object ' Key=カテゴリ名, Item=Collection of course names
    Set categoryMap = CreateObject("Scripting.Dictionary")
    ' カテゴリごとに紐づくコース一覧を保持し、複数カテゴリ選択時のグルーピングに利用
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
        ' エクセル1.xlsxの明細行を走査して該当スケジュールを抽出
        courseName = Trim$(wsDetail.Cells(rowIdx, "C").value)
        Dim rowStatus As String
        Dim parsedStartText As String
        Dim rawStart As String
        rawStart = Trim$(CStr(wsDetail.Cells(rowIdx, "G").value))
        If courseSchedules.Exists(courseName) Then
            If TryParseDate(wsDetail.Cells(rowIdx, "G").value, tmpDate, wsDetail.Cells(rowIdx, "A").value) Then
                ' 開催開始日の解析に成功した場合のみ判定を行う
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
                ' 日付が解析できなかった行は統計用にカウント
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
        ' 既存の一覧表示領域をクリアしてから再生成
        .Clear
        .ClearFormats
    End With

    '=== ★ カテゴリタイトルとヘッダー行を整形 ============================
    ' ヘッダ表示: 選択された最初のカテゴリ名をA1に表示
    If selectedCategories.Count > 0 Then
    wsOutput.Range("A1").value = "〇" & selectedCategories(1)
    ' 一番目に選択されたカテゴリ名を大見出しとして表示
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
    ' 罫線適用を除外する行番号を管理（カテゴリ見出しや空行など）

    ' 出力: G3 が全てならカテゴリごとの見出しを付けて出力、そうでなければ従来どおり全コースを順に出力
    If isAllCategories Then
        Dim catKey As Variant
        For Each catKey In categoryMap.Keys
            ' 各カテゴリごとに見出しとコース情報を出力
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
                ' カテゴリに紐づく各コースのスケジュールを展開
                courseName = cCol(i)
                If courseSchedules.Exists(courseName) Then
                    Set scheduleList = courseSchedules(courseName)
                    If scheduleList.Count > 0 Then
                        firstRowForCourse = outputRow
                        For scheduleIdx = 1 To scheduleList.Count
                            ' 各開催日程を1行ずつ書き出し
                            scheduleInfo = scheduleList(scheduleIdx)
                            If scheduleIdx = 1 Then
                                wsOutput.Cells(outputRow, "A").value = courseName
                            Else
                                wsOutput.Cells(outputRow, "A").value = vbNullString
                            End If
                            locationName = Trim$(CStr(scheduleInfo(0)))
                            wsOutput.Cells(outputRow, "B").value = locationName
                            If Len(locationName) > 0 Then
                                ' 場所が空でなければURL辞書からリンクを付与
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
                            ' 残席数0は「満席」表示に差し替え
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
                            ' コース名が複数行に渡る場合はA列とB列を適切に結合
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
            ' 単一カテゴリ指定時はコース一覧をそのまま出力
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

                ' 残席数は元データをそのまま表示（数値以外の表現にも対応）
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
        ' 条件に合致するデータが無い場合のメッセージ表示
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
            ' 罫線をカテゴリ行とデータ行で切り替え
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
    ' 処理中に例外が発生した場合は内容をユーザへ通知
    MsgBox "スケジュール生成中にエラーが発生しました。" & vbCrLf & _
           "詳細: " & Err.Description, vbCritical + vbOKOnly, "GenerateCategorySchedule"
    Resume CleanExit

End Sub

Private Function StripTrackingParameters(ByVal url As String) As String
    ' 既存URLからメルマガ用のUTMパラメータを除去してベースURLを得る
    Dim workUrl As String
    Dim fragment As String
    Dim hashPos As Long
    Dim qPos As Long
    Dim params As Variant
    Dim param As Variant
    Dim filtered As String
    Dim lowerParam As String

    workUrl = url
    hashPos = InStr(workUrl, "#")
    If hashPos > 0 Then
        ' アンカー(#以降)は一旦切り離して後で付け直す
        fragment = Mid$(workUrl, hashPos)
        workUrl = Left$(workUrl, hashPos - 1)
    End If

    qPos = InStr(workUrl, "?")
    If qPos = 0 Then
        ' クエリ文字列が存在しない場合はそのまま返却
        StripTrackingParameters = workUrl & fragment
        Exit Function
    End If

    params = Split(Mid$(workUrl, qPos + 1), "&")
    ' 既存のクエリパラメータからUTM系のみ除外して再構築
    For Each param In params
        lowerParam = LCase$(Trim$(param))
        If Len(lowerParam) > 0 Then
            If Not (Left$(lowerParam, Len("utm_source=mail")) = "utm_source=mail" _
                Or Left$(lowerParam, Len("utm_medium=email")) = "utm_medium=email" _
                Or Left$(lowerParam, Len("utm_campaign=mail")) = "utm_campaign=mail") Then
                If Len(filtered) = 0 Then
                    filtered = param
                Else
                    filtered = filtered & "&" & param
                End If
            End If
        End If
    Next param

    If Len(filtered) > 0 Then
        StripTrackingParameters = Left$(workUrl, qPos - 1) & "?" & filtered & fragment
    Else
        StripTrackingParameters = Left$(workUrl, qPos - 1) & fragment
    End If
End Function

Private Function BuildTrackedUrl(ByVal baseUrl As String, ByVal dateParam As String, ByVal alphaCode As String, Optional ByVal encodeAmpersands As Boolean = False) As String
    ' 追跡用UTMパラメータを付与したURLを生成（必要に応じて&を&amp;に変換）
    Dim separator As String
    Dim query As String
    Dim finalUrl As String
    Dim workUrl As String
    Dim lowerWorkUrl As String
    Dim targetCampaign As String
    Dim hasSource As Boolean
    Dim hasMedium As Boolean
    Dim hasCampaign As Boolean

    If Len(baseUrl) = 0 Then
        ' URLが空の場合はそのまま返す
        BuildTrackedUrl = baseUrl
        Exit Function
    End If

    If encodeAmpersands Then
        ' HTML出力時は事前に&を戻してから処理する
        workUrl = Replace$(baseUrl, "&amp;", "&")
    Else
        workUrl = baseUrl
    End If

    query = "utm_source=mail&utm_medium=email&utm_campaign=mail" & dateParam & "edu_" & alphaCode
    ' 付与する追跡パラメータを生成
    lowerWorkUrl = LCase$(workUrl)
    targetCampaign = LCase$("utm_campaign=mail" & dateParam & "edu_" & alphaCode)
    hasSource = (InStr(lowerWorkUrl, "utm_source=mail") > 0)
    hasMedium = (InStr(lowerWorkUrl, "utm_medium=email") > 0)
    hasCampaign = (InStr(lowerWorkUrl, targetCampaign) > 0)
    ' 既存URLに同じUTMパラメータが含まれているかを判定

    If hasSource And hasMedium And hasCampaign Then
        ' 既に同じパラメータが揃っている場合は再付与しない
        If encodeAmpersands Then
            finalUrl = workUrl
        Else
            finalUrl = baseUrl
        End If
    Else
        If hasSource Or hasMedium Or hasCampaign Then
            ' 追跡情報が部分的に含まれている場合はいったん削除
            workUrl = StripTrackingParameters(workUrl)
        End If

        If InStr(workUrl, "?") > 0 Then
            ' 既存のクエリ有無に応じて接続文字を切り替える
            If Right$(workUrl, 1) = "?" Or Right$(workUrl, 1) = "&" Then
                separator = ""
            Else
                separator = "&"
            End If
        Else
            separator = "?"
        End If

        finalUrl = workUrl & separator & query
    End If

    If encodeAmpersands Then
        ' HTML用にアンパサンドをエンコード
        finalUrl = Replace$(finalUrl, "&", "&amp;")
    End If

    BuildTrackedUrl = finalUrl
End Function

Private Function NextAlpha(ByVal alpha As String) As String
    ' メール追跡用パラメータの英字連番を算出（zの次はaa）
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







