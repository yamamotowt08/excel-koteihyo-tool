Attribute VB_Name = "GanttMaker"
'======================================================================
'  工程表 自動作成マクロ
'
'  「入力」シートに作業項目と開始日・終了日・日数を入れて
'  MakeSchedule を実行すると、「工程表」シートに
'  月を上旬・中旬・下旬に区切った工程表と塗りつぶしバーを作成します。
'
'  ・開始日＋終了日 → 日数を計算
'  ・開始日＋日数   → 終了日を計算
'  ・終了日＋日数   → 開始日を計算
'  ・日数だけ       → 1つ上の期間の翌日から開始
'  ・開始日だけ     → 1日だけの作業
'  ・作業項目が空欄で日付がある行 → 上の作業の追加期間（同じ行にバーを追加）
'
'  日付が変わったときは「変更後」の欄に入れると、
'  その作業の行が上下2段になり、上段に変更前、下段に変更後のバーを描きます。
'  ・変更後 開始日だけ → 日数は変更前と同じでずらす
'  ・変更後 終了日だけ → 開始日は変更前のまま（延長・短縮）
'  ・変更後 日数だけ   → 開始日は変更前のまま
'======================================================================
Option Explicit

Private Const IN_SHEET As String = "入力"
Private Const OUT_SHEET As String = "工程表"
Private Const HOL_SHEET As String = "祝日"
Private Const OUT_FONT As String = "ＭＳ 明朝"
Private Const WORKDAY_MODE As String = "土日祝除く"
Private Const DATE_FMT As String = "[$-411]ggge""年""m""月""d""日"""
Private Const MONTH_FMT As String = "[$-411]e""年""m""月"""

' 入力シートの列
Private Const IC_NAME As Long = 1     ' 作業項目
Private Const IC_START As Long = 2    ' 開始日
Private Const IC_END As Long = 3      ' 終了日
Private Const IC_DAYS As Long = 4     ' 日数
Private Const IC_CSTART As Long = 5   ' 変更後 開始日
Private Const IC_CEND As Long = 6     ' 変更後 終了日
Private Const IC_CDAYS As Long = 7    ' 変更後 日数
Private Const IC_COLOR As Long = 8    ' バー色
Private Const IC_NOTE As Long = 9     ' 備考
Private Const IC_RES As Long = 11     ' 計算結果（開始日・終了日・日数 ×変更前／変更後の6列）

' 工程表シートの列・行
Private Const OC_NAME As Long = 1
Private Const OC_DAYS As Long = 2
Private Const OC_CHART As Long = 3
Private Const OR_LEGEND As Long = 14
Private Const OR_PROJECT As Long = 15
Private Const OR_PERIOD As Long = 16
Private Const OR_MONTH As Long = 17
Private Const OR_TEN As Long = 18
Private Const OR_FIRST As Long = 19
Private Const ROW_H As Double = 18     ' 通常の行の高さ
Private Const ROW_H2 As Double = 28    ' 変更がある行（上下2段）の高さ

' バーの描き方
Private Const STYLE_FILL As Long = 0
Private Const STYLE_DASH As Long = 1
Private Const STYLE_OUTLINE As Long = 2

' Office の定数（参照設定なしで動くよう数値で書く）
Private Const SHAPE_RECT As Long = 1
Private Const MSO_TRUE As Long = -1
Private Const MSO_FALSE As Long = 0
Private Const LINE_DASH As Long = 4      ' 破線（msoLineDash）

' 作業（工程表の1行）
Private itemName() As String
Private itemNote() As String
Private itemDays() As Long        ' 変更前の日数合計
Private itemCDays() As Long       ' 変更後の日数合計
Private itemColor() As Long
Private itemHasChg() As Boolean
Private nItems As Long

' 期間（バー1本）
Private perItem() As Long
Private perHasOrig() As Boolean
Private perStart() As Date
Private perEnd() As Date
Private perHasChg() As Boolean
Private perCStart() As Date
Private perCEnd() As Date
Private perColor() As Long
Private nPers As Long

Private useWorkDays As Boolean
Private holRange As Range
Private monthBase As Date
Private chgColor As Long
Private oldStyle As Long

'----------------------------------------------------------------------
' 工程表を作成（ボタンから呼ばれるメイン処理）
'----------------------------------------------------------------------
Public Sub MakeSchedule()
    Dim wsIn As Worksheet, ws As Worksheet

    On Error GoTo EH
    Set wsIn = ThisWorkbook.Worksheets(IN_SHEET)

    Application.ScreenUpdating = False
    ReadInput wsIn
    If nPers = 0 Then
        Application.ScreenUpdating = True
        MsgBox "開始日・終了日・日数が入力された作業がありません。", vbExclamation, "工程表の作成"
        Exit Sub
    End If

    Set ws = GetOutSheet()
    DrawSchedule ws
    Application.ScreenUpdating = True

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "工程表の作成"
End Sub

'----------------------------------------------------------------------
' 入力シートを読んで作業と期間を集める
'----------------------------------------------------------------------
Private Sub ReadInput(ByVal wsIn As Worksheet)
    Dim firstRow As Long, lastRow As Long, r As Long, c As Long, n As Long
    Dim nm As String, note As String, defColor As Long, rowColor As Long
    Dim s As Date, e As Date, d As Long, lastEnd As Date, haveLast As Boolean
    Dim hasS As Boolean, hasE As Boolean, hasD As Boolean, hasOrig As Boolean
    Dim cs As Date, ce As Date, cd As Long
    Dim hasCS As Boolean, hasCE As Boolean, hasCD As Boolean, hasChg As Boolean

    useWorkDays = (CfgText("cfg_DayMode") = WORKDAY_MODE)
    LoadHolidays
    defColor = CellColor(CfgRange("cfg_BarColor"), RGB(0, 0, 0))
    chgColor = CellColor(CfgRange("cfg_ChangeColor"), RGB(192, 0, 0))
    Select Case CfgText("cfg_OldStyle")
        Case "点線": oldStyle = STYLE_DASH
        Case "白抜き": oldStyle = STYLE_OUTLINE
        Case Else: oldStyle = STYLE_FILL
    End Select

    If CfgRange("tbl_Header") Is Nothing Then
        Err.Raise vbObjectError + 1, , "入力シートの作業表（名前「tbl_Header」）が見つかりません。"
    End If
    firstRow = CfgRange("tbl_Header").Row + 1
    lastRow = firstRow - 1
    For c = IC_NAME To IC_CDAYS
        r = wsIn.Cells(wsIn.Rows.Count, c).End(xlUp).Row
        If r > lastRow Then lastRow = r
    Next c
    wsIn.Range(wsIn.Cells(firstRow, IC_RES), wsIn.Cells(lastRow + 200, IC_RES + 5)).ClearContents

    n = lastRow - firstRow + 2
    ReDim itemName(1 To n): ReDim itemNote(1 To n)
    ReDim itemDays(1 To n): ReDim itemCDays(1 To n)
    ReDim itemColor(1 To n): ReDim itemHasChg(1 To n)
    ReDim perItem(1 To n): ReDim perHasOrig(1 To n)
    ReDim perStart(1 To n): ReDim perEnd(1 To n)
    ReDim perHasChg(1 To n): ReDim perCStart(1 To n): ReDim perCEnd(1 To n)
    ReDim perColor(1 To n)
    nItems = 0: nPers = 0: haveLast = False

    For r = firstRow To lastRow
        nm = CellText(wsIn.Cells(r, IC_NAME))
        hasS = ReadDate(wsIn.Cells(r, IC_START), "開始日", s)
        hasE = ReadDate(wsIn.Cells(r, IC_END), "終了日", e)
        hasD = ReadDays(wsIn.Cells(r, IC_DAYS), "日数", d)
        hasCS = ReadDate(wsIn.Cells(r, IC_CSTART), "変更後の開始日", cs)
        hasCE = ReadDate(wsIn.Cells(r, IC_CEND), "変更後の終了日", ce)
        hasCD = ReadDays(wsIn.Cells(r, IC_CDAYS), "変更後の日数", cd)
        hasOrig = hasS Or hasE Or hasD
        hasChg = hasCS Or hasCE Or hasCD

        If nm <> "" Then
            ' 新しい作業
            nItems = nItems + 1
            itemName(nItems) = nm
            itemNote(nItems) = CellText(wsIn.Cells(r, IC_NOTE))
            itemColor(nItems) = CellColor(wsIn.Cells(r, IC_COLOR), defColor)
            rowColor = itemColor(nItems)
        ElseIf hasOrig Or hasChg Then
            ' 作業項目が空欄 → 上の作業の追加期間
            If nItems = 0 Then RowError r, "作業項目が空欄です。"
            rowColor = CellColor(wsIn.Cells(r, IC_COLOR), itemColor(nItems))
            note = CellText(wsIn.Cells(r, IC_NOTE))
            If note <> "" Then
                If itemNote(nItems) <> "" Then itemNote(nItems) = itemNote(nItems) & "、"
                itemNote(nItems) = itemNote(nItems) & note
            End If
        End If

        ' ---- 変更前（当初）の期間 ----
        If hasOrig Then
            If hasS And hasE Then
                If e < s Then RowError r, "終了日が開始日より前になっています。"
            ElseIf hasS And hasD Then
                e = AddDays(s, d)
            ElseIf hasE And hasD Then
                s = SubDays(e, d)
            ElseIf hasD Then
                If Not haveLast Then RowError r, "開始日を入力してください（日数だけの行は1つ上の期間の翌日から始まります）。"
                s = NextStart(lastEnd)
                e = AddDays(s, d)
            ElseIf hasS Then
                e = s
            Else
                s = e
            End If
            d = CountDays(s, e)
            lastEnd = e
            haveLast = True
        End If

        ' ---- 変更後の期間（足りない所は変更前から補う） ----
        If hasChg Then
            If hasCS And hasCE Then
                ' そのまま
            ElseIf hasCS And hasCD Then
                ce = AddDays(cs, cd)
            ElseIf hasCE And hasCD Then
                cs = SubDays(ce, cd)
            ElseIf Not hasOrig Then
                If hasCD Then RowError r, "変更後の開始日か終了日を入力してください。"
                If hasCS Then ce = cs Else cs = ce
            ElseIf hasCS Then
                ce = AddDays(cs, IIf(d < 1, 1, d))      ' 日数はそのままでずらす
            ElseIf hasCE Then
                cs = s                                  ' 開始日はそのまま
            Else
                cs = s
                ce = AddDays(cs, cd)
            End If
            If ce < cs Then RowError r, "変更後の終了日が開始日より前になっています。"
            cd = CountDays(cs, ce)
        End If

        If hasOrig Or hasChg Then
            nPers = nPers + 1
            perItem(nPers) = nItems
            perHasOrig(nPers) = hasOrig
            perStart(nPers) = s
            perEnd(nPers) = e
            perHasChg(nPers) = hasChg
            perCStart(nPers) = cs
            perCEnd(nPers) = ce
            perColor(nPers) = rowColor
            If hasOrig Then
                itemDays(nItems) = itemDays(nItems) + d
                wsIn.Cells(r, IC_RES).Value = s
                wsIn.Cells(r, IC_RES + 1).Value = e
                wsIn.Cells(r, IC_RES + 2).Value = d
            End If
            If hasChg Then
                itemHasChg(nItems) = True
                itemCDays(nItems) = itemCDays(nItems) + cd
                wsIn.Cells(r, IC_RES + 3).Value = cs
                wsIn.Cells(r, IC_RES + 4).Value = ce
                wsIn.Cells(r, IC_RES + 5).Value = cd
            ElseIf hasOrig Then
                itemCDays(nItems) = itemCDays(nItems) + d
            End If
        End If
    Next r
End Sub

'----------------------------------------------------------------------
' 工程表シートを描く
'----------------------------------------------------------------------
Private Sub DrawSchedule(ByVal ws As Worksheet)
    Dim i As Long, k As Long, c As Long, r As Long, it As Long
    Dim minS As Date, maxE As Date, curS As Date, curE As Date, first As Boolean
    Dim nMonths As Long, lastCol As Long, lastRow As Long, colW As Double
    Dim c1 As Long, c2 As Long, rg As Range, shp As Shape, anyChg As Boolean
    Dim barH As Double, pct As Double, rowTop As Double, rowHt As Double
    Dim s As Date, e As Date, dTop As String, dBot As String

    ' 表示する範囲（全バー）と、開始年月日・終了予定年月日（変更後の予定）
    first = True
    For i = 1 To nPers
        If perHasOrig(i) Then
            If first Or perStart(i) < minS Then minS = perStart(i)
            If first Or perEnd(i) > maxE Then maxE = perEnd(i)
            first = False
        End If
        If perHasChg(i) Then
            If first Or perCStart(i) < minS Then minS = perCStart(i)
            If first Or perCEnd(i) > maxE Then maxE = perCEnd(i)
            first = False
            s = perCStart(i): e = perCEnd(i)
        Else
            s = perStart(i): e = perEnd(i)
        End If
        If i = 1 Or s < curS Then curS = s
        If i = 1 Or e > curE Then curE = e
    Next i
    For i = 1 To nItems
        If itemHasChg(i) Then anyChg = True
    Next i

    monthBase = DateSerial(Year(minS), Month(minS), 1)
    nMonths = (Year(maxE) - Year(minS)) * 12 + Month(maxE) - Month(minS) + 1
    lastCol = OC_CHART + nMonths * 3          ' 備考の列
    lastRow = OR_FIRST + nItems - 1

    ClearSheet ws

    ' ---- 列幅 ----
    ws.Columns(OC_NAME).ColumnWidth = 40
    ws.Columns(OC_DAYS).ColumnWidth = 6
    colW = 72 / (nMonths * 3)
    If colW < 2.63 Then colW = 2.63
    If colW > 8 Then colW = 8
    ws.Range(ws.Columns(OC_CHART), ws.Columns(lastCol - 1)).ColumnWidth = colW
    ws.Columns(lastCol).ColumnWidth = 22

    ' ---- 表の上（様式番号・表題・作成日・宛先・作成者・前文） ----
    ws.Cells(1, 1).Value = CfgText("cfg_FormNo")

    With MergeRow(ws, 3, OC_NAME, lastCol)
        .Cells(1, 1).Value = CfgText("cfg_Title")
        .HorizontalAlignment = xlCenter
        .Font.Size = 16
    End With
    ws.Rows(3).RowHeight = 26

    With MergeRow(ws, 4, SpanStart(ws, lastCol, 20), lastCol)
        PutDateOrText .Cells(1, 1), CfgValue("cfg_Date")
        .HorizontalAlignment = xlRight
    End With

    ws.Cells(6, 1).Value = CfgText("cfg_To")
    ws.Cells(6, 1).Font.Size = 11

    c1 = SpanStart(ws, lastCol, 50)
    WriteAuthor ws, 7, c1, lastCol, "作成者　所在地　", CfgText("cfg_Address")
    WriteAuthor ws, 8, c1, lastCol, "　　　　会社名　", CfgText("cfg_Company")
    WriteAuthor ws, 9, c1, lastCol, "　　　　名　前　", CfgText("cfg_Name")
    WriteAuthor ws, 10, c1, lastCol, "　　　　管理者　", CfgText("cfg_Manager")
    WriteAuthor ws, 11, c1, lastCol, "　　　　担当者　", CfgText("cfg_Staff")
    WriteAuthor ws, 12, c1, lastCol, "　　　　確認先　", CfgText("cfg_Contact")

    ws.Cells(13, 1).Value = CfgText("cfg_Preface")

    ' 変更がある場合は凡例
    If anyChg Then
        With MergeRow(ws, OR_LEGEND, SpanStart(ws, lastCol, 50), lastCol)
            .Cells(1, 1).Value = "（上段：変更前　下段：変更後）"
            .HorizontalAlignment = xlRight
            .Font.Color = chgColor
        End With
    End If

    ' ---- 業務名・開始・終了 ----
    ws.Cells(OR_PROJECT, OC_NAME).Value = "業務名"
    With MergeRow(ws, OR_PROJECT, OC_DAYS, lastCol)
        .Cells(1, 1).Value = CfgText("cfg_Project")
        .HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    ws.Rows(OR_PROJECT).RowHeight = 30

    c1 = OC_CHART + (nMonths * 3) \ 2
    c2 = SpanEnd(ws, c1, 14, lastCol - 1)
    ws.Cells(OR_PERIOD, OC_NAME).Value = "開始年月日"
    With MergeRow(ws, OR_PERIOD, OC_DAYS, c1 - 1)
        .Cells(1, 1).Value = curS
        .NumberFormat = DATE_FMT
        .HorizontalAlignment = xlCenter
    End With
    With MergeRow(ws, OR_PERIOD, c1, c2)
        .Cells(1, 1).Value = "終了予定年月日"
        .HorizontalAlignment = xlDistributed
        .ShrinkToFit = True
    End With
    With MergeRow(ws, OR_PERIOD, c2 + 1, lastCol)
        .Cells(1, 1).Value = curE
        .NumberFormat = DATE_FMT
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(OR_PROJECT, OC_NAME), ws.Cells(OR_PERIOD, OC_NAME))
        .HorizontalAlignment = xlDistributed
        .IndentLevel = 1
    End With
    ws.Rows(OR_PERIOD).RowHeight = 18

    ' ---- 見出し（月・旬） ----
    ws.Cells(OR_MONTH, OC_NAME).Value = "月別"
    ws.Cells(OR_MONTH, OC_NAME).HorizontalAlignment = xlRight
    ws.Cells(OR_TEN, OC_NAME).Value = "工種"
    ws.Cells(OR_TEN, OC_NAME).HorizontalAlignment = xlLeft
    With ws.Range(ws.Cells(OR_MONTH, OC_DAYS), ws.Cells(OR_TEN, OC_DAYS))
        .Merge
        .Cells(1, 1).Value = "日数"
        .HorizontalAlignment = xlCenter
    End With
    For k = 0 To nMonths - 1
        c = OC_CHART + k * 3
        With MergeRow(ws, OR_MONTH, c, c + 2)
            .Cells(1, 1).Value = DateAdd("m", k, monthBase)
            .NumberFormat = MONTH_FMT
            .HorizontalAlignment = xlCenter
            .ShrinkToFit = True
        End With
        ws.Cells(OR_TEN, c).Value = 10
        ws.Cells(OR_TEN, c + 1).Value = 20
    Next k
    With ws.Range(ws.Cells(OR_TEN, OC_CHART), ws.Cells(OR_TEN, lastCol - 1))
        .Font.Size = 7
        .HorizontalAlignment = xlRight
    End With
    With ws.Range(ws.Cells(OR_MONTH, lastCol), ws.Cells(OR_TEN, lastCol))
        .Merge
        .Cells(1, 1).Value = "備　　考"
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(OR_MONTH).RowHeight = 16
    ws.Rows(OR_TEN).RowHeight = 12

    ' ---- 作業行 ----
    ws.Range(ws.Cells(OR_FIRST, OC_NAME), ws.Cells(lastRow, OC_NAME)).NumberFormat = "@"
    ws.Range(ws.Cells(OR_FIRST, lastCol), ws.Cells(lastRow, lastCol)).NumberFormat = "@"
    For i = 1 To nItems
        r = OR_FIRST + i - 1
        ws.Cells(r, OC_NAME).Value = itemName(i)
        ws.Cells(r, lastCol).Value = itemNote(i)
        If itemHasChg(i) Then
            ' 日数は上段に変更前、下段に変更後
            ws.Rows(r).RowHeight = ROW_H2
            If itemDays(i) > 0 Then dTop = CStr(itemDays(i)) Else dTop = "－"
            dBot = CStr(itemCDays(i))
            With ws.Cells(r, OC_DAYS)
                .Value = dTop & vbLf & dBot
                .WrapText = True
                .Characters(Len(dTop) + 2, Len(dBot)).Font.Color = chgColor
            End With
        Else
            ws.Rows(r).RowHeight = ROW_H
            If itemDays(i) > 0 Then ws.Cells(r, OC_DAYS).Value = itemDays(i)
        End If
    Next i
    ws.Range(ws.Cells(OR_FIRST, OC_NAME), ws.Cells(lastRow, OC_NAME)).ShrinkToFit = True
    ws.Range(ws.Cells(OR_FIRST, lastCol), ws.Cells(lastRow, lastCol)).ShrinkToFit = True
    ws.Range(ws.Cells(OR_FIRST, OC_DAYS), ws.Cells(lastRow, OC_DAYS)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(OR_PROJECT, OC_NAME), ws.Cells(lastRow, lastCol)).VerticalAlignment = xlCenter

    ' ---- 罫線（月の境は実線、旬の境は点線） ----
    Set rg = ws.Range(ws.Cells(OR_PROJECT, OC_NAME), ws.Cells(lastRow, lastCol))
    rg.Borders.LineStyle = xlContinuous
    rg.Borders.Weight = xlThin
    ws.Cells(OR_MONTH, OC_NAME).Borders(xlEdgeBottom).LineStyle = xlNone
    For k = 0 To nMonths - 1
        c = OC_CHART + k * 3
        ws.Range(ws.Cells(OR_TEN, c), ws.Cells(OR_TEN, c + 2)).Borders(xlInsideVertical).LineStyle = xlNone
        With ws.Range(ws.Cells(OR_FIRST, c + 1), ws.Cells(lastRow, c + 2))
            .Borders(xlEdgeLeft).LineStyle = xlDot
            .Borders(xlInsideVertical).LineStyle = xlDot
        End With
    Next k

    ' 「月別／工種」の斜線
    With ws.Cells(OR_MONTH, OC_NAME)
        Set shp = ws.Shapes.AddLine(.Left, .Top, .Left + .Width, .Top + .Height + ws.Cells(OR_TEN, OC_NAME).Height)
    End With
    shp.Line.ForeColor.RGB = RGB(0, 0, 0)
    shp.Line.Weight = 0.75
    shp.Name = "Diagonal"

    ' ---- バー ----
    pct = Val(CfgText("cfg_BarHeight"))
    If pct < 10 Or pct > 90 Then pct = 35
    barH = ROW_H * pct / 100
    For i = 1 To nPers
        it = perItem(i)
        r = OR_FIRST + it - 1
        rowTop = ws.Rows(r).Top
        rowHt = ws.Rows(r).Height
        If Not itemHasChg(it) Then
            DrawBar ws, i, "", perStart(i), perEnd(i), rowTop + rowHt / 2, barH, perColor(i), STYLE_FILL
        Else
            ' 上段：変更前　下段：変更後（変更のない期間はそのまま下段にも描く）
            If perHasOrig(i) Then
                DrawBar ws, i, "Old", perStart(i), perEnd(i), rowTop + rowHt * 0.3, barH, perColor(i), oldStyle
            End If
            If perHasChg(i) Then
                DrawBar ws, i, "New", perCStart(i), perCEnd(i), rowTop + rowHt * 0.72, barH, chgColor, STYLE_FILL
            Else
                DrawBar ws, i, "New", perStart(i), perEnd(i), rowTop + rowHt * 0.72, barH, chgColor, STYLE_FILL
            End If
        End If
    Next i

    SetupPrint ws, lastRow, lastCol
End Sub

' バーを1本描く（yMid はバーの縦の中心）
Private Sub DrawBar(ByVal ws As Worksheet, ByVal idx As Long, ByVal tag As String, _
                    ByVal s As Date, ByVal e As Date, ByVal yMid As Double, ByVal h As Double, _
                    ByVal color As Long, ByVal style As Long)
    Dim x1 As Double, x2 As Double, shp As Shape

    x1 = DateX(ws, s, False)
    x2 = DateX(ws, e, True)
    If x2 - x1 < 2 Then x2 = x1 + 2

    If style = STYLE_DASH Then
        Set shp = ws.Shapes.AddLine(x1, yMid, x2, yMid)
        shp.Line.ForeColor.RGB = color
        shp.Line.Weight = h * 0.6
        shp.Line.DashStyle = LINE_DASH
    Else
        Set shp = ws.Shapes.AddShape(SHAPE_RECT, x1, yMid - h / 2, x2 - x1, h)
        If style = STYLE_OUTLINE Then
            shp.Fill.Visible = MSO_TRUE
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
            shp.Line.Visible = MSO_TRUE
            shp.Line.ForeColor.RGB = color
            shp.Line.Weight = 1
        Else
            shp.Fill.Visible = MSO_TRUE
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = color
            shp.Line.Visible = MSO_FALSE
        End If
        shp.Shadow.Visible = MSO_FALSE
    End If
    shp.Name = "Bar" & idx & tag
    shp.Placement = xlMoveAndSize
End Sub

' 日付 → 工程表上の横位置（ポイント）。atEnd=True ならその日の終わり
Private Function DateX(ByVal ws As Worksheet, ByVal d As Date, ByVal atEnd As Boolean) As Double
    Dim mi As Long, dy As Long, j As Long, s0 As Long, ln As Long, pos As Double

    mi = (Year(d) - Year(monthBase)) * 12 + Month(d) - Month(monthBase)
    dy = Day(d)
    If dy <= 10 Then
        j = 0: s0 = 1: ln = 10
    ElseIf dy <= 20 Then
        j = 1: s0 = 11: ln = 10
    Else
        j = 2: s0 = 21: ln = Day(DateSerial(Year(d), Month(d) + 1, 0)) - 20
    End If
    pos = dy - s0
    If atEnd Then pos = pos + 1
    With ws.Cells(OR_FIRST, OC_CHART + mi * 3 + j)
        DateX = .Left + .Width * pos / ln
    End With
End Function

Private Sub WriteAuthor(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, _
                        ByVal label As String, ByVal txt As String)
    If txt = "" Then Exit Sub
    With MergeRow(ws, r, c1, c2)
        .Cells(1, 1).Value = label & txt
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub SetupPrint(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    ' プリンタが無い環境でも止まらないようにエラーは無視する
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
        .PrintTitleRows = "$" & OR_PROJECT & ":$" & OR_TEN
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------
' 日数の計算（暦日 / 土日祝除く）
'----------------------------------------------------------------------
Private Function AddDays(ByVal s As Date, ByVal d As Long) As Date
    If useWorkDays Then AddDays = WorkDayEx(s - 1, d) Else AddDays = s + d - 1
End Function

Private Function SubDays(ByVal e As Date, ByVal d As Long) As Date
    If useWorkDays Then SubDays = WorkDayEx(e + 1, -d) Else SubDays = e - d + 1
End Function

Private Function NextStart(ByVal e As Date) As Date
    If useWorkDays Then NextStart = WorkDayEx(e, 1) Else NextStart = e + 1
End Function

Private Function CountDays(ByVal s As Date, ByVal e As Date) As Long
    If Not useWorkDays Then
        CountDays = e - s + 1
    ElseIf holRange Is Nothing Then
        CountDays = Application.WorksheetFunction.NetworkDays(s, e)
    Else
        CountDays = Application.WorksheetFunction.NetworkDays(s, e, holRange)
    End If
End Function

Private Function WorkDayEx(ByVal d As Date, ByVal n As Long) As Date
    If holRange Is Nothing Then
        WorkDayEx = Application.WorksheetFunction.WorkDay(d, n)
    Else
        WorkDayEx = Application.WorksheetFunction.WorkDay(d, n, holRange)
    End If
End Function

Private Sub LoadHolidays()
    Dim wsH As Worksheet, last As Long
    Set holRange = Nothing
    On Error Resume Next
    Set wsH = ThisWorkbook.Worksheets(HOL_SHEET)
    On Error GoTo 0
    If wsH Is Nothing Then Exit Sub
    last = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row
    If last >= 2 Then Set holRange = wsH.Range(wsH.Cells(2, 1), wsH.Cells(last, 1))
End Sub

'----------------------------------------------------------------------
' セルの読み取り
'----------------------------------------------------------------------
Private Function ReadDate(ByVal cell As Range, ByVal label As String, ByRef result As Date) As Boolean
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then RowError cell.Row, label & "がエラー値になっています。"
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then Exit Function
        If Not IsDate(v) Then RowError cell.Row, label & "「" & v & "」を日付として読み取れません。"
    ElseIf VarType(v) <> vbDate And Not IsNumeric(v) Then
        RowError cell.Row, label & "を日付として読み取れません。"
    End If
    result = CDate(Int(CDbl(CDate(v))))
    ReadDate = True
End Function

Private Function ReadDays(ByVal cell As Range, ByVal label As String, ByRef result As Long) As Boolean
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then RowError cell.Row, label & "がエラー値になっています。"
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then Exit Function
    End If
    If Not IsNumeric(v) Then RowError cell.Row, label & "は数字で入力してください。"
    If CDbl(v) < 1 Then RowError cell.Row, label & "は1以上で入力してください。"
    result = -Int(-CDbl(v))
    ReadDays = True
End Function

Private Function CellText(ByVal cell As Range) As String
    If IsError(cell.Value) Then Exit Function
    CellText = Trim$(CStr(cell.Value))
End Function

Private Function CellColor(ByVal cell As Range, ByVal def As Long) As Long
    If cell Is Nothing Then
        CellColor = def
    ElseIf cell.Interior.ColorIndex = xlColorIndexNone Then
        CellColor = def
    Else
        CellColor = cell.Interior.Color
    End If
End Function

Private Sub RowError(ByVal r As Long, ByVal msg As String)
    Err.Raise vbObjectError + 2, , "入力シート " & r & " 行目：" & msg
End Sub

'----------------------------------------------------------------------
' 名前付きセル（入力シートの設定欄）
'----------------------------------------------------------------------
Private Function CfgRange(ByVal nm As String) As Range
    On Error Resume Next
    Set CfgRange = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
End Function

Private Function CfgValue(ByVal nm As String) As Variant
    Dim rg As Range
    Set rg = CfgRange(nm)
    If rg Is Nothing Then
        CfgValue = Empty
    ElseIf IsError(rg.Cells(1, 1).Value) Then
        CfgValue = Empty
    Else
        CfgValue = rg.Cells(1, 1).Value
    End If
End Function

Private Function CfgText(ByVal nm As String) As String
    CfgText = Trim$(CStr(CfgValue(nm)))
End Function

'----------------------------------------------------------------------
' シート操作の小道具
'----------------------------------------------------------------------
Private Function GetOutSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(IN_SHEET))
        ws.Name = OUT_SHEET
    End If
    Set GetOutSheet = ws
End Function

Private Sub ClearSheet(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Font.Name = OUT_FONT
    ws.Cells.Font.Size = 9
    ws.Columns.ColumnWidth = 9
    ws.Rows.RowHeight = 13.5
End Sub

Private Function MergeRow(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As Range
    Dim rg As Range
    Set rg = ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    rg.Merge
    Set MergeRow = rg
End Function

Private Sub PutDateOrText(ByVal cell As Range, ByVal v As Variant)
    If IsEmpty(v) Then Exit Sub
    If VarType(v) = vbDate Then
        cell.Value = v
        cell.NumberFormat = DATE_FMT
    Else
        cell.Value = CStr(v)
    End If
End Sub

' lastCol から左へ、幅の合計が w 以上になる最初の列
Private Function SpanStart(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal w As Double) As Long
    Dim c As Long, total As Double
    c = lastCol
    total = ws.Columns(c).ColumnWidth
    Do While total < w And c > OC_CHART
        c = c - 1
        total = total + ws.Columns(c).ColumnWidth
    Loop
    SpanStart = c
End Function

' c1 から右へ、幅の合計が w 以上になる最後の列（maxCol まで）
Private Function SpanEnd(ByVal ws As Worksheet, ByVal c1 As Long, ByVal w As Double, ByVal maxCol As Long) As Long
    Dim c As Long, total As Double
    c = c1
    total = ws.Columns(c).ColumnWidth
    Do While total < w And c < maxCol
        c = c + 1
        total = total + ws.Columns(c).ColumnWidth
    Loop
    SpanEnd = c
End Function
