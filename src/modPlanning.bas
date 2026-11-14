Attribute VB_Name = "modPlanning"
' Strategy Tools - Gantt / timeline and issue tree (Frameworks group).
' Both are one group: ST_TYPE = GANTT / ISSUETREE.
Option Explicit

Private Const DIAMOND As Long = 4            ' msoShapeDiamond
Private Function Grid() As Long: Grid = RGB(225, 229, 235): End Function
Private Function HeadFill() As Long: HeadFill = RGB(238, 241, 245): End Function
Private Function Muted() As Long: Muted = RGB(90, 100, 118): End Function
Private Function Amber() As Long: Amber = RGB(240, 162, 2): End Function

' ================================================================ Gantt

Public Sub ST_Gantt()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim rows As String, span As String
    rows = InputBox("How many activity rows? (2 to 10)", ST_APP_TITLE, "4")
    If Len(Trim$(rows)) = 0 Then Exit Sub
    If Not IsNumeric(rows) Then
        Msg "Please enter a number from 2 to 10."
        Exit Sub
    End If
    span = InputBox("Time span, in months or weeks (e.g. ""6 months"" or ""12 weeks""):", ST_APP_TITLE, "6 months")
    If Len(Trim$(span)) = 0 Then Exit Sub
    AddGantt CLng(rows), span
    Exit Sub
fail:
    ReportError "Gantt"
End Sub

' Parses "6 months" / "12 weeks" / "6m" / "12w". Returns column count (0 if
' invalid) and sets weeks = True for a weekly scale.
Public Function ParseSpan(ByVal span As String, ByRef weeks As Boolean) As Long
    Dim s As String, num As String, i As Long, ch As String
    s = LCase$(Trim$(span))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then num = num & ch Else Exit For
    Next i
    If Len(num) = 0 Then Exit Function
    weeks = (InStr(s, "w") > 0)
    If Not weeks And InStr(s, "m") = 0 Then Exit Function
    ParseSpan = CLng(num)
    If ParseSpan < 2 Or ParseSpan > IIf(weeks, 26, 24) Then ParseSpan = 0
End Function

Public Function AddGantt(ByVal nRows As Long, ByVal span As String) As Object
    If nRows < 2 Or nRows > 10 Then
        Msg "Please enter a number of rows from 2 to 10."
        Exit Function
    End If
    Dim weeks As Boolean, nCols As Long
    nCols = ParseSpan(span, weeks)
    If nCols = 0 Then
        Msg "Please give the span as a number of months (2-24) or weeks (2-26), e.g. ""6 months"" or ""12 weeks""."
        Exit Function
    End If
    Dim sld As Object, l As Single, t As Single, w As Single, h As Single
    Set sld = CurrentSlide()
    PlanArea sld, l, t, w, h
    Const HEAD As Single = 22
    Dim labelW As Single, gx As Single, gw As Single, colW As Single, pitch As Single, barH As Single
    labelW = w * 0.24
    gx = l + labelW
    gw = w - labelW
    colW = gw / nCols
    pitch = (h - HEAD - 6) / (nRows + 1)          ' +1 row for milestones
    If pitch > 34 Then pitch = 34
    barH = pitch * 0.55
    Dim parts As New Collection, shp As Object, i As Long, y As Single
    ' Header band and period labels.
    Set shp = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, gx, t, gw, HEAD)
    shp.Fill.Solid: shp.Fill.ForeColor.RGB = HeadFill(): NoOutline shp
    parts.Add shp
    For i = 1 To nCols
        Set shp = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, gx + (i - 1) * colW, t, colW, HEAD)
        PlainText shp, IIf(weeks, "W" & i, Format(DateSerial(Year(Date), Month(Date) + i - 1, 1), "mmm")), 10, False, Muted(), ppAlignCenter
        parts.Add shp
        If i > 1 Then
            Set shp = sld.Shapes.AddLine(gx + (i - 1) * colW, t + HEAD, gx + (i - 1) * colW, t + HEAD + pitch * (nRows + 1))
            shp.Line.ForeColor.RGB = Grid(): shp.Line.Weight = 0.75
            parts.Add shp
        End If
    Next i
    ' Activity rows: label + a staggered bar to stretch into place.
    Dim barLen As Single, barX As Single
    barLen = gw * 0.3
    For i = 1 To nRows
        y = t + HEAD + (i - 1) * pitch
        Set shp = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, l, y, labelW - 8, pitch)
        PlainText shp, "Activity " & i, 12, False, 0, ppAlignLeft
        shp.TextFrame.TextRange.Font.Color.ObjectThemeColor = 13
        parts.Add shp
        barX = gx + (gw - barLen) * (i - 1) / (nRows - 1)
        Set shp = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, barX, y + (pitch - barH) / 2, barLen, barH)
        shp.Fill.Solid: shp.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1: NoOutline shp
        shp.Name = "ST_Bar" & i
        parts.Add shp
    Next i
    ' Milestone row with two markers to copy and drag.
    y = t + HEAD + nRows * pitch
    Set shp = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, l, y, labelW - 8, pitch)
    PlainText shp, "Milestones", 12, False, 0, ppAlignLeft
    shp.TextFrame.TextRange.Font.Color.ObjectThemeColor = 13
    parts.Add shp
    Dim m As Long, dsz As Single
    dsz = barH * 1.1
    For m = 1 To 2
        Set shp = sld.Shapes.AddShape(DIAMOND, gx + gw * IIf(m = 1, 0.3, 0.85) - dsz / 2, y + (pitch - dsz) / 2, dsz, dsz)
        shp.Fill.Solid: shp.Fill.ForeColor.RGB = Amber(): NoOutline shp
        parts.Add shp
    Next m
    Set AddGantt = GroupParts(sld, parts, "GANTT", "Gantt (" & nRows & " rows)")
End Function

' ================================================================ issue tree

Public Sub ST_IssueTree()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim src As Object, sr As Object
    Set sr = SelectedShapes()
    If Not sr Is Nothing Then
        If sr.Count = 1 Then
            If sr(1).HasTextFrame = ST_TRUE Then Set src = sr(1)
        End If
    End If
    If src Is Nothing Then
        AddOutlineStarter
        Msg "Edit this outline: one item per line, press Tab to indent a level (up to 4)." & vbCr & vbCr & _
            "Then select the box and click Issue Tree again."
        Exit Sub
    End If
    If BuildIssueTree(src) Is Nothing Then Exit Sub
    Exit Sub
fail:
    ReportError "Issue tree"
End Sub

Public Function AddOutlineStarter() As Object
    Dim sld As Object, tb As Object, i As Long, levels As Variant
    Set sld = CurrentSlide()
    Set tb = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, ContentLeft(), ActivePresentation.PageSetup.SlideHeight * 0.25, 360, 200)
    tb.TextFrame.TextRange.Text = "How can we grow profit by 20%?" & vbCr & "Increase revenue" & vbCr & "Raise prices" & vbCr & _
        "Sell more volume" & vbCr & "Reduce costs" & vbCr & "Cut fixed costs" & vbCr & "Cut variable costs"
    levels = Array(1, 2, 3, 3, 2, 3, 3)
    For i = 1 To 7
        With tb.TextFrame.TextRange.Paragraphs(i)
            .IndentLevel = levels(i - 1)
            .ParagraphFormat.Bullet.Visible = ST_TRUE
        End With
    Next i
    tb.Name = "Issue tree outline"
    tb.Select
    Set AddOutlineStarter = tb
End Function

' Builds the tree from src's paragraphs (indent level, plus leading tabs),
' deletes src, returns the group (Nothing if the outline is unusable).
Public Function BuildIssueTree(ByVal src As Object) As Object
    Dim n As Long, i As Long, txt As String, lvl As Long, cnt As Long
    Dim texts() As String, levels() As Long, parent() As Long, depth As Long
    n = src.TextFrame.TextRange.Paragraphs.Count
    ReDim texts(1 To n): ReDim levels(1 To n): ReDim parent(1 To n)
    For i = 1 To n
        txt = src.TextFrame.TextRange.Paragraphs(i).Text
        txt = Replace(Replace(txt, vbCr, ""), Chr(11), "")
        lvl = src.TextFrame.TextRange.Paragraphs(i).IndentLevel
        Do While Left$(txt, 1) = vbTab
            lvl = lvl + 1
            txt = Mid$(txt, 2)
        Loop
        txt = Trim$(txt)
        If Len(txt) > 0 Then
            cnt = cnt + 1
            texts(cnt) = txt
            levels(cnt) = lvl
        End If
    Next i
    If cnt < 2 Then
        Msg "The outline needs at least two lines: a question and the branches under it."
        Exit Function
    End If
    ' Normalise so the shallowest item is level 1, then cap at 4.
    Dim minL As Long
    minL = levels(1)
    For i = 1 To cnt
        If levels(i) < minL Then minL = levels(i)
    Next i
    For i = 1 To cnt
        levels(i) = levels(i) - minL + 1
        If i > 1 Then
            If levels(i) > levels(i - 1) + 1 Then levels(i) = levels(i - 1) + 1   ' no skipped levels
        End If
        If levels(i) > 4 Then levels(i) = 4
        If levels(i) > depth Then depth = levels(i)
    Next i
    ' Parent = nearest earlier item one level up.
    For i = 1 To cnt
        parent(i) = 0
        Dim j As Long
        For j = i - 1 To 1 Step -1
            If levels(j) = levels(i) - 1 Then parent(i) = j: Exit For
        Next j
    Next i
    ' Vertical slots: leaves in order, parents centred on their children.
    Dim isLeaf() As Boolean, leaves As Long, ySlot() As Double
    ReDim isLeaf(1 To cnt): ReDim ySlot(1 To cnt)
    For i = 1 To cnt
        isLeaf(i) = True
        If i < cnt Then
            If levels(i + 1) > levels(i) Then isLeaf(i) = False
        End If
        If isLeaf(i) Then leaves = leaves + 1: ySlot(i) = leaves - 1
    Next i
    Dim lv As Long, sumY As Double, kids As Long
    For lv = depth To 1 Step -1
        For i = 1 To cnt
            If levels(i) = lv And Not isLeaf(i) Then
                sumY = 0: kids = 0
                For j = 1 To cnt
                    If parent(j) = i Then sumY = sumY + ySlot(j): kids = kids + 1
                Next j
                If kids > 0 Then ySlot(i) = sumY / kids
            End If
        Next i
    Next lv
    ' Geometry.
    Dim sld As Object, l As Single, t As Single, w As Single, h As Single
    Set sld = src.Parent
    PlanArea sld, l, t, w, h
    Const COL_GAP As Single = 34
    Dim colW As Single, slot As Single, boxH As Single
    colW = (w - (depth - 1) * COL_GAP) / depth
    slot = h / leaves
    boxH = slot * 0.78
    If boxH > 46 Then boxH = 46
    Dim parts As New Collection, box As Object, x() As Single, cy() As Single
    ReDim x(1 To cnt): ReDim cy(1 To cnt)
    For i = 1 To cnt
        x(i) = l + (levels(i) - 1) * (colW + COL_GAP)
        cy(i) = t + slot * (ySlot(i) + 0.5)
        Set box = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, x(i), cy(i) - boxH / 2, colW, boxH)
        box.Fill.Solid
        NoOutline box
        Select Case levels(i)
            Case 1
                box.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
                PlainText box, texts(i), 13, True, RGB(255, 255, 255), ppAlignCenter
            Case 2
                box.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
                box.Fill.Transparency = 0.82
                PlainText box, texts(i), 12, True, 0, ppAlignCenter
                box.TextFrame.TextRange.Font.Color.ObjectThemeColor = 13
            Case Else
                box.Fill.ForeColor.RGB = RGB(241, 244, 248)
                PlainText box, texts(i), 11, False, 0, ppAlignLeft
                box.TextFrame.TextRange.Font.Color.ObjectThemeColor = 13
        End Select
        box.TextFrame.VerticalAnchor = 3
        box.TextFrame.MarginLeft = 8: box.TextFrame.MarginRight = 8
        parts.Add box
    Next i
    ' Elbow connectors parent -> child.
    Dim pts(1 To 4, 1 To 2) As Single, ln As Object, p As Long, mx As Single
    For i = 1 To cnt
        p = parent(i)
        If p > 0 Then
            mx = x(p) + colW + COL_GAP / 2
            pts(1, 1) = x(p) + colW: pts(1, 2) = cy(p)
            pts(2, 1) = mx: pts(2, 2) = cy(p)
            pts(3, 1) = mx: pts(3, 2) = cy(i)
            pts(4, 1) = x(i): pts(4, 2) = cy(i)
            Set ln = sld.Shapes.AddPolyline(pts)
            ln.Fill.Visible = ST_FALSE
            ln.Line.ForeColor.RGB = RGB(138, 147, 163)
            ln.Line.Weight = 1.5
            parts.Add ln
        End If
    Next i
    src.Delete
    Set BuildIssueTree = GroupParts(sld, parts, "ISSUETREE", "Issue tree")
End Function

' ================================================================ helpers

' Area under the title, above the footer (same rules as the other frameworks).
Private Sub PlanArea(ByVal sld As Object, ByRef l As Single, ByRef t As Single, ByRef w As Single, ByRef h As Single)
    Dim slideH As Single
    slideH = ActivePresentation.PageSetup.SlideHeight
    l = ContentLeft()
    w = ActivePresentation.PageSetup.SlideWidth - 2 * l
    t = slideH * 0.22
    On Error Resume Next
    If sld.Shapes.HasTitle = ST_TRUE Then
        If sld.Shapes.Title.Top + sld.Shapes.Title.Height < slideH * 0.4 Then t = sld.Shapes.Title.Top + sld.Shapes.Title.Height + 12
    End If
    On Error GoTo 0
    h = slideH * 0.86 - t
End Sub

Private Sub PlainText(ByVal shp As Object, ByVal text As String, ByVal size As Single, ByVal bold As Boolean, _
                      ByVal colour As Long, ByVal align As Long)
    With shp.TextFrame
        .WordWrap = ST_TRUE
        .AutoSize = ppAutoSizeNone
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        With .TextRange
            .Text = text
            .ParagraphFormat.Alignment = align
            .ParagraphFormat.Bullet.Visible = ST_FALSE
            .Font.Size = size
            .Font.Bold = IIf(bold, ST_TRUE, ST_FALSE)
            .Font.Color.RGB = colour
        End With
    End With
End Sub

Private Function GroupParts(ByVal sld As Object, ByVal parts As Collection, ByVal kind As String, ByVal nm As String) As Object
    Dim arr() As Variant, names() As Variant, i As Long, grp As Object
    ReDim arr(1 To parts.Count)
    ReDim names(1 To parts.Count)
    For i = 1 To parts.Count
        Set arr(i) = parts(i)
        names(i) = "ST_" & kind & "_" & i
    Next i
    Set grp = GroupShapes(sld, arr, names)
    grp.Tags.Add "ST_TYPE", kind
    grp.Name = nm
    grp.Select
    Set GroupParts = grp
End Function
