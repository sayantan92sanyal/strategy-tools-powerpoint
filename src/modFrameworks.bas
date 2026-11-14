Attribute VB_Name = "modFrameworks"
' Strategy Tools - frameworks: 2x2 matrix, process flow, key takeaways panel,
' rating scale and CAGR arrow. Each is one group tagged ST_TYPE = MATRIX /
' PROCESS / TAKEAWAY / RATING / CAGR.
Option Explicit

Private Const ARROW_TRIANGLE As Long = 2      ' msoArrowheadTriangle
Private Const THEME_TEXT1 As Long = 13       ' msoThemeColorText1
Private Const SHAPE_ROUNDED As Long = 5      ' msoShapeRoundedRectangle

Private Function Ink() As Long: Ink = RGB(24, 32, 47): End Function
Private Function AxisGrey() As Long: AxisGrey = RGB(90, 100, 118): End Function

' ---------------------------------------------------------------- body area

' The slide area under the title and above the source footer.
Private Sub BodyArea(ByVal sld As Object, ByRef l As Single, ByRef t As Single, ByRef w As Single, ByRef h As Single)
    Dim slideW As Single, slideH As Single, bottom As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    l = ContentLeft()
    w = slideW - 2 * l
    t = slideH * 0.22
    On Error Resume Next
    If sld.Shapes.HasTitle = ST_TRUE Then
        If sld.Shapes.Title.Top + sld.Shapes.Title.Height < slideH * 0.4 Then t = sld.Shapes.Title.Top + sld.Shapes.Title.Height + 12
    End If
    On Error GoTo 0
    bottom = slideH * 0.86
    h = bottom - t
End Sub

Private Function NeedSlide() As Boolean
    NeedSlide = HasEditableSlide()
    If Not NeedSlide Then Msg "Open a slide in Normal view first."
End Function

Private Sub StyleText(ByVal shp As Object, ByVal text As String, ByVal size As Single, ByVal bold As Boolean, _
                      ByVal colour As Long, ByVal align As Long, ByVal anchor As Long)
    With shp.TextFrame
        .WordWrap = ST_TRUE
        .VerticalAnchor = anchor
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

Private Function Finish(ByVal sld As Object, ByVal parts As Collection, ByVal kind As String, ByVal nm As String) As Object
    Dim arr() As Variant, names() As Variant, i As Long
    ReDim arr(1 To parts.Count)
    ReDim names(1 To parts.Count)
    For i = 1 To parts.Count
        Set arr(i) = parts(i)
        names(i) = "ST_" & kind & "_" & i
    Next i
    Dim grp As Object
    Set grp = GroupShapes(sld, arr, names)
    grp.Tags.Add "ST_TYPE", kind
    grp.Name = nm
    grp.Select
    Set Finish = grp
End Function

' ---------------------------------------------------------------- 2x2 matrix

Public Sub ST_Matrix()
    On Error GoTo fail
    If Not NeedSlide() Then Exit Sub
    AddMatrix
    Exit Sub
fail:
    ReportError "2x2 matrix"
End Sub

Public Function AddMatrix() As Object
    Dim sld As Object, l As Single, t As Single, w As Single, h As Single
    Set sld = CurrentSlide()
    BodyArea sld, l, t, w, h
    Const AXIS_GAP As Single = 26     ' room for the vertical axis label
    Const BASE_GAP As Single = 22     ' room for the horizontal axis label
    Const GAP As Single = 4
    Dim gx As Single, gy As Single, gw As Single, gh As Single, qw As Single, qh As Single
    gx = l + AXIS_GAP: gy = t: gw = w - AXIS_GAP: gh = h - BASE_GAP
    qw = (gw - GAP) / 2: qh = (gh - GAP) / 2
    Dim parts As New Collection, q As Object, i As Long, col As Long, rw As Long
    For i = 1 To 4
        col = (i - 1) Mod 2: rw = (i - 1) \ 2
        Set q = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, gx + col * (qw + GAP), gy + rw * (qh + GAP), qw, qh)
        q.Fill.Solid
        If i = 2 Then
            q.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
            q.Fill.Transparency = 0.84
        ElseIf i = 3 Then
            q.Fill.ForeColor.RGB = RGB(246, 247, 249)
        Else
            q.Fill.ForeColor.RGB = RGB(236, 239, 243)
        End If
        NoOutline q
        StyleText q, "Quadrant " & i, 14, True, 0, ppAlignLeft, 1
        q.TextFrame.TextRange.Font.Color.ObjectThemeColor = ST_THEME_ACCENT1
        q.TextFrame.MarginLeft = 10: q.TextFrame.MarginTop = 8
        parts.Add q
    Next i
    Dim ax As Object
    Set ax = sld.Shapes.AddLine(gx - 6, gy + gh + 6, gx + gw, gy + gh + 6)
    ax.Line.ForeColor.RGB = AxisGrey(): ax.Line.Weight = 2: ax.Line.EndArrowheadStyle = ARROW_TRIANGLE
    parts.Add ax
    Set ax = sld.Shapes.AddLine(gx - 6, gy + gh + 6, gx - 6, gy - 2)
    ax.Line.ForeColor.RGB = AxisGrey(): ax.Line.Weight = 2: ax.Line.EndArrowheadStyle = ARROW_TRIANGLE
    parts.Add ax
    Dim lab As Object
    Set lab = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, gx + gw - 240, gy + gh + 8, 240, 16)
    StyleText lab, "X axis label " & ChrW(8594), 11, False, AxisGrey(), ppAlignRight, 1
    lab.TextFrame.MarginRight = 0
    parts.Add lab
    Set lab = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, 0, 0, 200, 16)
    StyleText lab, "Y axis label " & ChrW(8594), 11, False, AxisGrey(), ppAlignRight, 1
    lab.Rotation = 270
    ' After a 270-degree turn the box is centred on the same point, so place by centre.
    lab.Left = (l + 8) - lab.Width / 2
    lab.Top = gy + gh / 2 - lab.Height / 2
    parts.Add lab
    Set AddMatrix = Finish(sld, parts, "MATRIX", "2x2 matrix")
End Function

' ---------------------------------------------------------------- process flow

Public Sub ST_ProcessFlow()
    On Error GoTo fail
    If Not NeedSlide() Then Exit Sub
    Dim answer As String
    answer = InputBox("How many steps? (2 to 7)", ST_APP_TITLE, "4")
    If Len(Trim$(answer)) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then
        Msg "Please enter a number from 2 to 7."
        Exit Sub
    End If
    AddProcess CLng(answer)
    Exit Sub
fail:
    ReportError "Process flow"
End Sub

Public Function AddProcess(ByVal n As Long) As Object
    If n < 2 Or n > 7 Then
        Msg "Please enter a number from 2 to 7."
        Exit Function
    End If
    Dim sld As Object, l As Single, t As Single, w As Single, h As Single
    Set sld = CurrentSlide()
    BodyArea sld, l, t, w, h
    Dim d As Single, pitch As Single, cy As Single, colW As Single, i As Long, cx As Single
    d = ActivePresentation.PageSetup.SlideHeight * 0.09
    colW = w / n
    pitch = colW
    cy = t + h * 0.12 + d / 2
    Dim parts As New Collection, ln As Object, c As Object, tb As Object
    ' Connecting line from the first to the last circle centre.
    Set ln = sld.Shapes.AddLine(l + colW / 2, cy, l + colW / 2 + (n - 1) * pitch, cy)
    ln.Line.ForeColor.RGB = RGB(201, 206, 214)
    ln.Line.Weight = 3
    parts.Add ln
    For i = 1 To n
        cx = l + colW / 2 + (i - 1) * pitch
        Set c = sld.Shapes.AddShape(ST_SHAPE_OVAL, cx - d / 2, cy - d / 2, d, d)
        c.Fill.Solid
        c.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
        NoOutline c
        StyleText c, CStr(i), 18, True, RGB(255, 255, 255), ppAlignCenter, 3
        c.TextFrame.WordWrap = ST_FALSE
        c.TextFrame.MarginLeft = 0: c.TextFrame.MarginRight = 0
        parts.Add c
        Set tb = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, cx - colW * 0.45, cy + d / 2 + 12, colW * 0.9, 60)
        StyleText tb, "Step " & i & vbCr & "Description", 12, False, AxisGrey(), ppAlignCenter, 1
        With tb.TextFrame.TextRange.Paragraphs(1).Font
            .Bold = ST_TRUE
            .Size = 14
            .Color.RGB = Ink()
        End With
        parts.Add tb
    Next i
    Set AddProcess = Finish(sld, parts, "PROCESS", "Process flow (" & n & ")")
End Function

' ---------------------------------------------------------------- key takeaways

Public Sub ST_Takeaways()
    On Error GoTo fail
    If Not NeedSlide() Then Exit Sub
    AddTakeaways
    Exit Sub
fail:
    ReportError "Key takeaways"
End Sub

Public Function AddTakeaways() As Object
    Dim sld As Object, existing As Collection
    Set sld = CurrentSlide()
    Set existing = TaggedShapes(sld, "TAKEAWAY")
    If existing.Count > 0 Then
        existing(1).Select
        Set AddTakeaways = existing(1)
        Exit Function
    End If
    Dim l As Single, t As Single, w As Single, h As Single, pw As Single, px As Single
    BodyArea sld, l, t, w, h
    pw = ActivePresentation.PageSetup.SlideWidth * 0.27
    px = l + w - pw
    Const HEAD As Single = 30
    Dim parts As New Collection, hd As Object, bd As Object
    Set hd = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, px, t, pw, HEAD)
    hd.Fill.Solid
    hd.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
    NoOutline hd
    StyleText hd, "Key takeaways", 14, True, RGB(255, 255, 255), ppAlignLeft, 3
    hd.TextFrame.MarginLeft = 10
    parts.Add hd
    Set bd = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, px, t + HEAD, pw, h - HEAD)
    bd.Fill.Solid
    bd.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
    bd.Fill.Transparency = 0.88
    NoOutline bd
    StyleText bd, "Takeaway 1" & vbCr & "Takeaway 2" & vbCr & "Takeaway 3", 14, False, 0, ppAlignLeft, 1
    With bd.TextFrame
        .MarginLeft = 10: .MarginRight = 10: .MarginTop = 10
        .TextRange.Font.Color.ObjectThemeColor = THEME_TEXT1
        With .TextRange.ParagraphFormat
            .SpaceAfter = 8
            .Bullet.Visible = ST_TRUE
            .Bullet.Type = ppBulletUnnumbered
            .Bullet.Character = 8226
        End With
    End With
    parts.Add bd
    Set AddTakeaways = Finish(sld, parts, "TAKEAWAY", "Key takeaways")
End Function

' ---------------------------------------------------------------- rating scale

Public Sub ST_Rating1(): RatingScale 1: End Sub
Public Sub ST_Rating2(): RatingScale 2: End Sub
Public Sub ST_Rating3(): RatingScale 3: End Sub
Public Sub ST_Rating4(): RatingScale 4: End Sub
Public Sub ST_Rating5(): RatingScale 5: End Sub

Public Sub RatingScale(ByVal score As Long)
    On Error GoTo fail
    If Not NeedSlide() Then Exit Sub
    Dim targets As Collection, g As Variant
    Set targets = SelectedVisuals("RATING")
    If targets.Count > 0 Then
        For Each g In targets
            StyleRating g, score
        Next g
        Exit Sub
    End If
    Const DOT As Single = 9
    Const DOT_GAP As Single = 3
    Dim sld As Object, x As Single, y As Single, i As Long, dots(1 To 5) As Object
    Set sld = CurrentSlide()
    x = (ActivePresentation.PageSetup.SlideWidth - (5 * DOT + 4 * DOT_GAP)) / 2
    y = (ActivePresentation.PageSetup.SlideHeight - DOT) / 2
    For i = 1 To 5
        Set dots(i) = sld.Shapes.AddShape(ST_SHAPE_OVAL, x + (i - 1) * (DOT + DOT_GAP), y, DOT, DOT)
        NoOutline dots(i)
    Next i
    Dim grp As Object
    Set grp = GroupShapes(sld, Array(dots(1), dots(2), dots(3), dots(4), dots(5)), _
                          Array("ST_Dot1", "ST_Dot2", "ST_Dot3", "ST_Dot4", "ST_Dot5"))
    grp.Tags.Add "ST_TYPE", "RATING"
    StyleRating grp, score
    grp.Select
    Exit Sub
fail:
    ReportError "Rating"
End Sub

Private Sub StyleRating(ByVal grp As Object, ByVal score As Long)
    Dim item As Object, n As Long
    For Each item In grp.GroupItems
        If Left$(item.Name, 6) = "ST_Dot" Then
            n = CLng(Mid$(item.Name, 7))
            item.Fill.Solid
            If n <= score Then
                item.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
            Else
                item.Fill.ForeColor.RGB = RGB(220, 225, 232)
            End If
        End If
    Next item
    grp.Tags.Add "ST_LEVEL", CStr(score)
    grp.Name = "Rating " & score & " of 5"
End Sub

' ---------------------------------------------------------------- CAGR arrow

Public Sub ST_Cagr()
    On Error GoTo fail
    If Not NeedSlide() Then Exit Sub
    Dim a As String, b As String, yrs As String
    a = InputBox("Start value:", ST_APP_TITLE, "")
    If Len(Trim$(a)) = 0 Then Exit Sub
    b = InputBox("End value:", ST_APP_TITLE, "")
    If Len(Trim$(b)) = 0 Then Exit Sub
    yrs = InputBox("Number of years between them:", ST_APP_TITLE, "")
    If Len(Trim$(yrs)) = 0 Then Exit Sub
    If Not (IsNumeric(a) And IsNumeric(b) And IsNumeric(yrs)) Then
        Msg "Please enter numbers only (e.g. 120, 185, 3)."
        Exit Sub
    End If
    AddCagr CDbl(a), CDbl(b), CDbl(yrs)
    Exit Sub
fail:
    ReportError "CAGR arrow"
End Sub

Public Function CagrLabel(ByVal startV As Double, ByVal endV As Double, ByVal years As Double) As String
    Dim r As Double, pct As Double, s As String
    r = (endV / startV) ^ (1 / years) - 1
    pct = r * 100
    If Abs(pct) < 10 Then s = Format(Abs(pct), "0.0") Else s = Format(Abs(pct), "0")
    CagrLabel = IIf(pct >= 0, "+", "-") & s & "% p.a."
End Function

' Joins the tops of two selected shapes (left to right) with an elbow arrow and
' a CAGR label; without two selected shapes, drops it mid-slide to drag into place.
Public Function AddCagr(ByVal startV As Double, ByVal endV As Double, ByVal years As Double) As Object
    If startV <= 0 Or endV <= 0 Or years <= 0 Then
        Msg "Start value, end value and years must all be above zero."
        Exit Function
    End If
    Dim sld As Object, sr As Object, a As Object, b As Object
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single, yTop As Single
    Set sld = CurrentSlide()
    Set sr = SelectedShapes()
    If Not sr Is Nothing Then
        If sr.Count = 2 Then
            Set a = sr(1): Set b = sr(2)
            If b.Left < a.Left Then Set a = sr(2): Set b = sr(1)
        End If
    End If
    If a Is Nothing Then
        x1 = ActivePresentation.PageSetup.SlideWidth * 0.3
        x2 = ActivePresentation.PageSetup.SlideWidth * 0.7
        y1 = ActivePresentation.PageSetup.SlideHeight * 0.45
        y2 = y1
    Else
        x1 = a.Left + a.Width / 2: y1 = a.Top - 4
        x2 = b.Left + b.Width / 2: y2 = b.Top - 4
    End If
    If y1 < y2 Then yTop = y1 Else yTop = y2
    yTop = yTop - 36
    If yTop < 4 Then yTop = 4
    Dim pts(1 To 4, 1 To 2) As Single, path As Object, lab As Object
    pts(1, 1) = x1: pts(1, 2) = y1
    pts(2, 1) = x1: pts(2, 2) = yTop
    pts(3, 1) = x2: pts(3, 2) = yTop
    pts(4, 1) = x2: pts(4, 2) = y2
    Set path = sld.Shapes.AddPolyline(pts)
    path.Fill.Visible = ST_FALSE
    With path.Line
        .Visible = ST_TRUE
        .ForeColor.RGB = Ink()
        .Weight = 1.5
        .EndArrowheadStyle = ARROW_TRIANGLE
    End With
    Set lab = sld.Shapes.AddShape(SHAPE_ROUNDED, (x1 + x2) / 2 - 50, yTop - 12, 100, 24)
    lab.Adjustments(1) = 0.5
    lab.Fill.Solid
    lab.Fill.ForeColor.RGB = RGB(255, 255, 255)
    lab.Line.ForeColor.RGB = Ink()
    lab.Line.Weight = 1
    StyleText lab, CagrLabel(startV, endV, years), 12, True, Ink(), ppAlignCenter, 3
    lab.TextFrame.WordWrap = ST_FALSE
    lab.TextFrame.MarginLeft = 4: lab.TextFrame.MarginRight = 4
    Dim parts As New Collection
    parts.Add path
    parts.Add lab
    Dim grp As Object
    Set grp = Finish(sld, parts, "CAGR", "CAGR " & CagrLabel(startV, endV, years))
    grp.Tags.Add "ST_TEXT", CagrLabel(startV, endV, years)
    Set AddCagr = grp
End Function
