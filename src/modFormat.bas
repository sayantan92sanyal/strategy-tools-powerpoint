Attribute VB_Name = "modFormat"
' Strategy Tools - one-click formatting: swap, match size, title harmonizer.
Option Explicit

' Swap two selected objects by their centre points. Sizes are unchanged.
Public Sub ST_Swap()
    On Error GoTo fail
    Dim sr As Object
    Set sr = SelectedShapes()
    If sr Is Nothing Then
        Msg "Select exactly 2 objects to swap."
        Exit Sub
    End If
    If sr.Count <> 2 Then
        Msg "Select exactly 2 objects to swap."
        Exit Sub
    End If
    Dim a As Object, b As Object
    Set a = sr(1)
    Set b = sr(2)
    Dim ax As Single, ay As Single, bx As Single, by As Single
    ax = a.Left + a.Width / 2: ay = a.Top + a.Height / 2
    bx = b.Left + b.Width / 2: by = b.Top + b.Height / 2
    a.Left = bx - a.Width / 2: a.Top = by - a.Height / 2
    b.Left = ax - b.Width / 2: b.Top = ay - b.Height / 2
    Exit Sub
fail:
    ReportError "Swap"
End Sub

Public Sub ST_SameSize()
    MatchSize True, True, "Same Size"
End Sub

Public Sub ST_SameWidth()
    MatchSize True, False, "Same Width"
End Sub

Public Sub ST_SameHeight()
    MatchSize False, True, "Same Height"
End Sub

' Resize every selected object to match the first one selected.
' The top-left corner of each object stays where it is.
Private Sub MatchSize(ByVal doWidth As Boolean, ByVal doHeight As Boolean, ByVal label As String)
    On Error GoTo fail
    Dim sr As Object
    Set sr = SelectedShapes()
    If sr Is Nothing Then
        Msg "Select 2 or more objects. The first one you click sets the size."
        Exit Sub
    End If
    If sr.Count < 2 Then
        Msg "Select 2 or more objects. The first one you click sets the size."
        Exit Sub
    End If
    Dim refW As Single, refH As Single
    refW = sr(1).Width
    refH = sr(1).Height
    Dim i As Long, s As Object, lockAR As Long
    For i = 2 To sr.Count
        Set s = sr(i)
        lockAR = s.LockAspectRatio
        s.LockAspectRatio = ST_FALSE
        If doWidth Then s.Width = refW
        If doHeight Then s.Height = refH
        s.LockAspectRatio = lockAR
    Next i
    Exit Sub
fail:
    ReportError label
End Sub

' Copy the selected title's font name and size to every other slide's title,
' skipping slides that use a Title Slide layout.
Public Sub ST_TitleHarmonizer()
    On Error GoTo fail
    Dim sr As Object
    Set sr = SelectedShapes()
    If sr Is Nothing Then
        Msg "Click into a slide title first, then press Title Harmonizer."
        Exit Sub
    End If
    Dim src As Object
    Set src = sr(1)
    If Not IsTitle(src) Then
        Msg "Click into a slide title first, then press Title Harmonizer."
        Exit Sub
    End If
    Dim fName As String, fSize As Single
    With src.TextFrame.TextRange
        If .Length > 0 Then
            fName = .Characters(1, 1).Font.Name
            fSize = .Characters(1, 1).Font.Size
        Else
            fName = .Font.Name
            fSize = .Font.Size
        End If
    End With

    Dim srcSlideId As Long
    srcSlideId = src.Parent.SlideID

    Dim sld As Object, ttl As Object, changed As Long
    For Each sld In ActivePresentation.Slides
        If sld.SlideID <> srcSlideId And Not IsTitleSlide(sld) Then
            Set ttl = SlideTitle(sld)
            If Not ttl Is Nothing Then
                With ttl.TextFrame.TextRange.Font
                    .Name = fName
                    .Size = fSize
                End With
                changed = changed + 1
            End If
        End If
    Next sld
    Msg "Title Harmonizer: updated " & changed & " slide title(s) to " & fName & " " & fSize & " pt."
    Exit Sub
fail:
    ReportError "Title Harmonizer"
End Sub

Private Function IsTitle(ByVal shp As Object) As Boolean
    On Error GoTo nope
    If shp.Type <> ST_TYPE_PLACEHOLDER Then GoTo nope
    Dim t As Long
    t = shp.PlaceholderFormat.Type
    IsTitle = (t = ppPlaceholderTitle Or t = ppPlaceholderCenterTitle)
    Exit Function
nope:
    IsTitle = False
End Function

Public Function IsTitleSlide(ByVal sld As Object) As Boolean
    On Error Resume Next
    IsTitleSlide = False
    If sld.Layout = ppLayoutTitle Then IsTitleSlide = True: Exit Function
    If LCase$(sld.CustomLayout.Name) = "title slide" Then IsTitleSlide = True
End Function

Private Function SlideTitle(ByVal sld As Object) As Object
    On Error GoTo nope
    Set SlideTitle = Nothing
    If sld.Shapes.HasTitle = ST_TRUE Then Set SlideTitle = sld.Shapes.Title
    Exit Function
nope:
    Set SlideTitle = Nothing
End Function

' Give selected shapes the same corner rounding (in points) as the first one
' clicked, so small and large boxes look identical.
Public Sub ST_SameCorners()
    On Error GoTo fail
    Dim sr As Object, refShape As Object, radius As Single, i As Long, s As Object, shortSide As Single
    Set sr = SelectedShapes()
    If sr Is Nothing Then GoTo needTwo
    If sr.Count < 2 Then GoTo needTwo
    Set refShape = sr(1)
    If Not HasCornerAdjustment(refShape) Then
        Msg "The first object you clicked has no adjustable corners. Start with a rounded rectangle."
        Exit Sub
    End If
    radius = refShape.Adjustments(1) * MinSide(refShape)
    For i = 2 To sr.Count
        Set s = sr(i)
        If HasCornerAdjustment(s) And s.AutoShapeType = refShape.AutoShapeType Then
            shortSide = MinSide(s)
            If shortSide > 0 Then
                If radius / shortSide > 0.5 Then s.Adjustments(1) = 0.5 Else s.Adjustments(1) = radius / shortSide
            End If
        End If
    Next i
    Exit Sub
needTwo:
    Msg "Select 2 or more rounded shapes. The first one you click sets the corners."
    Exit Sub
fail:
    ReportError "Same corners"
End Sub

Private Function HasCornerAdjustment(ByVal s As Object) As Boolean
    On Error GoTo nope
    HasCornerAdjustment = (s.Type = 1 And s.Adjustments.Count >= 1)   ' 1 = msoAutoShape
    Exit Function
nope:
    HasCornerAdjustment = False
End Function

Private Function MinSide(ByVal s As Object) As Single
    If s.Width < s.Height Then MinSide = s.Width Else MinSide = s.Height
End Function

' Copy fill, outline, shadow and text styling (font, colour, bold, italic)
' from the first object clicked to the others. Sizes and positions stay.
Public Sub ST_MatchFormat()
    On Error GoTo fail
    Dim sr As Object, refShape As Object, i As Long, s As Object
    Set sr = SelectedShapes()
    If sr Is Nothing Then GoTo needTwo
    If sr.Count < 2 Then GoTo needTwo
    Set refShape = sr(1)
    refShape.PickUp
    Dim hasText As Boolean, fName As String, fColour As Long, fBold As Long, fItalic As Long
    On Error Resume Next
    hasText = (refShape.HasTextFrame = ST_TRUE)
    If hasText Then
        With refShape.TextFrame.TextRange.Characters(1, 1).Font
            fName = .Name: fColour = .Color.RGB: fBold = .Bold: fItalic = .Italic
        End With
    End If
    On Error GoTo fail
    For i = 2 To sr.Count
        Set s = sr(i)
        Dim l As Single, t As Single, w As Single, h As Single
        l = s.Left: t = s.Top: w = s.Width: h = s.Height
        s.Apply
        s.Left = l: s.Top = t: s.Width = w: s.Height = h
        If hasText And Len(fName) > 0 Then
            On Error Resume Next
            If s.HasTextFrame = ST_TRUE Then
                With s.TextFrame.TextRange.Font
                    .Name = fName: .Color.RGB = fColour: .Bold = fBold: .Italic = fItalic
                End With
            End If
            On Error GoTo fail
        End If
    Next i
    Exit Sub
needTwo:
    Msg "Select 2 or more objects. The first one you click is the one to copy from."
    Exit Sub
fail:
    ReportError "Match format"
End Sub
