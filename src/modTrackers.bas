Attribute VB_Name = "modTrackers"
' Strategy Tools - trackers: small solid circles with a number, letter or
' short custom text. "Next" looks at the trackers already on the slide.
' Each tracker is tagged ST_TYPE = TRACKER, ST_KIND = NUM / LET / CUSTOM.
Option Explicit

Private Const TRACKER_SIZE As Single = 21.6   ' 0.3 inch
Private Const MAX_CUSTOM As Long = 3

Public Sub ST_TrackerNumber(): AddTracker "NUM", "": End Sub
Public Sub ST_TrackerLetter(): AddTracker "LET", "": End Sub

Public Sub ST_TrackerCustom()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim answer As String
    answer = Trim$(InputBox("Tracker text (up to " & MAX_CUSTOM & " characters):", ST_APP_TITLE, ""))
    If Len(answer) = 0 Then Exit Sub
    AddTracker "CUSTOM", Left$(answer, MAX_CUSTOM)
    Exit Sub
fail:
    ReportError "Tracker"
End Sub

' kind: NUM, LET or CUSTOM. text is only used for CUSTOM.
Public Sub AddTracker(ByVal kind As String, ByVal text As String)
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim sld As Object, shp As Object, last As Object, best As Long, v As Long
    Set sld = CurrentSlide()

    ' Find the highest tracker of this kind already on the slide.
    For Each shp In sld.Shapes
        If TagOf(shp, "ST_TYPE") = "TRACKER" And TagOf(shp, "ST_KIND") = kind Then
            v = TrackerValue(kind, TagOf(shp, "ST_TEXT"))
            If v >= best Then best = v: Set last = shp
        End If
    Next shp

    Select Case kind
        Case "NUM": text = CStr(best + 1)
        Case "LET": text = LetterFor(best + 1)
    End Select

    ' Start below the previous tracker of this kind (or mid-slide for the
    ' first), then step down / across until the spot is free of trackers.
    Dim x As Single, y As Single, startY As Single, tries As Long, pitch As Single
    pitch = TRACKER_SIZE * 1.8
    If last Is Nothing Then
        x = (ActivePresentation.PageSetup.SlideWidth - TRACKER_SIZE) / 2
        y = (ActivePresentation.PageSetup.SlideHeight - TRACKER_SIZE) / 2
    Else
        x = last.Left
        y = last.Top + pitch
    End If
    startY = y
    Do While SpotTaken(sld, x, y) And tries < 60
        y = y + pitch
        If y + TRACKER_SIZE > ActivePresentation.PageSetup.SlideHeight Then
            x = x + pitch
            y = startY
        End If
        tries = tries + 1
    Loop

    Dim t As Object
    Set t = sld.Shapes.AddShape(ST_SHAPE_OVAL, x, y, TRACKER_SIZE, TRACKER_SIZE)
    t.Fill.Solid
    t.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
    NoOutline t
    With t.TextFrame
        .WordWrap = ST_FALSE
        .AutoSize = ppAutoSizeNone
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = 3          ' msoAnchorMiddle
        With .TextRange
            .Text = text
            .ParagraphFormat.Alignment = ppAlignCenter
            .Font.Bold = ST_TRUE
            .Font.Color.RGB = RGB(255, 255, 255)
            .Font.Size = IIf(Len(text) >= 3, 8, IIf(Len(text) = 2, 9.5, 11))
        End With
    End With
    t.Tags.Add "ST_TYPE", "TRACKER"
    t.Tags.Add "ST_KIND", kind
    t.Tags.Add "ST_TEXT", text
    t.Name = "Tracker " & text
    t.Select
    Exit Sub
fail:
    ReportError "Tracker"
End Sub

Private Function SpotTaken(ByVal sld As Object, ByVal x As Single, ByVal y As Single) As Boolean
    Dim shp As Object
    For Each shp In sld.Shapes
        If TagOf(shp, "ST_TYPE") = "TRACKER" Then
            If Abs(shp.Left - x) < TRACKER_SIZE And Abs(shp.Top - y) < TRACKER_SIZE Then
                SpotTaken = True
                Exit Function
            End If
        End If
    Next shp
End Function

Private Function TrackerValue(ByVal kind As String, ByVal text As String) As Long
    Dim i As Long
    On Error GoTo nope
    Select Case kind
        Case "NUM": TrackerValue = CLng(text)
        Case "LET"
            ' A=1 ... Z=26, AA=27 ...
            For i = 1 To Len(text)
                TrackerValue = TrackerValue * 26 + (Asc(UCase$(Mid$(text, i, 1))) - 64)
            Next i
        Case Else: TrackerValue = 1
    End Select
    Exit Function
nope:
    TrackerValue = 0
End Function

Private Function LetterFor(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(65 + (n - 1) Mod 26) & s
        n = (n - 1) \ 26
    Loop
    LetterFor = s
End Function
