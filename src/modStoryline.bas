Attribute VB_Name = "modStoryline"
' Strategy Tools - storyline: agenda + highlighted dividers, section tags,
' storyline overview slide and the deck check report.
'
' Slide tags (ST_SLIDE): AGENDA (the master agenda you type into), DIVIDER
' (a copy placed before a section, ST_SECTION = its number), STORYLINE and
' DECKCHECK (generated report slides; Clean for Sending removes them).
' Shape tags (ST_TYPE): AGENDA_NUM / AGENDA_ITEM / AGENDA_BAND with ST_INDEX,
' and SECTIONTAG for the small top-right section label.
Option Explicit

Private Const MIN_SECTIONS As Long = 2
Private Const MAX_SECTIONS As Long = 8
Private Const LINES_PER_SLIDE As Long = 16

Private Function DimGrey() As Long: DimGrey = RGB(201, 206, 214): End Function
Private Function DimText() As Long: DimText = RGB(150, 158, 171): End Function

' ================================================================ agenda

Public Sub ST_AgendaNew()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    If Not AgendaMaster() Is Nothing Then
        If Not Confirm("This deck already has an agenda slide. Add another one anyway?") Then Exit Sub
    End If
    Dim answer As String
    answer = InputBox("How many sections? (" & MIN_SECTIONS & " to " & MAX_SECTIONS & ")", ST_APP_TITLE, "4")
    If Len(Trim$(answer)) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then
        Msg "Please enter a number from " & MIN_SECTIONS & " to " & MAX_SECTIONS & "."
        Exit Sub
    End If
    AddAgenda CLng(answer)
    Exit Sub
fail:
    ReportError "New agenda"
End Sub

' Inserts the master agenda after the current slide and returns it.
Public Function AddAgenda(ByVal n As Long) As Object
    If n < MIN_SECTIONS Or n > MAX_SECTIONS Then
        Msg "Please enter a number from " & MIN_SECTIONS & " to " & MAX_SECTIONS & "."
        Exit Function
    End If
    Dim sld As Object, i As Long
    Set sld = ActivePresentation.Slides.AddSlide(CurrentSlide().SlideIndex + 1, TitleOnlyLayout())
    sld.Tags.Add "ST_SLIDE", "AGENDA"
    SetSlideTitle sld, "Agenda"
    Dim slideW As Single, slideH As Single, left0 As Single, top0 As Single, pitch As Single, box As Single, width As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    left0 = ContentLeft()
    top0 = slideH * 0.24
    pitch = slideH * 0.66 / n
    If pitch > slideH * 0.12 Then pitch = slideH * 0.12
    box = pitch * 0.72
    width = slideW - 2 * left0
    For i = 1 To n
        Dim y As Single, band As Object, num As Object, item As Object
        y = top0 + (i - 1) * pitch
        ' Highlight band (hidden on the master, shown on the current divider row).
        Set band = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, left0, y, width, box)
        band.Fill.Solid
        band.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
        band.Fill.Transparency = 0.86
        NoOutline band
        band.Visible = ST_FALSE
        TagAgenda band, "AGENDA_BAND", i
        Set num = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, left0, y, box, box)
        num.Fill.Solid
        num.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
        NoOutline num
        With num.TextFrame
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .WordWrap = ST_FALSE
            .VerticalAnchor = 3
            With .TextRange
                .Text = CStr(i)
                .ParagraphFormat.Alignment = ppAlignCenter
                .Font.Bold = ST_TRUE
                .Font.Size = 16
                .Font.Color.RGB = RGB(255, 255, 255)
            End With
        End With
        TagAgenda num, "AGENDA_NUM", i
        Set item = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, left0 + box + 12, y, width - box - 12, box)
        With item.TextFrame
            .WordWrap = ST_TRUE
            .AutoSize = ppAutoSizeNone
            .VerticalAnchor = 3
            .MarginLeft = 6
            With .TextRange
                .Text = "Section " & i
                .ParagraphFormat.Alignment = ppAlignLeft
                .Font.Size = 20
            End With
        End With
        item.Top = y: item.Height = box     ' undo the automatic shrink-to-fit
        TagAgenda item, "AGENDA_ITEM", i
    Next i
    ActiveWindow.View.GotoSlide sld.SlideIndex
    Set AddAgenda = sld
End Function

Private Sub TagAgenda(ByVal shp As Object, ByVal kind As String, ByVal i As Long)
    shp.Tags.Add "ST_TYPE", kind
    shp.Tags.Add "ST_INDEX", CStr(i)
    shp.Name = kind & "_" & i
End Sub

' The master agenda slide (the one you type into), or Nothing.
Public Function AgendaMaster() As Object
    Dim sld As Object
    For Each sld In ActivePresentation.Slides
        If TagOf(sld, "ST_SLIDE") = "AGENDA" Then
            Set AgendaMaster = sld
            Exit Function
        End If
    Next sld
End Function

Public Function SectionCount() As Long
    Dim m As Object
    Set m = AgendaMaster()
    If m Is Nothing Then Exit Function
    SectionCount = TaggedShapes(m, "AGENDA_ITEM").Count
End Function

Public Function SectionName(ByVal i As Long) As String
    Dim m As Object, shp As Object
    Set m = AgendaMaster()
    If m Is Nothing Then Exit Function
    For Each shp In m.Shapes
        If TagOf(shp, "ST_TYPE") = "AGENDA_ITEM" And TagOf(shp, "ST_INDEX") = CStr(i) Then
            SectionName = Trim$(Replace(shp.TextFrame.TextRange.Text, vbCr, " "))
            Exit Function
        End If
    Next shp
End Function

' Puts a highlighted copy of the agenda in front of the current slide.
Public Sub ST_AgendaDivider()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim m As Object
    Set m = AgendaMaster()
    If m Is Nothing Then
        Msg "Create the agenda first (Agenda > New agenda), then add dividers."
        Exit Sub
    End If
    Dim here As Long, k As Long, sld As Object
    here = CurrentSlide().SlideIndex
    If here <= m.SlideIndex Then
        Msg "Dividers go after the agenda slide. Select the first slide of a section further on."
        Exit Sub
    End If
    ' Section number = dividers already before this slide + 1.
    For Each sld In ActivePresentation.Slides
        If sld.SlideIndex < here And TagOf(sld, "ST_SLIDE") = "DIVIDER" Then k = k + 1
    Next sld
    k = k + 1
    If k > SectionCount() Then
        Msg "Every section already has a divider before this slide. Add a section to the agenda first."
        Exit Sub
    End If
    Dim d As Object
    Set d = MakeDivider(here, k)
    RenumberDividers
    ActiveWindow.View.GotoSlide d.SlideIndex
    Exit Sub
fail:
    ReportError "Agenda divider"
End Sub

' Copies the master agenda to position idx, styled for section k.
Public Function MakeDivider(ByVal idx As Long, ByVal k As Long) As Object
    Dim d As Object
    Set d = AgendaMaster().Duplicate.Item(1)
    d.MoveTo idx
    d.Tags.Add "ST_SLIDE", "DIVIDER"
    StyleDivider d, k
    Set MakeDivider = d
End Function

Private Sub StyleDivider(ByVal d As Object, ByVal k As Long)
    Dim shp As Object, i As Long, current As Boolean
    d.Tags.Add "ST_SECTION", CStr(k)
    For Each shp In d.Shapes
        If Len(TagOf(shp, "ST_INDEX")) > 0 Then
            i = CLng(TagOf(shp, "ST_INDEX"))
            current = (i = k)
            Select Case TagOf(shp, "ST_TYPE")
                Case "AGENDA_BAND"
                    shp.Visible = IIf(current, ST_TRUE, ST_FALSE)
                Case "AGENDA_NUM"
                    If current Then shp.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1 Else shp.Fill.ForeColor.RGB = DimGrey()
                Case "AGENDA_ITEM"
                    With shp.TextFrame.TextRange.Font
                        .Bold = IIf(current, ST_TRUE, ST_FALSE)
                        If current Then .Color.ObjectThemeColor = ST_THEME_ACCENT1 Else .Color.RGB = DimText()
                    End With
            End Select
        End If
    Next shp
End Sub

' Dividers are numbered by their order in the deck.
Private Sub RenumberDividers()
    Dim sld As Object, k As Long
    For Each sld In ActivePresentation.Slides
        If TagOf(sld, "ST_SLIDE") = "DIVIDER" Then
            k = k + 1
            If TagOf(sld, "ST_SECTION") <> CStr(k) Then StyleDivider sld, k
        End If
    Next sld
End Sub

' Rebuild every divider from the master (names, number of sections) and
' refresh section tags if the deck uses them.
Public Sub ST_AgendaUpdate()
    On Error GoTo fail
    If AgendaMaster() Is Nothing Then
        Msg "This deck has no agenda yet (Agenda > New agenda)."
        Exit Sub
    End If
    Dim positions As New Collection, sld As Object, i As Long, hadTags As Boolean, n As Long
    hadTags = DeckHasSectionTags()
    For i = ActivePresentation.Slides.Count To 1 Step -1
        Set sld = ActivePresentation.Slides(i)
        If TagOf(sld, "ST_SLIDE") = "DIVIDER" Then
            positions.Add i
            sld.Delete
        End If
    Next i
    ' positions were collected last-first; re-insert first-last.
    For i = positions.Count To 1 Step -1
        n = n + 1
        If n <= SectionCount() Then MakeDivider positions(i), n
    Next i
    If hadTags Then ApplySectionTags
    Msg "Agenda updated: " & IIf(n > SectionCount(), SectionCount(), n) & " divider(s)" & IIf(hadTags, " and section tags", "") & " refreshed."
    Exit Sub
fail:
    ReportError "Update agenda"
End Sub

' ================================================================ section tags

Public Sub ST_SectionTags()
    On Error GoTo fail
    If AgendaMaster() Is Nothing Then
        Msg "Create the agenda and its dividers first (Agenda menu)."
        Exit Sub
    End If
    Dim n As Long
    n = ApplySectionTags()
    Msg "Section tags added to " & n & " slide(s). Run Update agenda after renaming sections or moving slides."
    Exit Sub
fail:
    ReportError "Section tags"
End Sub

' Tags every content slide with the section of the divider before it.
' Returns how many slides carry a tag.
Public Function ApplySectionTags() As Long
    Dim sld As Object, k As Long, kind As String, label As String
    For Each sld In ActivePresentation.Slides
        kind = TagOf(sld, "ST_SLIDE")
        RemoveSectionTag sld
        If kind = "DIVIDER" Then
            k = CLng(TagOf(sld, "ST_SECTION"))
        ElseIf kind = "" And k > 0 Then
            label = k & "  " & ChrW(183) & "  " & UCase$(SectionName(k))
            AddSectionTag sld, label
            ApplySectionTags = ApplySectionTags + 1
        End If
    Next sld
End Function

Private Sub RemoveSectionTag(ByVal sld As Object)
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        If TagOf(sld.Shapes(i), "ST_TYPE") = "SECTIONTAG" Then sld.Shapes(i).Delete
    Next i
End Sub

Private Sub AddSectionTag(ByVal sld As Object, ByVal label As String)
    Dim shp As Object, slideW As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    Set shp = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, slideW * 0.6, ActivePresentation.PageSetup.SlideHeight * 0.03, 100, 16)
    shp.Fill.Solid
    shp.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
    shp.Fill.Transparency = 0.86
    NoOutline shp
    With shp.TextFrame
        .WordWrap = ST_FALSE
        .MarginLeft = 6: .MarginRight = 6: .MarginTop = 1: .MarginBottom = 1
        .VerticalAnchor = 3
        With .TextRange
            .Text = label
            .ParagraphFormat.Alignment = ppAlignCenter
            .Font.Bold = ST_TRUE
            .Font.Size = 8
            .Font.Color.ObjectThemeColor = ST_THEME_ACCENT1
        End With
        .AutoSize = ppAutoSizeShapeToFitText
    End With
    shp.Left = slideW - slideW * 0.03 - shp.Width
    shp.Tags.Add "ST_TYPE", "SECTIONTAG"
    shp.Name = "Section tag"
    ArrangeTopRight sld
End Sub

Private Function DeckHasSectionTags() As Boolean
    Dim sld As Object
    For Each sld In ActivePresentation.Slides
        If TaggedShapes(sld, "SECTIONTAG").Count > 0 Then
            DeckHasSectionTags = True
            Exit Function
        End If
    Next sld
End Function

' When a slide has both a section tag and a stamp, put the stamp to the
' left of the tag so they don't overlap.
Public Sub ArrangeTopRight(ByVal sld As Object)
    Dim tags As Collection, stamps As Collection
    Set tags = TaggedShapes(sld, "SECTIONTAG")
    Set stamps = TaggedShapes(sld, "STAMP")
    If tags.Count = 0 Or stamps.Count = 0 Then Exit Sub
    stamps(1).Left = tags(1).Left - stamps(1).Width - 8
    stamps(1).Top = tags(1).Top + (tags(1).Height - stamps(1).Height) / 2
End Sub

' ================================================================ storyline

Public Sub ST_Storyline()
    On Error GoTo fail
    If Presentations.Count = 0 Then
        Msg "Open a presentation first."
        Exit Sub
    End If
    Dim lines As Collection, made As Long
    RemoveReportSlides "STORYLINE"
    Set lines = StorylineLines()
    made = WriteReport("Storyline", "STORYLINE", lines)
    Msg "Storyline added at the end of the deck (" & made & " slide" & IIf(made = 1, "", "s") & "). Clean for Sending removes it."
    Exit Sub
fail:
    ReportError "Storyline"
End Sub

' One entry per slide: "#3  Title" or a section header line.
' Entries starting with "##" are headers (written bold).
Public Function StorylineLines() As Collection
    Dim c As New Collection, sld As Object, kind As String, k As Long, t As String
    For Each sld In ActivePresentation.Slides
        kind = TagOf(sld, "ST_SLIDE")
        Select Case kind
            Case "DIVIDER"
                k = CLng(TagOf(sld, "ST_SECTION"))
                c.Add "##" & k & ". " & SectionName(k)
            Case "STORYLINE", "DECKCHECK", "AGENDA"
                ' skip report and agenda slides
            Case Else
                t = SlideTitleText(sld)
                If Len(t) = 0 Then t = "(no title)"
                c.Add sld.SlideIndex & vbTab & t
        End Select
    Next sld
    Set StorylineLines = c
End Function

' ================================================================ deck check

Public Sub ST_DeckCheck()
    On Error GoTo fail
    If Presentations.Count = 0 Then
        Msg "Open a presentation first."
        Exit Sub
    End If
    Dim issues As Collection
    RemoveReportSlides "DECKCHECK"
    Set issues = DeckIssues()
    If issues.Count = 0 Then
        Msg "Deck check: no issues found."
        Exit Sub
    End If
    WriteReport "Deck check: " & issues.Count & " issue" & IIf(issues.Count = 1, "", "s"), "DECKCHECK", issues
    ActiveWindow.View.GotoSlide ActivePresentation.Slides.Count - ReportSlideCount("DECKCHECK") + 1
    Msg "Deck check found " & issues.Count & " issue(s). They're listed on the report slide at the end. Clean for Sending removes it."
    Exit Sub
fail:
    ReportError "Deck check"
End Sub

Public Function DeckIssues() As Collection
    Dim c As New Collection, sld As Object, shp As Object, heading As String, body As String
    Dim slideW As Single, slideH As Single, kind As String
    heading = ActivePresentation.SlideMaster.Theme.ThemeFontScheme.MajorFont(1).Name
    body = ActivePresentation.SlideMaster.Theme.ThemeFontScheme.MinorFont(1).Name
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    For Each sld In ActivePresentation.Slides
        kind = TagOf(sld, "ST_SLIDE")
        If kind <> "STORYLINE" And kind <> "DECKCHECK" Then
            Dim n As Long, fonts As String
            n = sld.SlideIndex
            fonts = ""
            If Len(SlideTitleText(sld)) = 0 Then c.Add n & vbTab & "No title"
            For Each shp In sld.Shapes
                Select Case TagOf(shp, "ST_TYPE")
                    Case "STAMP": c.Add n & vbTab & "Stamp still on slide (""" & TagOf(shp, "ST_TEXT") & """)"
                    Case "STICKY": c.Add n & vbTab & "Sticky note still on slide"
                End Select
                If IsEmptyPlaceholder(shp) Then c.Add n & vbTab & "Empty placeholder (""" & shp.Name & """)"
                If shp.Left < -1 Or shp.Top < -1 Or shp.Left + shp.Width > slideW + 1 Or shp.Top + shp.Height > slideH + 1 Then
                    c.Add n & vbTab & "Object partly off the slide (""" & shp.Name & """)"
                End If
                If TextOverflows(shp) Then c.Add n & vbTab & "Text overflows its box (""" & shp.Name & """)"
                fonts = StrayFontsIn(shp, heading, body, fonts)
            Next shp
            If Len(fonts) > 0 Then c.Add n & vbTab & "Non-template font: " & Mid$(fonts, 3)
        End If
    Next sld
    Set DeckIssues = c
End Function

Private Function IsEmptyPlaceholder(ByVal shp As Object) As Boolean
    Dim t As Long
    On Error GoTo nope
    If shp.Type <> ST_TYPE_PLACEHOLDER Then Exit Function
    t = shp.PlaceholderFormat.Type
    ' Titles are reported as "No title"; footer-type placeholders are fine empty.
    If t = ppPlaceholderTitle Or t = ppPlaceholderCenterTitle Then Exit Function
    If t = ppPlaceholderSlideNumber Or t = ppPlaceholderDate Or t = ppPlaceholderFooter Then Exit Function
    If shp.HasTextFrame <> ST_TRUE Then Exit Function
    IsEmptyPlaceholder = (Len(Trim$(shp.TextFrame.TextRange.Text)) = 0)
nope:
End Function

Private Function TextOverflows(ByVal shp As Object) As Boolean
    On Error GoTo nope
    If shp.HasTextFrame <> ST_TRUE Then Exit Function
    If Len(shp.TextFrame.TextRange.Text) = 0 Then Exit Function
    If shp.TextFrame.AutoSize <> ppAutoSizeNone Then Exit Function
    TextOverflows = (shp.TextFrame.TextRange.BoundHeight > shp.Height + 2)
nope:
End Function

' Appends ", Font" for each stray font in shp (and its group items) not yet in list.
Private Function StrayFontsIn(ByVal shp As Object, ByVal heading As String, ByVal body As String, ByVal list As String) As String
    Dim i As Long, n As Long, nm As String
    StrayFontsIn = list
    On Error GoTo done
    If shp.Type = ST_TYPE_GROUP Then
        n = shp.GroupItems.Count
        For i = 1 To n
            StrayFontsIn = StrayFontsIn(shp.GroupItems(i), heading, body, StrayFontsIn)
        Next i
        Exit Function
    End If
    If shp.HasTextFrame <> ST_TRUE Then Exit Function
    n = shp.TextFrame.TextRange.Runs.Count
    For i = 1 To n
        nm = shp.TextFrame.TextRange.Runs(i).Font.Name
        If Len(nm) > 0 And StrComp(nm, heading, vbTextCompare) <> 0 And StrComp(nm, body, vbTextCompare) <> 0 _
           And Left$(nm, 1) <> "+" And InStr(1, StrayFontsIn & ",", ", " & nm & ",", vbTextCompare) = 0 Then
            StrayFontsIn = StrayFontsIn & ", " & nm
        End If
    Next i
done:
End Function

' ================================================================ report slides

' Writes lines onto report slides at the end of the deck, LINES_PER_SLIDE per
' slide. Lines are "<slide no><TAB><text>" or "##header". Returns slide count.
Private Function WriteReport(ByVal title As String, ByVal kind As String, ByVal lines As Collection) As Long
    Dim total As Long, pages As Long, pg As Long, i As Long, first As Long, last As Long
    total = lines.Count
    pages = (total + LINES_PER_SLIDE - 1) \ LINES_PER_SLIDE
    If pages = 0 Then pages = 1
    For pg = 1 To pages
        Dim sld As Object, tb As Object, txt As String, ln As String
        Set sld = ActivePresentation.Slides.AddSlide(ActivePresentation.Slides.Count + 1, TitleOnlyLayout())
        sld.Tags.Add "ST_SLIDE", kind
        SetSlideTitle sld, title & IIf(pages > 1, " (" & pg & "/" & pages & ")", "")
        first = (pg - 1) * LINES_PER_SLIDE + 1
        last = first + LINES_PER_SLIDE - 1
        If last > total Then last = total
        txt = ""
        For i = first To last
            ln = lines(i)
            If Left$(ln, 2) = "##" Then ln = Mid$(ln, 3) Else ln = "Slide " & Replace(ln, vbTab, "   ", 1, 1)
            txt = txt & IIf(Len(txt) > 0, vbCr, "") & ln
        Next i
        Set tb = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, ContentLeft(), ActivePresentation.PageSetup.SlideHeight * 0.2, _
                                       ActivePresentation.PageSetup.SlideWidth - 2 * ContentLeft(), ActivePresentation.PageSetup.SlideHeight * 0.7)
        With tb.TextFrame
            .WordWrap = ST_TRUE
            .TextRange.Text = txt
            .TextRange.Font.Size = 15
            .TextRange.ParagraphFormat.Bullet.Visible = ST_FALSE
            .TextRange.ParagraphFormat.SpaceAfter = 3
        End With
        ' Bold the section headers.
        For i = first To last
            If Left$(lines(i), 2) = "##" Then
                With tb.TextFrame.TextRange.Paragraphs(i - first + 1).Font
                    .Bold = ST_TRUE
                    .Color.ObjectThemeColor = ST_THEME_ACCENT1
                End With
            End If
        Next i
        tb.Name = kind & " text"
    Next pg
    WriteReport = pages
End Function

Public Sub RemoveReportSlides(ByVal kind As String)
    Dim i As Long
    For i = ActivePresentation.Slides.Count To 1 Step -1
        If TagOf(ActivePresentation.Slides(i), "ST_SLIDE") = kind Then ActivePresentation.Slides(i).Delete
    Next i
End Sub

Public Function ReportSlideCount(ByVal kind As String) As Long
    Dim sld As Object
    For Each sld In ActivePresentation.Slides
        If TagOf(sld, "ST_SLIDE") = kind Then ReportSlideCount = ReportSlideCount + 1
    Next sld
End Function

' ================================================================ helpers

Public Function SlideTitleText(ByVal sld As Object) As String
    On Error GoTo nope
    If sld.Shapes.HasTitle = ST_TRUE Then
        SlideTitleText = Trim$(Replace(Replace(sld.Shapes.Title.TextFrame.TextRange.Text, vbCr, " "), Chr(11), " "))
    End If
nope:
End Function

Private Sub SetSlideTitle(ByVal sld As Object, ByVal text As String)
    On Error GoTo nope
    If sld.Shapes.HasTitle = ST_TRUE Then sld.Shapes.Title.TextFrame.TextRange.Text = text
nope:
End Sub

Public Function TitleOnlyLayout() As Object
    Dim lay As Object
    For Each lay In ActivePresentation.SlideMaster.CustomLayouts
        If InStr(1, lay.Name, "Title Only", vbTextCompare) > 0 Then
            Set TitleOnlyLayout = lay
            Exit Function
        End If
    Next lay
    Set TitleOnlyLayout = ActivePresentation.SlideMaster.CustomLayouts(1)
End Function
