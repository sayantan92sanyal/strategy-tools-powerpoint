Attribute VB_Name = "modReview"
' Strategy Tools - deck hygiene: status stamps, sticky notes, source footer,
' back-up labels, clean-for-sending and the deck font fixer.
' Everything this module adds is tagged (ST_TYPE = STAMP / STICKY / FOOTER /
' BACKUP) so it can be found, updated and removed reliably.
Option Explicit

Private Const ANCHOR_BOTTOM As Long = 4      ' msoAnchorBottom
Private Const ANCHOR_TOP As Long = 1         ' msoAnchorTop

Private Function StampRed() As Long: StampRed = RGB(198, 40, 40): End Function
Private Function NoteGrey() As Long: NoteGrey = RGB(90, 100, 118): End Function

' ---------------------------------------------------------------- stamps

Public Sub ST_StampDraft():        AddStamp "Draft":          End Sub
Public Sub ST_StampPreliminary():  AddStamp "Preliminary":    End Sub
Public Sub ST_StampDiscussion():   AddStamp "For discussion": End Sub
Public Sub ST_StampConfidential(): AddStamp "Confidential":   End Sub

Public Sub ST_StampCustom()
    Dim answer As String
    answer = Trim$(InputBox("Stamp text:", ST_APP_TITLE, ""))
    If Len(answer) > 0 Then AddStamp Left$(answer, 30)
End Sub

' Adds (or re-labels) the stamp on the selected slides / current slide.
Public Sub AddStamp(ByVal label As String)
    On Error GoTo fail
    Dim slides As Collection, sld As Variant
    Set slides = TargetSlides()
    If slides.Count = 0 Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    For Each sld In slides
        StampOn sld, label
    Next sld
    Exit Sub
fail:
    ReportError "Stamp"
End Sub

' Copies the current slide's stamp text to every slide.
Public Sub ST_StampAllSlides()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim found As Collection, label As String, sld As Object
    Set found = TaggedShapes(CurrentSlide(), "STAMP")
    If found.Count = 0 Then
        Msg "Add a stamp to this slide first, then use this to copy it to every slide."
        Exit Sub
    End If
    label = TagOf(found(1), "ST_TEXT")
    For Each sld In ActivePresentation.Slides
        StampOn sld, label
    Next sld
    Msg "Stamp """ & label & """ is now on all " & ActivePresentation.Slides.Count & " slides."
    Exit Sub
fail:
    ReportError "Stamp"
End Sub

Public Sub ST_StampRemoveAll()
    Msg "Removed " & RemoveTagged(ActivePresentation, "STAMP") & " stamp(s)."
End Sub

Private Sub StampOn(ByVal sld As Object, ByVal label As String)
    Const W As Single = 112
    Const H As Single = 22
    Dim existing As Collection, shp As Object, slideW As Single, slideH As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    Set existing = TaggedShapes(sld, "STAMP")
    If existing.Count > 0 Then
        Set shp = existing(1)
    Else
        Set shp = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, slideW - slideW * 0.03 - W, slideH * 0.03, W, H)
        shp.Fill.Visible = ST_FALSE
        With shp.Line
            .Visible = ST_TRUE
            .ForeColor.RGB = StampRed()
            .Weight = 1.5
        End With
        With shp.TextFrame
            .WordWrap = ST_FALSE
            .MarginLeft = 4: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = 3
            With .TextRange
                .ParagraphFormat.Alignment = ppAlignCenter
                .Font.Bold = ST_TRUE
                .Font.Size = 10
                .Font.Color.RGB = StampRed()
            End With
        End With
        shp.Tags.Add "ST_TYPE", "STAMP"
    End If
    shp.TextFrame.TextRange.Text = UCase$(label)
    shp.Tags.Add "ST_TEXT", label
    shp.Name = "Stamp " & label
    ArrangeTopRight sld
End Sub

' ---------------------------------------------------------------- sticky notes

Public Sub ST_StickyAdd()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Const W As Single = 150
    Const H As Single = 80
    Dim shp As Object, who As String
    Set shp = CurrentSlide().Shapes.AddShape(ST_SHAPE_RECTANGLE, _
        (ActivePresentation.PageSetup.SlideWidth - W) / 2, (ActivePresentation.PageSetup.SlideHeight - H) / 2, W, H)
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = RGB(255, 233, 122)
    shp.Line.ForeColor.RGB = RGB(229, 201, 74)
    shp.Line.Weight = 0.75
    shp.Rotation = -2
    who = Initials()
    With shp.TextFrame
        .WordWrap = ST_TRUE
        .AutoSize = ppAutoSizeNone
        .VerticalAnchor = ANCHOR_TOP
        .MarginLeft = 7: .MarginRight = 7: .MarginTop = 6: .MarginBottom = 6
        .TextRange.Text = IIf(Len(who) > 0, vbCr & "- " & who, "")
        .TextRange.ParagraphFormat.Alignment = ppAlignLeft
        .TextRange.Font.Size = 11
        .TextRange.Font.Color.RGB = RGB(74, 61, 0)
    End With
    shp.Tags.Add "ST_TYPE", "STICKY"
    shp.Name = "Sticky note"
    shp.TextFrame.TextRange.Characters(1, 0).Select
    Exit Sub
fail:
    ReportError "Sticky note"
End Sub

Public Sub ST_StickyRemoveAll()
    Msg "Removed " & RemoveTagged(ActivePresentation, "STICKY") & " sticky note(s)."
End Sub

' Initials from the deck's last author (normally the person editing).
Private Function Initials() As String
    Dim nm As String, parts() As String, i As Long
    On Error Resume Next
    nm = Trim$(ActivePresentation.BuiltInDocumentProperties("Last Author").Value)
    If Len(nm) = 0 Then nm = Trim$(ActivePresentation.BuiltInDocumentProperties("Author").Value)
    On Error GoTo 0
    If Len(nm) = 0 Then Exit Function
    parts = Split(nm, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 And Len(Initials) < 3 Then Initials = Initials & UCase$(Left$(parts(i), 1))
    Next i
End Function

' ---------------------------------------------------------------- source footer

Public Sub ST_FooterSource():     AddFooter False: End Sub
Public Sub ST_FooterSourceNote(): AddFooter True:  End Sub

Public Sub AddFooter(ByVal withNote As Boolean)
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim sld As Object, existing As Collection, shp As Object, slideW As Single, slideH As Single
    Set sld = CurrentSlide()
    Set existing = TaggedShapes(sld, "FOOTER")
    If existing.Count > 0 Then
        ' One footer per slide: select the existing one instead.
        Set shp = existing(1)
        With shp.TextFrame.TextRange
            .Characters(.Length + 1, 0).Select
        End With
        Exit Sub
    End If
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    Set shp = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, ContentLeft(), slideH * 0.88, slideW * 0.62, 20)
    With shp.TextFrame
        .WordWrap = ST_TRUE
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = ANCHOR_BOTTOM
        .TextRange.Text = IIf(withNote, "Note: " & vbCr, "") & "Source: "
        .TextRange.ParagraphFormat.Alignment = ppAlignLeft
        .TextRange.ParagraphFormat.Bullet.Visible = ST_FALSE
        .TextRange.Font.Size = 8
        .TextRange.Font.Color.RGB = NoteGrey()
        .AutoSize = ppAutoSizeShapeToFitText
    End With
    ' Keep the bottom edge fixed at 94% of the slide height on every slide.
    shp.Top = slideH * 0.94 - shp.Height
    shp.Tags.Add "ST_TYPE", "FOOTER"
    shp.Name = "Source footer"
    With shp.TextFrame.TextRange
        .Characters(.Length + 1, 0).Select
    End With
    Exit Sub
fail:
    ReportError "Source footer"
End Sub

' ---------------------------------------------------------------- back-up

Public Sub ST_BackupMark()
    On Error GoTo fail
    Dim slides As Collection, sld As Variant, n As Long
    Set slides = TargetSlides()
    If slides.Count = 0 Then
        Msg "Select the slides to mark (in the thumbnail pane), or open one in Normal view."
        Exit Sub
    End If
    For Each sld In slides
        If TaggedShapes(sld, "BACKUP").Count = 0 Then
            BackupLabelOn sld
            n = n + 1
        End If
    Next sld
    If n = 0 Then
        Msg "Those slides are already marked as back-up."
        Exit Sub
    End If
    Dim first As Long
    first = slides(1).SlideIndex
    If HasAppendixBefore(first) Then
        Msg "Marked " & n & " slide(s) as back-up."
    ElseIf Confirm("Marked " & n & " slide(s) as back-up." & vbCr & vbCr & _
                   "Insert an ""Appendix"" divider slide in front of them?") Then
        InsertAppendixBefore first
    End If
    Exit Sub
fail:
    ReportError "Back-up label"
End Sub

Public Sub ST_BackupRemove()
    Dim slides As Collection, sld As Variant, n As Long, shp As Variant
    Set slides = TargetSlides()
    For Each sld In slides
        For Each shp In TaggedShapes(sld, "BACKUP")
            shp.Delete
            n = n + 1
        Next shp
    Next sld
    Msg "Removed " & n & " back-up label(s)."
End Sub

Public Sub BackupLabelOn(ByVal sld As Object)
    Dim shp As Object
    Set shp = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, ContentLeft(), _
        ActivePresentation.PageSetup.SlideHeight * 0.025, 58, 15)
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = NoteGrey()
    NoOutline shp
    With shp.TextFrame
        .WordWrap = ST_FALSE
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = 3
        With .TextRange
            .Text = "BACK-UP"
            .ParagraphFormat.Alignment = ppAlignCenter
            .Font.Bold = ST_TRUE
            .Font.Size = 8
            .Font.Color.RGB = RGB(255, 255, 255)
        End With
    End With
    shp.Tags.Add "ST_TYPE", "BACKUP"
    shp.Name = "Back-up label"
End Sub

' True if the slide just before idx is an Appendix divider.
Private Function HasAppendixBefore(ByVal idx As Long) As Boolean
    On Error Resume Next
    If idx <= 1 Then Exit Function
    HasAppendixBefore = (LCase$(Trim$(ActivePresentation.Slides(idx - 1).Shapes.Title.TextFrame.TextRange.Text)) = "appendix")
End Function

' Speaker-notes body text of a slide ("" if none).
' Note: never combine On Error Resume Next with a loop in VBA - an error in
' the loop header makes it spin forever. Errors here jump out instead.
Public Function NotesText(ByVal sld As Object) As String
    Dim body As Object
    Set body = NotesBody(sld)
    If Not body Is Nothing Then NotesText = body.TextFrame.TextRange.Text
End Function

Public Sub SetNotesText(ByVal sld As Object, ByVal text As String)
    Dim body As Object
    Set body = NotesBody(sld)
    If Not body Is Nothing Then body.TextFrame.TextRange.Text = text
End Sub

Private Function NotesBody(ByVal sld As Object) As Object
    Dim shapesOnPage As Object, i As Long, n As Long
    On Error GoTo none
    Set shapesOnPage = sld.NotesPage.Shapes
    n = shapesOnPage.Count
    For i = 1 To n
        If IsBodyPlaceholder(shapesOnPage(i)) Then
            Set NotesBody = shapesOnPage(i)
            Exit Function
        End If
    Next i
none:
End Function

Private Function IsBodyPlaceholder(ByVal shp As Object) As Boolean
    On Error GoTo nope
    If shp.Type <> ST_TYPE_PLACEHOLDER Then Exit Function
    IsBodyPlaceholder = (shp.PlaceholderFormat.Type = ppPlaceholderBody)
nope:
End Function

Private Sub DeleteComments(ByVal sld As Object)
    Dim n As Long, i As Long
    On Error GoTo done
    n = sld.Comments.Count
    For i = n To 1 Step -1
        sld.Comments(i).Delete
    Next i
done:
End Sub

' Inserts an "Appendix" divider at position idx, using a section-header
' layout when the template has one.
Public Function InsertAppendixBefore(ByVal idx As Long) As Object
    Dim lay As Object, chosen As Object, sld As Object
    For Each lay In ActivePresentation.SlideMaster.CustomLayouts
        If InStr(1, lay.Name, "Section", vbTextCompare) > 0 Then Set chosen = lay: Exit For
    Next lay
    If chosen Is Nothing Then
        For Each lay In ActivePresentation.SlideMaster.CustomLayouts
            If InStr(1, lay.Name, "Title Only", vbTextCompare) > 0 Then Set chosen = lay: Exit For
        Next lay
    End If
    If chosen Is Nothing Then Set chosen = ActivePresentation.SlideMaster.CustomLayouts(1)
    Set sld = ActivePresentation.Slides.AddSlide(idx, chosen)
    On Error Resume Next
    If sld.Shapes.HasTitle = ST_TRUE Then sld.Shapes.Title.TextFrame.TextRange.Text = "Appendix"
    Set InsertAppendixBefore = sld
End Function

' ---------------------------------------------------------------- clean for sending

Public Sub ST_CleanForSending()
    On Error GoTo fail
    If Presentations.Count = 0 Then
        Msg "Open a presentation first."
        Exit Sub
    End If
    Dim src As Object
    Set src = ActivePresentation
    If Len(src.Path) = 0 Then
        Msg "Save the presentation first. Clean for sending works on a copy saved next to it."
        Exit Sub
    End If
    If Not Confirm("Make a clean copy for sending?" & vbCr & vbCr & _
                   "The copy has no stamps, sticky notes, speaker notes or comments. This deck is not changed.") Then Exit Sub
    Dim copyPath As String
    copyPath = CleanCopy(src)
    Msg "Clean copy saved and opened:" & vbCr & copyPath
    Exit Sub
fail:
    ReportError "Clean for sending"
End Sub

' Saves "<name> (client).pptx" next to src, cleans it, leaves it open.
Public Function CleanCopy(ByVal src As Object) As String
    Dim full As String, nm As String, ext As String, copyPath As String, dot As Long
    full = src.FullName
    nm = src.Name
    dot = InStrRev(nm, ".")
    If dot > 0 Then ext = Mid$(nm, dot) Else ext = ".pptx"   ' keep the original format
    copyPath = Left$(full, Len(full) - Len(nm)) & Left$(nm, Len(nm) - IIf(dot > 0, Len(ext), 0)) & " (client)" & ext
    ' Leave text editing first: saving while a text cursor is active can stall.
    On Error Resume Next
    ActiveWindow.Selection.Unselect
    On Error GoTo 0
    src.SaveCopyAs copyPath
    Dim cp As Object
    Set cp = Presentations.Open(copyPath)
    CleanDeck cp
    cp.Save
    CleanCopy = copyPath
End Function

Public Sub CleanDeck(ByVal p As Object)
    Dim sld As Object
    RemoveTagged p, "STAMP"
    RemoveTagged p, "STICKY"
    Dim i As Long, kind As String
    For i = p.Slides.Count To 1 Step -1
        kind = TagOf(p.Slides(i), "ST_SLIDE")
        If kind = "STORYLINE" Or kind = "DECKCHECK" Then p.Slides(i).Delete
    Next i
    For Each sld In p.Slides
        SetNotesText sld, ""
        DeleteComments sld
    Next sld
End Sub

' Deletes every top-level shape tagged kind in presentation p. Returns count.
Public Function RemoveTagged(ByVal p As Object, ByVal kind As String) As Long
    Dim sld As Object, i As Long
    For Each sld In p.Slides
        For i = sld.Shapes.Count To 1 Step -1
            If TagOf(sld.Shapes(i), "ST_TYPE") = kind Then
                sld.Shapes(i).Delete
                RemoveTagged = RemoveTagged + 1
            End If
        Next i
    Next sld
End Function

' ---------------------------------------------------------------- font fixer

Public Sub ST_FixFonts()
    On Error GoTo fail
    If Presentations.Count = 0 Then
        Msg "Open a presentation first."
        Exit Sub
    End If
    Dim heading As String, body As String, stray As Collection, f As Variant, list As String
    heading = ActivePresentation.SlideMaster.Theme.ThemeFontScheme.MajorFont(1).Name
    body = ActivePresentation.SlideMaster.Theme.ThemeFontScheme.MinorFont(1).Name
    Set stray = StrayFonts(heading, body)
    If stray.Count = 0 Then
        Msg "All text already uses the template fonts (" & heading & " / " & body & ")."
        Exit Sub
    End If
    For Each f In stray
        list = list & vbCr & "  " & f
    Next f
    If Not Confirm("These fonts aren't the template's (" & heading & " / " & body & "):" & list & vbCr & vbCr & _
                   "Replace them with " & body & " everywhere, including masters and notes?") Then Exit Sub
    For Each f In stray
        ActivePresentation.Fonts.Replace CStr(f), body
    Next f
    Msg "Replaced " & stray.Count & " font(s) with " & body & "."
    Exit Sub
fail:
    ReportError "Fix fonts"
End Sub

Public Function StrayFonts(ByVal heading As String, ByVal body As String) As Collection
    Dim c As New Collection, i As Long, nm As String
    For i = 1 To ActivePresentation.Fonts.Count
        nm = ActivePresentation.Fonts(i).Name
        If StrComp(nm, heading, vbTextCompare) <> 0 And StrComp(nm, body, vbTextCompare) <> 0 _
           And Not IsSymbolFont(nm) And Left$(nm, 1) <> "+" Then c.Add nm
    Next i
    Set StrayFonts = c
End Function

' Bullet and symbol fonts are left alone.
Private Function IsSymbolFont(ByVal nm As String) As Boolean
    Select Case LCase$(nm)
        Case "wingdings", "wingdings 2", "wingdings 3", "symbol", "webdings", "zapf dingbats", "monotype sorts"
            IsSymbolFont = True
    End Select
End Function
