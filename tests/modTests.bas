Attribute VB_Name = "modTests"
' Developer tests. Only built into dist/StrategyTools-test.pptm, never the add-in.
' Each T_ function returns "PASS ..." or "FAIL ...".
Option Explicit

Private Function NewDeck(ByVal nSlides As Long, ByVal layout As Long) As Object
    Dim p As Object, i As Long
    Set p = Presentations.Add(ST_TRUE)
    For i = 1 To nSlides
        p.Slides.Add i, layout
    Next i
    p.Windows(1).Activate
    ActiveWindow.View.GotoSlide 1
    Set NewDeck = p
End Function

Private Sub CloseDeck(ByVal p As Object)
    On Error Resume Next
    p.Saved = ST_TRUE
    p.Close
End Sub

Private Function Near(ByVal a As Single, ByVal b As Single) As Boolean
    Near = Abs(a - b) < 0.6
End Function

Public Function T_Swap() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, a As Object, b As Object
    Set p = NewDeck(1, ppLayoutBlank)
    Set a = p.Slides(1).Shapes.AddShape(ST_SHAPE_RECTANGLE, 10, 10, 100, 50)
    Set b = p.Slides(1).Shapes.AddShape(ST_SHAPE_RECTANGLE, 300, 200, 40, 40)
    a.Select
    b.Select ST_FALSE
    ST_Swap
    ' A centre was (60,35), B centre was (320,220)
    If Near(a.Left + 50, 320) And Near(a.Top + 25, 220) And Near(b.Left + 20, 60) And Near(b.Top + 20, 35) _
       And Near(a.Width, 100) And Near(b.Width, 40) Then
        T_Swap = "PASS"
    Else
        T_Swap = "FAIL a=" & a.Left & "," & a.Top & " b=" & b.Left & "," & b.Top
    End If
    ' wrong selection count
    a.Select
    ST_Swap
    If InStr(gLastMessage, "exactly 2") = 0 Then T_Swap = T_Swap & " FAIL(no msg for 1 shape)"
    CloseDeck p
    Exit Function
fail:
    T_Swap = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_SameSize() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, a As Object, b As Object, c As Object, r As String
    Set p = NewDeck(1, ppLayoutBlank)
    Set a = p.Slides(1).Shapes.AddShape(ST_SHAPE_RECTANGLE, 10, 10, 100, 50)
    Set b = p.Slides(1).Shapes.AddShape(ST_SHAPE_RECTANGLE, 300, 200, 40, 30)
    Set c = p.Slides(1).Shapes.AddShape(ST_SHAPE_OVAL, 500, 100, 70, 90)
    c.LockAspectRatio = ST_TRUE
    ' select B first: B is the reference
    b.Select
    a.Select ST_FALSE
    c.Select ST_FALSE
    ST_SameSize
    If Near(a.Width, 40) And Near(a.Height, 30) And Near(c.Width, 40) And Near(c.Height, 30) _
       And Near(a.Left, 10) And Near(a.Top, 10) And Near(b.Width, 40) And c.LockAspectRatio = ST_TRUE Then
        r = "PASS size"
    Else
        r = "FAIL size a=" & a.Width & "x" & a.Height & " c=" & c.Width & "x" & c.Height
    End If
    ' width only, reference A (now 40x30) -> make A first after resizing
    a.Width = 120: a.Height = 60
    a.Select
    b.Select ST_FALSE
    ST_SameWidth
    If Near(b.Width, 120) And Near(b.Height, 30) Then r = r & " PASS width" Else r = r & " FAIL width " & b.Width & "x" & b.Height
    a.Select
    b.Select ST_FALSE
    ST_SameHeight
    If Near(b.Width, 120) And Near(b.Height, 60) Then r = r & " PASS height" Else r = r & " FAIL height " & b.Width & "x" & b.Height
    a.Select
    ST_SameSize
    If InStr(gLastMessage, "2 or more") = 0 Then r = r & " FAIL(no msg)"
    T_SameSize = r
    CloseDeck p
    Exit Function
fail:
    T_SameSize = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_Title() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, i As Long, r As String, origName As String, origSize As Single
    Set p = NewDeck(1, ppLayoutTitle)
    For i = 2 To 4
        p.Slides.Add i, ppLayoutText
    Next i
    p.Slides.Add 5, ppLayoutBlank          ' no title placeholder
    For i = 1 To 4
        p.Slides(i).Shapes.Title.TextFrame.TextRange.Text = "Title " & i
    Next i
    origName = p.Slides(1).Shapes.Title.TextFrame.TextRange.Font.Name
    origSize = p.Slides(1).Shapes.Title.TextFrame.TextRange.Font.Size
    With p.Slides(2).Shapes.Title.TextFrame.TextRange.Font
        .Name = "Georgia"
        .Size = 31
    End With
    ActiveWindow.View.GotoSlide 2
    p.Slides(2).Shapes.Title.Select
    ST_TitleHarmonizer
    r = "PASS"
    For i = 3 To 4
        With p.Slides(i).Shapes.Title.TextFrame.TextRange.Font
            If .Name <> "Georgia" Or .Size <> 31 Then r = "FAIL slide " & i & " " & .Name & " " & .Size
        End With
    Next i
    With p.Slides(1).Shapes.Title.TextFrame.TextRange.Font
        If .Name <> origName Or .Size <> origSize Then r = r & " FAIL title slide changed"
    End With
    If InStr(gLastMessage, "2 slide") = 0 Then r = r & " FAIL msg=" & gLastMessage
    ' text cursor inside the title also works
    p.Slides(3).Shapes.Title.TextFrame.TextRange.Font.Size = 20
    ActiveWindow.View.GotoSlide 3
    p.Slides(3).Shapes.Title.TextFrame.TextRange.Characters(2, 0).Select
    ST_TitleHarmonizer
    If p.Slides(4).Shapes.Title.TextFrame.TextRange.Font.Size <> 20 Then r = r & " FAIL via text cursor"
    ' non-title selection gives a message
    p.Slides(3).Shapes(2).Select
    ST_TitleHarmonizer
    If InStr(gLastMessage, "Click into a slide title") = 0 Then r = r & " FAIL(no msg)"
    T_Title = r
    CloseDeck p
    Exit Function
fail:
    T_Title = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_Visuals() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, r As String, sld As Object, g As Object, n As Long
    Set p = NewDeck(1, ppLayoutBlank)
    Set sld = p.Slides(1)
    ActiveWindow.Selection.Unselect
    ST_Harvey50
    If sld.Shapes.Count <> 1 Then r = "FAIL harvey count " & sld.Shapes.Count: GoTo done
    Set g = sld.Shapes(1)
    If g.Type <> ST_TYPE_GROUP Or g.Tags("ST_TYPE") <> "HARVEY" Or g.Tags("ST_LEVEL") <> "50" Then r = "FAIL harvey tags": GoTo done
    If Not Near(g.Width, 28.8) Then r = "FAIL harvey size " & g.Width: GoTo done
    ' update in place (select group, pick 75)
    g.Select
    ST_Harvey75
    If sld.Shapes.Count <> 1 Or sld.Shapes(1).Tags("ST_LEVEL") <> "75" Then r = "FAIL harvey update": GoTo done
    ActiveWindow.Selection.Unselect
    ST_MeterMedium
    If sld.Shapes.Count <> 2 Then r = "FAIL meter count": GoTo done
    Set g = sld.Shapes(2)
    If g.Tags("ST_TYPE") <> "METER" Or g.Tags("ST_LEVEL") <> "2" Then r = "FAIL meter tags": GoTo done
    g.Select
    ST_MeterHigh
    If sld.Shapes.Count <> 2 Or sld.Shapes(2).Tags("ST_LEVEL") <> "3" Then r = "FAIL meter update": GoTo done
    ' selecting a non-visual shape adds a new visual rather than failing
    Dim rect As Object
    Set rect = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, 10, 10, 50, 50)
    rect.Select
    ST_Harvey25
    If sld.Shapes.Count <> 4 Then r = "FAIL harvey with rect selected": GoTo done
    r = "PASS"
done:
    T_Visuals = r
    CloseDeck p
    Exit Function
fail:
    T_Visuals = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_TextBox() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, tb As Object
    Set p = NewDeck(1, ppLayoutBlank)
    ST_AddTextBox
    Set tb = p.Slides(1).Shapes(1)
    If tb.HasTextFrame = ST_TRUE And Near(tb.Left + tb.Width / 2, p.PageSetup.SlideWidth / 2) _
       And ActiveWindow.Selection.Type = ppSelectionText Then
        T_TextBox = "PASS"
    Else
        T_TextBox = "FAIL sel=" & ActiveWindow.Selection.Type
    End If
    CloseDeck p
    Exit Function
fail:
    T_TextBox = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_NoWindow() As String
    ' Messages instead of errors when nothing is selected
    On Error GoTo fail
    gQuiet = True
    Dim p As Object
    Set p = NewDeck(1, ppLayoutBlank)
    ActiveWindow.Selection.Unselect
    ST_SameWidth
    If InStr(gLastMessage, "2 or more") = 0 Then T_NoWindow = "FAIL samewidth": CloseDeck p: Exit Function
    ST_TitleHarmonizer
    If InStr(gLastMessage, "title") = 0 Then T_NoWindow = "FAIL title": CloseDeck p: Exit Function
    T_NoWindow = "PASS"
    CloseDeck p
    Exit Function
fail:
    T_NoWindow = "FAIL error " & Err.Description
    CloseDeck p
End Function

' Renders every visual on one slide and exports a PNG for a visual check.
Public Function T_Render() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, sld As Object, levels As Variant, i As Long, path As String
    Set p = NewDeck(1, ppLayoutBlank)
    Set sld = p.Slides(1)
    levels = Array(0, 25, 50, 75, 100)
    For i = 0 To 4
        ActiveWindow.Selection.Unselect
        HarveyBall levels(i)
        With sld.Shapes(sld.Shapes.Count)
            .Width = 80: .Height = 80
            .Left = 60 + i * 170: .Top = 120
        End With
    Next i
    For i = 1 To 3
        ActiveWindow.Selection.Unselect
        Meter i
        With sld.Shapes(sld.Shapes.Count)
            .LockAspectRatio = ST_TRUE
            .Height = 80
            .Left = 200 + (i - 1) * 200: .Top = 320
        End With
    Next i
    path = Environ("TMPDIR") & "st_render.png"
    sld.Export path, "PNG", 1280, 720
    T_Render = "PASS " & path
    CloseDeck p
    Exit Function
fail:
    T_Render = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_HarveyDetail() As String
    On Error GoTo fail
    gQuiet = True
    gLastMessage = ""
    Dim p As Object, g As Object, it As Object, r As String
    Set p = NewDeck(1, ppLayoutBlank)
    p.Slides(1).Shapes.AddShape ST_SHAPE_RECTANGLE, 5, 5, 20, 20
    ActiveWindow.Selection.Unselect
    ST_Harvey25
    Set g = p.Slides(1).Shapes(p.Slides(1).Shapes.Count)
    r = "msg=[" & gLastMessage & "] items=" & g.GroupItems.Count
    For Each it In g.GroupItems
        r = r & " | " & it.Name & " vis=" & it.Visible & " line=" & it.Line.Visible
        If it.Name = "ST_Fill" Then r = r & " adj=" & it.Adjustments(1) & "," & it.Adjustments(2)
    Next it
    T_HarveyDetail = r
    CloseDeck p
    Exit Function
fail:
    T_HarveyDetail = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_Status() As String
    On Error GoTo fail
    gQuiet = True
    gLastMessage = ""
    Dim p As Object, sld As Object, g As Object, r As String, x As Single
    Set p = NewDeck(1, ppLayoutBlank)
    Set sld = p.Slides(1)
    ActiveWindow.Selection.Unselect
    ST_StatusComplete
    If sld.Shapes.Count <> 1 Then r = "FAIL count " & sld.Shapes.Count & " msg=" & gLastMessage: GoTo done
    Set g = sld.Shapes(1)
    If g.Tags("ST_TYPE") <> "STATUS" Or g.Tags("ST_LEVEL") <> "COMPLETE" Or g.GroupItems.Count <> 2 Then r = "FAIL complete": GoTo done
    If Not Near(g.Width, 28.8) Then r = "FAIL size " & g.Width: GoTo done
    ' move it, then change status in place: position kept, still one shape
    g.Left = 100: g.Top = 50
    g.Select
    ST_StatusDelayed
    Set g = sld.Shapes(1)
    If sld.Shapes.Count <> 1 Or g.Tags("ST_LEVEL") <> "DELAYED" Or g.GroupItems.Count <> 3 Then r = "FAIL update " & sld.Shapes.Count: GoTo done
    If Not (Near(g.Left, 100) And Near(g.Top, 50)) Then r = "FAIL update moved it " & g.Left & "," & g.Top: GoTo done
    ActiveWindow.Selection.Unselect
    ST_StatusInProgress
    ActiveWindow.Selection.Unselect
    ST_StatusNotStarted
    If sld.Shapes.Count <> 3 Then r = "FAIL add more": GoTo done
    If sld.Shapes(2).Tags("ST_LEVEL") <> "INPROGRESS" Or sld.Shapes(2).GroupItems.Count < 5 Then r = "FAIL dotted tick": GoTo done
    If gLastMessage <> "" Then r = "FAIL msg=" & gLastMessage: GoTo done
    r = "PASS"
done:
    T_Status = r
    CloseDeck p
    Exit Function
fail:
    T_Status = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_Chevrons() As String
    On Error GoTo fail
    gQuiet = True
    gLastMessage = ""
    Dim p As Object, sld As Object, g As Object, r As String, i As Long, W As Single
    Set p = NewDeck(1, ppLayoutTitleOnly)
    Set sld = p.Slides(1)
    W = p.PageSetup.SlideWidth
    AddChevrons 4
    Set g = sld.Shapes(sld.Shapes.Count)
    If g.Tags("ST_TYPE") <> "CHEVRONS" Or g.GroupItems.Count <> 8 Then r = "FAIL group msg=" & gLastMessage: GoTo done
    If g.Left < 0 Or g.Left + g.Width > W + 0.5 Then r = "FAIL off slide": GoTo done
    If g.Top + g.Height > p.PageSetup.SlideHeight Then r = "FAIL too tall": GoTo done
    ' aligned with the title
    If Not Near(g.Left, sld.Shapes.Title.Left) Then r = "FAIL not aligned to title " & g.Left & " vs " & sld.Shapes.Title.Left: GoTo done
    ' chevrons left-to-right, details under them
    Dim c1 As Object, c2 As Object, d1 As Object
    For i = 1 To g.GroupItems.Count
        If g.GroupItems(i).Name = "ST_Chevron1" Then Set c1 = g.GroupItems(i)
        If g.GroupItems(i).Name = "ST_Chevron2" Then Set c2 = g.GroupItems(i)
        If g.GroupItems(i).Name = "ST_Detail1" Then Set d1 = g.GroupItems(i)
    Next i
    If c1 Is Nothing Or c2 Is Nothing Or d1 Is Nothing Then r = "FAIL names": GoTo done
    If c2.Left <= c1.Left Then r = "FAIL order": GoTo done
    If d1.Top < c1.Top + c1.Height Then r = "FAIL detail not below": GoTo done
    If c1.TextFrame.TextRange.Text <> "Phase 1" Then r = "FAIL label": GoTo done
    ' out of range gives a message
    AddChevrons 12
    If InStr(gLastMessage, "2 to 8") = 0 And InStr(gLastMessage, "from 2") = 0 Then r = "FAIL range msg=" & gLastMessage: GoTo done
    r = "PASS"
done:
    T_Chevrons = r
    CloseDeck p
    Exit Function
fail:
    T_Chevrons = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_Trackers() As String
    On Error GoTo fail
    gQuiet = True
    gLastMessage = ""
    Dim p As Object, sld As Object, r As String, s As String, i As Long, shp As Object
    Set p = NewDeck(1, ppLayoutBlank)
    Set sld = p.Slides(1)
    ST_TrackerNumber
    ST_TrackerNumber
    ST_TrackerNumber
    ST_TrackerLetter
    ST_TrackerLetter
    AddTracker "CUSTOM", "2a"
    For Each shp In sld.Shapes
        s = s & shp.TextFrame.TextRange.Text & ","
    Next shp
    If s <> "1,2,3,A,B,2a," Then r = "FAIL texts=" & s & " msg=" & gLastMessage: GoTo done
    ' no two trackers in the same spot
    For i = 2 To 3
        If Near(sld.Shapes(i).Top, sld.Shapes(i - 1).Top) And Near(sld.Shapes(i).Left, sld.Shapes(i - 1).Left) Then r = "FAIL stacked": GoTo done
    Next i
    ' deleting 2 and adding again continues from the highest (4)
    sld.Shapes(2).Delete
    ST_TrackerNumber
    If sld.Shapes(sld.Shapes.Count).TextFrame.TextRange.Text <> "4" Then r = "FAIL next after delete": GoTo done
    r = "PASS"
done:
    T_Trackers = r
    CloseDeck p
    Exit Function
fail:
    T_Trackers = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

' Renders the new visuals for a visual check.
Public Function T_Render2() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, sld As Object, i As Long, path As String, st As Variant
    Set p = NewDeck(1, ppLayoutTitleOnly)
    Set sld = p.Slides(1)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Project roadmap"
    AddChevrons 4
    st = Array("COMPLETE", "INPROGRESS", "DELAYED", "NOTSTARTED")
    For i = 0 To 3
        ActiveWindow.Selection.Unselect
        StatusIcon st(i)
        With sld.Shapes(sld.Shapes.Count)
            .LockAspectRatio = ST_TRUE
            .Width = 40
            .Left = 120 + i * 70: .Top = 440
        End With
    Next i
    For i = 1 To 3
        ST_TrackerNumber
    Next i
    ST_TrackerLetter
    ST_TrackerLetter
    path = Environ("TMPDIR") & "st_render2.png"
    sld.Export path, "PNG", 1280, 720
    T_Render2 = "PASS " & path
    CloseDeck p
    Exit Function
fail:
    T_Render2 = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_TrackerOverlap() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, sld As Object, i As Long, j As Long
    Set p = NewDeck(1, ppLayoutBlank)
    Set sld = p.Slides(1)
    ST_TrackerNumber
    ST_TrackerNumber
    ST_TrackerLetter
    ST_TrackerLetter
    AddTracker "CUSTOM", "x"
    For i = 1 To sld.Shapes.Count
        For j = i + 1 To sld.Shapes.Count
            If Abs(sld.Shapes(i).Left - sld.Shapes(j).Left) < 21 And Abs(sld.Shapes(i).Top - sld.Shapes(j).Top) < 21 Then
                T_TrackerOverlap = "FAIL overlap " & sld.Shapes(i).TextFrame.TextRange.Text & "/" & sld.Shapes(j).TextFrame.TextRange.Text
                CloseDeck p
                Exit Function
            End If
        Next j
    Next i
    T_TrackerOverlap = "PASS"
    CloseDeck p
    Exit Function
fail:
    T_TrackerOverlap = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_Stamps() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, r As String, i As Long
    Set p = NewDeck(3, ppLayoutBlank)
    ST_StampDraft
    If TaggedShapes(p.Slides(1), "STAMP").Count <> 1 Then r = "FAIL add": GoTo done
    ST_StampPreliminary
    If TaggedShapes(p.Slides(1), "STAMP").Count <> 1 Then r = "FAIL relabel duplicated": GoTo done
    If TaggedShapes(p.Slides(1), "STAMP")(1).TextFrame.TextRange.Text <> "PRELIMINARY" Then r = "FAIL relabel text": GoTo done
    ST_StampAllSlides
    For i = 1 To 3
        If TaggedShapes(p.Slides(i), "STAMP").Count <> 1 Then r = "FAIL all slides " & i: GoTo done
    Next i
    ST_StampRemoveAll
    For i = 1 To 3
        If TaggedShapes(p.Slides(i), "STAMP").Count <> 0 Then r = "FAIL remove " & i: GoTo done
    Next i
    r = "PASS"
done:
    T_Stamps = r
    CloseDeck p
    Exit Function
fail:
    T_Stamps = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_StickyFooter() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, r As String, f As Object, H As Single
    Set p = NewDeck(2, ppLayoutTitleOnly)
    H = p.PageSetup.SlideHeight
    ST_StickyAdd
    If TaggedShapes(p.Slides(1), "STICKY").Count <> 1 Then r = "FAIL sticky add " & gLastMessage: GoTo done
    ST_StickyRemoveAll
    If TaggedShapes(p.Slides(1), "STICKY").Count <> 0 Then r = "FAIL sticky remove": GoTo done
    AddFooter True
    AddFooter True
    If TaggedShapes(p.Slides(1), "FOOTER").Count <> 1 Then r = "FAIL footer count": GoTo done
    Set f = TaggedShapes(p.Slides(1), "FOOTER")(1)
    If Left$(f.TextFrame.TextRange.Text, 6) <> "Note: " Then r = "FAIL footer text " & f.TextFrame.TextRange.Text: GoTo done
    If Abs(f.Top + f.Height - H * 0.94) > 1 Then r = "FAIL footer bottom " & (f.Top + f.Height): GoTo done
    If Abs(f.Left - p.Slides(1).Shapes.Title.Left) > 1 Then r = "FAIL footer left": GoTo done
    ActiveWindow.View.GotoSlide 2
    AddFooter False
    Dim f2 As Object
    Set f2 = TaggedShapes(p.Slides(2), "FOOTER")(1)
    If Abs((f2.Top + f2.Height) - (f.Top + f.Height)) > 1 Or Abs(f2.Left - f.Left) > 0.5 Then r = "FAIL footer not same spot": GoTo done
    r = "PASS"
done:
    T_StickyFooter = r
    CloseDeck p
    Exit Function
fail:
    T_StickyFooter = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_Backup() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, r As String
    Set p = NewDeck(3, ppLayoutTitleOnly)
    ActiveWindow.View.GotoSlide 3
    ST_BackupMark                      ' quiet: answers Yes to the appendix question
    If p.Slides.Count <> 4 Then r = "FAIL appendix not inserted (" & p.Slides.Count & ")": GoTo done
    If p.Slides(3).Shapes.Title.TextFrame.TextRange.Text <> "Appendix" Then r = "FAIL appendix title": GoTo done
    If TaggedShapes(p.Slides(4), "BACKUP").Count <> 1 Then r = "FAIL label": GoTo done
    ActiveWindow.View.GotoSlide 4
    ST_BackupMark
    If TaggedShapes(p.Slides(4), "BACKUP").Count <> 1 Then r = "FAIL label duplicated": GoTo done
    ST_BackupRemove
    If TaggedShapes(p.Slides(ActiveWindow.View.Slide.SlideIndex), "BACKUP").Count <> 0 Then r = "FAIL remove": GoTo done
    r = "PASS"
done:
    T_Backup = r
    CloseDeck p
    Exit Function
fail:
    T_Backup = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_Clean() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, r As String, path As String, cp As Object, shp As Object
    Set p = NewDeck(2, ppLayoutTitleOnly)
    path = Environ("TMPDIR") & "st_clean_test.pptx"
    On Error Resume Next
    Kill path
    Kill Environ("TMPDIR") & "st_clean_test (client).pptx"
    On Error GoTo fail
    p.SaveAs path, ppSaveAsOpenXMLPresentation
    ST_StampDraft
    ST_StickyAdd
    SetNotesText p.Slides(1), "speaker note"
    If NotesText(p.Slides(1)) <> "speaker note" Then r = "FAIL could not set notes": GoTo done
    Dim copyPath As String
    copyPath = CleanCopy(p)
    Set cp = ActivePresentation
    If cp.FullName = p.FullName Then r = "FAIL copy not opened": GoTo done
    If TaggedShapes(cp.Slides(1), "STAMP").Count <> 0 Or TaggedShapes(cp.Slides(1), "STICKY").Count <> 0 Then r = "FAIL copy not clean": GoTo done
    If Len(NotesText(cp.Slides(1))) > 0 Then r = "FAIL notes kept": GoTo done
    If TaggedShapes(p.Slides(1), "STAMP").Count <> 1 Then r = "FAIL original changed": GoTo done
    r = "PASS " & copyPath
done:
    T_Clean = r
    If Not cp Is Nothing Then CloseDeck cp
    CloseDeck p
    Exit Function
fail:
    T_Clean = "FAIL error " & Err.Description & " msg=" & gLastMessage
    If Not cp Is Nothing Then CloseDeck cp
    CloseDeck p
End Function

Public Function T_FontsCornersMatch() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, r As String, tb As Object, body As String, a As Object, b As Object
    Set p = NewDeck(1, ppLayoutBlank)
    body = p.SlideMaster.Theme.ThemeFontScheme.MinorFont(1).Name
    Set tb = p.Slides(1).Shapes.AddTextbox(ST_TEXT_HORIZONTAL, 10, 10, 200, 30)
    tb.TextFrame.TextRange.Text = "Stray font"
    tb.TextFrame.TextRange.Font.Name = "Courier New"
    ST_FixFonts
    If tb.TextFrame.TextRange.Font.Name <> body Then r = "FAIL fonts: " & tb.TextFrame.TextRange.Font.Name & " msg=" & gLastMessage: GoTo done
    ' corners: radius 10pt on a 100x50 box carried to a 200x150 box
    Set a = p.Slides(1).Shapes.AddShape(5, 50, 100, 100, 50)
    a.Adjustments(1) = 0.2
    Set b = p.Slides(1).Shapes.AddShape(5, 300, 100, 200, 150)
    a.Select
    b.Select ST_FALSE
    ST_SameCorners
    If Abs(b.Adjustments(1) * 150 - 10) > 0.5 Then r = "FAIL corners " & b.Adjustments(1): GoTo done
    ' match format: fill + font copied, size kept
    a.Fill.ForeColor.RGB = RGB(200, 30, 30)
    a.TextFrame.TextRange.Text = "ref"
    a.TextFrame.TextRange.Font.Name = "Georgia"
    b.TextFrame.TextRange.Text = "target"
    a.Select
    b.Select ST_FALSE
    ST_MatchFormat
    If b.Fill.ForeColor.RGB <> RGB(200, 30, 30) Then r = "FAIL match fill": GoTo done
    If b.TextFrame.TextRange.Font.Name <> "Georgia" Then r = "FAIL match font": GoTo done
    If Abs(b.Width - 200) > 0.5 Or Abs(b.Height - 150) > 0.5 Then r = "FAIL match changed size": GoTo done
    r = "PASS"
done:
    T_FontsCornersMatch = r
    CloseDeck p
    Exit Function
fail:
    T_FontsCornersMatch = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_GetStep() As String
    T_GetStep = gLastMessage
End Function

Public Function T_CleanSteps() As String
    gQuiet = True
    Dim p As Object, path As String
    gLastMessage = "1 new deck"
    Set p = NewDeck(1, ppLayoutTitleOnly)
    path = Environ("TMPDIR") & "st_step.pptx"
    gLastMessage = "2 saveas"
    p.SaveAs path, ppSaveAsOpenXMLPresentation
    gLastMessage = "3 stamp"
    ST_StampDraft
    gLastMessage = "4 sticky"
    ST_StickyAdd
    gLastMessage = "5 notes"
    SetNotesText p.Slides(1), "speaker note"
    gLastMessage = "6 savecopyas"
    p.SaveCopyAs Environ("TMPDIR") & "st_step (client).pptx"
    gLastMessage = "7 open copy"
    Dim cp As Object
    Set cp = Presentations.Open(Environ("TMPDIR") & "st_step (client).pptx")
    gLastMessage = "8 clean"
    CleanDeck cp
    gLastMessage = "9 save copy"
    cp.Save
    gLastMessage = "10 done"
    CloseDeck cp
    CloseDeck p
    T_CleanSteps = "done"
End Function

Private Sub SetItem(ByVal m As Object, ByVal i As Long, ByVal text As String)
    Dim shp As Object
    For Each shp In m.Shapes
        If TagOf(shp, "ST_TYPE") = "AGENDA_ITEM" And TagOf(shp, "ST_INDEX") = CStr(i) Then shp.TextFrame.TextRange.Text = text
    Next shp
End Sub

Private Function TitledDeck(ByVal n As Long) As Object
    Dim p As Object, i As Long
    Set p = NewDeck(n, ppLayoutTitleOnly)
    For i = 1 To n
        p.Slides(i).Shapes.Title.TextFrame.TextRange.Text = "Title " & i
    Next i
    Set TitledDeck = p
End Function

Public Function T_Agenda() As String
    On Error GoTo fail
    gQuiet = True
    gLastMessage = ""
    Dim p As Object, r As String, m As Object, d As Object, i As Long, tag As Collection
    Set p = TitledDeck(5)              ' Title 1..5
    ActiveWindow.View.GotoSlide 1
    Set m = AddAgenda(3)
    If m Is Nothing Then r = "FAIL no agenda " & gLastMessage: GoTo done
    If m.SlideIndex <> 2 Or SectionCount() <> 3 Then r = "FAIL agenda slot/count": GoTo done
    SetItem m, 1, "Alpha": SetItem m, 2, "Beta": SetItem m, 3, "Gamma"
    ' deck now: 1 T1, 2 Agenda, 3 T2, 4 T3, 5 T4, 6 T5
    ActiveWindow.View.GotoSlide 3
    ST_AgendaDivider                   ' before T2 -> section 1
    ActiveWindow.View.GotoSlide 6      ' T4 (was 5, +1 after divider)
    ST_AgendaDivider                   ' section 2
    Set d = p.Slides(3)
    If TagOf(d, "ST_SLIDE") <> "DIVIDER" Or TagOf(d, "ST_SECTION") <> "1" Then r = "FAIL divider 1 " & gLastMessage: GoTo done
    If TagOf(p.Slides(6), "ST_SECTION") <> "2" Then r = "FAIL divider 2 at " & TagOf(p.Slides(6), "ST_SLIDE"): GoTo done
    ' highlight band visible only on the current row
    Dim shp As Object, visibleBands As String
    For Each shp In p.Slides(6).Shapes
        If TagOf(shp, "ST_TYPE") = "AGENDA_BAND" And shp.Visible = ST_TRUE Then visibleBands = visibleBands & TagOf(shp, "ST_INDEX")
    Next shp
    If visibleBands <> "2" Then r = "FAIL band " & visibleBands: GoTo done
    ' tags: T2,T3 -> section 1 ; T4,T5 -> section 2 ; T1 untagged
    ApplySectionTags
    If TaggedShapes(p.Slides(1), "SECTIONTAG").Count <> 0 Then r = "FAIL T1 tagged": GoTo done
    Set tag = TaggedShapes(p.Slides(4), "SECTIONTAG")
    If tag.Count <> 1 Then r = "FAIL T2 tag count": GoTo done
    If InStr(tag(1).TextFrame.TextRange.Text, "ALPHA") = 0 Then r = "FAIL T2 tag text " & tag(1).TextFrame.TextRange.Text: GoTo done
    If InStr(TaggedShapes(p.Slides(8), "SECTIONTAG")(1).TextFrame.TextRange.Text, "BETA") = 0 Then r = "FAIL T5 tag": GoTo done
    ' stamp sits left of the tag
    ActiveWindow.View.GotoSlide 4
    ST_StampDraft
    If TaggedShapes(p.Slides(4), "STAMP")(1).Left + TaggedShapes(p.Slides(4), "STAMP")(1).Width > TaggedShapes(p.Slides(4), "SECTIONTAG")(1).Left Then r = "FAIL stamp overlaps tag": GoTo done
    ' rename + update keeps dividers in place and refreshes names
    SetItem m, 1, "Market"
    ST_AgendaUpdate
    If TagOf(p.Slides(3), "ST_SLIDE") <> "DIVIDER" Or TagOf(p.Slides(6), "ST_SLIDE") <> "DIVIDER" Then r = "FAIL update moved dividers": GoTo done
    If InStr(TaggedShapes(p.Slides(4), "SECTIONTAG")(1).TextFrame.TextRange.Text, "MARKET") = 0 Then r = "FAIL update tag text": GoTo done
    Dim nameOnDivider As String
    For Each shp In p.Slides(3).Shapes
        If TagOf(shp, "ST_TYPE") = "AGENDA_ITEM" And TagOf(shp, "ST_INDEX") = "1" Then nameOnDivider = shp.TextFrame.TextRange.Text
    Next shp
    If nameOnDivider <> "Market" Then r = "FAIL divider name " & nameOnDivider: GoTo done
    If p.Slides.Count <> 8 Then r = "FAIL slide count " & p.Slides.Count: GoTo done
    ' all sections used: a 4th divider is refused (only 3 sections, but 2 used so 3rd ok)
    ActiveWindow.View.GotoSlide 8
    ST_AgendaDivider
    ActiveWindow.View.GotoSlide 9
    ST_AgendaDivider
    If InStr(gLastMessage, "already has a divider") = 0 Then r = "FAIL too many dividers msg=" & gLastMessage: GoTo done
    r = "PASS"
done:
    T_Agenda = r
    CloseDeck p
    Exit Function
fail:
    T_Agenda = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_StorylineCheck() As String
    On Error GoTo fail
    gQuiet = True
    gLastMessage = ""
    Dim p As Object, r As String, txt As String, last As Object, tb As Object
    Set p = TitledDeck(4)
    ' issues: slide 2 no title, slide 3 shape off slide + overflow, slide 4 stray font + sticky
    p.Slides(2).Shapes.Title.TextFrame.TextRange.Text = ""
    p.Slides(3).Shapes.AddShape ST_SHAPE_RECTANGLE, p.PageSetup.SlideWidth - 20, 100, 80, 40
    Set tb = p.Slides(3).Shapes.AddTextbox(ST_TEXT_HORIZONTAL, 50, 200, 120, 20)
    tb.TextFrame.AutoSize = ppAutoSizeNone
    tb.TextFrame.WordWrap = ST_TRUE
    tb.Height = 20
    tb.TextFrame.TextRange.Text = "A long line of text that will certainly wrap onto several lines in this narrow box and overflow it"
    Set tb = p.Slides(4).Shapes.AddTextbox(ST_TEXT_HORIZONTAL, 50, 300, 200, 30)
    tb.TextFrame.TextRange.Text = "odd font"
    tb.TextFrame.TextRange.Font.Name = "Courier New"
    ActiveWindow.View.GotoSlide 4
    ST_StickyAdd
    ' storyline
    ST_Storyline
    ST_Storyline                        ' twice -> still one storyline slide
    If ReportSlideCount("STORYLINE") <> 1 Then r = "FAIL storyline count " & ReportSlideCount("STORYLINE"): GoTo done
    Set last = p.Slides(p.Slides.Count)
    txt = last.Shapes(last.Shapes.Count).TextFrame.TextRange.Text
    If InStr(txt, "Title 3") = 0 Or InStr(txt, "(no title)") = 0 Then r = "FAIL storyline text " & txt: GoTo done
    ' deck check
    ST_DeckCheck
    If ReportSlideCount("DECKCHECK") <> 1 Then r = "FAIL deckcheck slide " & gLastMessage: GoTo done
    Set last = p.Slides(p.Slides.Count)
    txt = last.Shapes(last.Shapes.Count).TextFrame.TextRange.Text
    If InStr(txt, "Slide 2") = 0 Or InStr(txt, "No title") = 0 Then r = "FAIL no-title " & txt: GoTo done
    If InStr(txt, "off the slide") = 0 Then r = "FAIL off-slide " & txt: GoTo done
    If InStr(txt, "overflows") = 0 Then r = "FAIL overflow " & txt: GoTo done
    If InStr(txt, "Courier New") = 0 Then r = "FAIL font " & txt: GoTo done
    If InStr(txt, "Sticky") = 0 Then r = "FAIL sticky " & txt: GoTo done
    If InStr(txt, "Storyline") > 0 Then r = "FAIL report checked itself": GoTo done
    ' clean removes both report slides
    CleanDeck p
    If ReportSlideCount("STORYLINE") + ReportSlideCount("DECKCHECK") <> 0 Then r = "FAIL clean kept reports": GoTo done
    r = "PASS"
done:
    T_StorylineCheck = r
    CloseDeck p
    Exit Function
fail:
    T_StorylineCheck = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_Render3() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, m As Object, d As String, i As Long
    Set p = TitledDeck(4)
    p.Slides(2).Shapes.Title.TextFrame.TextRange.Text = "Two-thirds of customers would switch for a better digital experience"
    ActiveWindow.View.GotoSlide 1
    Set m = AddAgenda(4)
    SetItem m, 1, "Market overview": SetItem m, 2, "Customer insights": SetItem m, 3, "Operating model": SetItem m, 4, "Next steps"
    ActiveWindow.View.GotoSlide 3
    ST_AgendaDivider
    ActiveWindow.View.GotoSlide 5
    ST_AgendaDivider
    ApplySectionTags
    ActiveWindow.View.GotoSlide 6
    ST_StampDraft
    ST_Storyline
    d = Environ("TMPDIR")
    p.Slides(2).Export d & "st_r3_agenda.png", "PNG", 1280, 720
    p.Slides(5).Export d & "st_r3_divider.png", "PNG", 1280, 720
    p.Slides(6).Export d & "st_r3_tagged.png", "PNG", 1280, 720
    p.Slides(p.Slides.Count).Export d & "st_r3_storyline.png", "PNG", 1280, 720
    T_Render3 = "PASS " & d
    CloseDeck p
    Exit Function
fail:
    T_Render3 = "FAIL error " & Err.Description
    CloseDeck p
End Function

Private Function InsideSlide(ByVal shp As Object, ByVal p As Object) As Boolean
    InsideSlide = shp.Left >= -1 And shp.Top >= -1 And shp.Left + shp.Width <= p.PageSetup.SlideWidth + 1 _
                  And shp.Top + shp.Height <= p.PageSetup.SlideHeight + 1
End Function

Public Function T_Frameworks() As String
    On Error GoTo fail
    gQuiet = True
    gLastMessage = ""
    Dim p As Object, r As String, g As Object, a As Object, b As Object, sld As Object
    Set p = TitledDeck(5)
    ' matrix
    ActiveWindow.View.GotoSlide 1
    Set g = AddMatrix()
    If g.Tags("ST_TYPE") <> "MATRIX" Or g.GroupItems.Count <> 8 Then r = "FAIL matrix parts " & g.GroupItems.Count: GoTo done
    If Not InsideSlide(g, p) Then r = "FAIL matrix off slide": GoTo done
    If g.Top < p.Slides(1).Shapes.Title.Top + p.Slides(1).Shapes.Title.Height Then r = "FAIL matrix over title": GoTo done
    ' process
    ActiveWindow.View.GotoSlide 2
    Set g = AddProcess(4)
    If g.GroupItems.Count <> 9 Then r = "FAIL process parts " & g.GroupItems.Count: GoTo done
    If Not InsideSlide(g, p) Then r = "FAIL process off slide": GoTo done
    If AddProcess(9) Is Nothing = False Then r = "FAIL process range": GoTo done
    ' takeaways (second call selects instead of duplicating)
    ActiveWindow.View.GotoSlide 3
    Set g = AddTakeaways()
    AddTakeaways
    If TaggedShapes(p.Slides(3), "TAKEAWAY").Count <> 1 Then r = "FAIL takeaway duplicated": GoTo done
    If g.Left < p.PageSetup.SlideWidth / 2 Or Not InsideSlide(g, p) Then r = "FAIL takeaway position": GoTo done
    ' rating: add, then update in place
    ActiveWindow.View.GotoSlide 4
    Set sld = p.Slides(4)
    ActiveWindow.Selection.Unselect
    ST_Rating3
    Set g = TaggedShapes(sld, "RATING")(1)
    If g.Tags("ST_LEVEL") <> "3" Then r = "FAIL rating level": GoTo done
    g.Select
    ST_Rating5
    If TaggedShapes(sld, "RATING").Count <> 1 Or TaggedShapes(sld, "RATING")(1).Tags("ST_LEVEL") <> "5" Then r = "FAIL rating update": GoTo done
    ' CAGR between two bars: 100 -> 200 over 3 years = +26% p.a.
    ActiveWindow.View.GotoSlide 5
    Set a = p.Slides(5).Shapes.AddShape(ST_SHAPE_RECTANGLE, 400, 300, 40, 100)
    Set b = p.Slides(5).Shapes.AddShape(ST_SHAPE_RECTANGLE, 150, 250, 40, 150)
    a.Select
    b.Select ST_FALSE
    Set g = AddCagr(100, 200, 3)
    If g.Tags("ST_TEXT") <> "+26% p.a." Then r = "FAIL cagr label " & g.Tags("ST_TEXT"): GoTo done
    If Abs(g.Left - 170) > 3 Or Abs(g.Left + g.Width - 420) > 3 Then r = "FAIL cagr span " & g.Left & "-" & (g.Left + g.Width): GoTo done
    If CagrLabel(100, 103, 1) <> "+3.0% p.a." Or CagrLabel(200, 100, 2) <> "-29% p.a." Then r = "FAIL cagr format " & CagrLabel(200, 100, 2): GoTo done
    ActiveWindow.Selection.Unselect
    If AddCagr(100, 150, 2) Is Nothing Then r = "FAIL cagr no selection": GoTo done
    If Not (AddCagr(0, 150, 2) Is Nothing) Then r = "FAIL cagr zero accepted": GoTo done
    r = "PASS"
done:
    T_Frameworks = r
    CloseDeck p
    Exit Function
fail:
    T_Frameworks = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_Render4() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, d As String, a As Object, b As Object, i As Long
    Set p = TitledDeck(3)
    p.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Prioritising initiatives by impact and ease"
    p.Slides(2).Shapes.Title.TextFrame.TextRange.Text = "Four-step approach to the transformation"
    p.Slides(3).Shapes.Title.TextFrame.TextRange.Text = "Revenue grew strongly while the market stalled"
    ActiveWindow.View.GotoSlide 1
    AddMatrix
    ActiveWindow.View.GotoSlide 2
    AddProcess 4
    ActiveWindow.View.GotoSlide 3
    For i = 0 To 3
        Set a = p.Slides(3).Shapes.AddShape(ST_SHAPE_RECTANGLE, 80 + i * 110, 470 - (150 + i * 50), 60, 150 + i * 50)
        a.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
        If i = 0 Then Set b = a
    Next i
    b.Select
    a.Select ST_FALSE
    AddCagr 150, 300, 3
    AddTakeaways
    ActiveWindow.Selection.Unselect
    ST_Rating4
    With TaggedShapes(p.Slides(3), "RATING")(1)
        .Left = 120: .Top = 490
    End With
    d = Environ("TMPDIR")
    For i = 1 To 3
        p.Slides(i).Export d & "st_r4_" & i & ".png", "PNG", 1280, 720
    Next i
    T_Render4 = "PASS"
    CloseDeck p
    Exit Function
fail:
    T_Render4 = "FAIL error " & Err.Description
    CloseDeck p
End Function

Public Function T_GanttTree() As String
    On Error GoTo fail
    gQuiet = True
    gLastMessage = ""
    Dim p As Object, r As String, g As Object, wk As Boolean, tb As Object, i As Long, boxes As Long, lines As Long
    Set p = TitledDeck(2)
    ' span parsing
    If ParseSpan("6 months", wk) <> 6 Or wk Then r = "FAIL parse months": GoTo done
    If ParseSpan("12 weeks", wk) <> 12 Or Not wk Then r = "FAIL parse weeks": GoTo done
    If ParseSpan("12w", wk) <> 12 Or ParseSpan("soon", wk) <> 0 Or ParseSpan("40 months", wk) <> 0 Then r = "FAIL parse edge": GoTo done
    ' gantt: 4 rows, 6 months -> header + 6 labels + 5 gridlines + 4 labels + 4 bars + milestone label + 2 diamonds = 23
    ActiveWindow.View.GotoSlide 1
    Set g = AddGantt(4, "6 months")
    If g Is Nothing Then r = "FAIL gantt " & gLastMessage: GoTo done
    If g.GroupItems.Count <> 23 Then r = "FAIL gantt parts " & g.GroupItems.Count: GoTo done
    If Not InsideSlide(g, p) Then r = "FAIL gantt off slide": GoTo done
    If Not (AddGantt(4, "forever") Is Nothing) Then r = "FAIL gantt bad span accepted": GoTo done
    ' issue tree: starter when nothing selected, then build from it
    ActiveWindow.View.GotoSlide 2
    ActiveWindow.Selection.Unselect
    ST_IssueTree
    If InStr(gLastMessage, "Edit this outline") = 0 Then r = "FAIL no starter msg": GoTo done
    Set tb = p.Slides(2).Shapes(p.Slides(2).Shapes.Count)
    Set g = BuildIssueTree(tb)
    If g Is Nothing Then r = "FAIL tree " & gLastMessage: GoTo done
    For i = 1 To g.GroupItems.Count
        If g.GroupItems(i).Type = 1 Then boxes = boxes + 1 Else lines = lines + 1
    Next i
    If boxes <> 7 Or lines <> 6 Then r = "FAIL tree parts boxes=" & boxes & " lines=" & lines: GoTo done
    If Not InsideSlide(g, p) Then r = "FAIL tree off slide": GoTo done
    If p.Slides(2).Shapes.Count <> 2 Then r = "FAIL outline not replaced (" & p.Slides(2).Shapes.Count & ")": GoTo done
    ' tab-indented plain text works too
    Set tb = p.Slides(2).Shapes.AddTextbox(ST_TEXT_HORIZONTAL, 50, 50, 300, 100)
    tb.TextFrame.TextRange.Text = "Root" & vbCr & vbTab & "A" & vbCr & vbTab & vbTab & "A1" & vbCr & vbTab & "B"
    Set g = BuildIssueTree(tb)
    If g.GroupItems.Count <> 7 Then r = "FAIL tab tree " & g.GroupItems.Count: GoTo done
    r = "PASS"
done:
    T_GanttTree = r
    CloseDeck p
    Exit Function
fail:
    T_GanttTree = "FAIL error " & Err.Description & " msg=" & gLastMessage
    CloseDeck p
End Function

Public Function T_Render5() As String
    On Error GoTo fail
    gQuiet = True
    Dim p As Object, d As String
    Set p = TitledDeck(2)
    p.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Transformation roadmap"
    p.Slides(2).Shapes.Title.TextFrame.TextRange.Text = "Where could the extra profit come from?"
    ActiveWindow.View.GotoSlide 1
    AddGantt 4, "6 months"
    ActiveWindow.View.GotoSlide 2
    BuildIssueTree AddOutlineStarter()
    d = Environ("TMPDIR")
    p.Slides(1).Export d & "st_r5_1.png", "PNG", 1280, 720
    p.Slides(2).Export d & "st_r5_2.png", "PNG", 1280, 720
    T_Render5 = "PASS"
    CloseDeck p
    Exit Function
fail:
    T_Render5 = "FAIL error " & Err.Description
    CloseDeck p
End Function
