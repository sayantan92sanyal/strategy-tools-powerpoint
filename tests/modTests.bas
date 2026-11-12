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
