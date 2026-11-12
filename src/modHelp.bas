Attribute VB_Name = "modHelp"
' Strategy Tools - keyboard-shortcut actions and the shortcuts cheat sheet.
Option Explicit

' Ctrl+Q: drop a ready-made text box in the middle of the current slide,
' selected and ready for typing.
Public Sub ST_AddTextBox()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Const W As Single = 216     ' 3 inches
    Const H As Single = 30
    Dim sld As Object, tb As Object
    Set sld = CurrentSlide()
    Set tb = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, _
        (ActivePresentation.PageSetup.SlideWidth - W) / 2, _
        (ActivePresentation.PageSetup.SlideHeight - H) / 2, W, H)
    tb.TextFrame.WordWrap = ST_TRUE
    tb.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    tb.TextFrame.TextRange.Font.Size = 14
    tb.Name = "Text Box"
    tb.TextFrame.TextRange.Select
    Exit Sub
fail:
    ReportError "New text box"
End Sub

' Cmd+Shift+O and the Help button: open the shortcuts cheat sheet.
' The page lives in ~/Library/Application Support/StrategyTools and is opened
' by an AppleScript helper that setup-shortcuts.sh installs (Office's sandbox
' only lets VBA run scripts from that helper folder).
Public Sub ST_ShowShortcuts()
    On Error GoTo fail
    AppleScriptTask "StrategyTools.scpt", "openCheatSheet", ""
    Exit Sub
fail:
    Msg "The shortcuts cheat sheet isn't installed yet." & vbCr & vbCr & _
        "Run 'Install Shortcuts' from the Strategy Tools folder (see INSTALL guide), then try again."
End Sub
