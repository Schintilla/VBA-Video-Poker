Attribute VB_Name = "Startup"
'Sub Auto_Open()
'    ActiveSheet.UnProtect
'    With Worksheets("Deck Data")
'        .Range("S2").Value = 0
'        .Range("S3").Value = 0
'        .Range("S4").Value = Application.Width
'        .Range("S5").Value = Application.Height
'        .Range("S6").Value = Application.WindowState
'    End With
'    With Worksheets("Display")
'        .DrawingObjects("window").Characters.Text = "Window Set"
'        .DrawingObjects("menu").Characters.Text = "Menu Reset"
'        .Select
'        .Range("B1").Select
'        .Range("AJ10") = False
'    End With
'    Application.ErrorCheckingOptions.BackgroundChecking = False
''    screenset
'    menuset
''    RibbonHeight = 100
''    If Application.CommandBars.Item("Ribbon").Height > RibbonHeight Then
''        Application.CommandBars.ExecuteMso "MinimizeRibbon"
''    End If
'    Worksheets("Display").Range(Cells(14 + 1, 3 + 31), Cells(14 + 4, 3 + 31)).Select
'End Sub

Sub Auto_Close()
    ActiveSheet.UnProtect
    Worksheets("Display").DrawingObjects("window").Characters.Text = "Window Reset"
    screenset
    Worksheets("Display").DrawingObjects.Ungroup
    On Error Resume Next
    'Worksheets("Display").DrawingObjects.Group
    MenuBars("Video Poker").Delete
    CommandBars.ExecuteMso "MinimizeRibbon"
    Application.ErrorCheckingOptions.BackgroundChecking = True
    On Error GoTo 0
End Sub
Sub menuset()
    With MenuBars.Add("Video Poker")
        With .Menus.Add("Game")
            .MenuItems.Add "Play", "screenset"
            .MenuItems.Add "New", "reset"
            .MenuItems.Add "Save", "savegame"
            .MenuItems.Add "-"
            .MenuItems.Add "Menu Toggle", "menureset"
            .MenuItems.Add "Edit Program", "screenset"
            .MenuItems.Add "-"
            .MenuItems.Add "Exit", "closepoker"
        End With
        With .Menus.Add("View")
            .MenuItems.Add "Entire Deck", "displaydeck"
            .MenuItems.Add "Next Five Cards", "nextfivecards"
            .MenuItems.Add "Hold Strategy", "gotohandedit"
            .MenuItems.Add "-"
            .MenuItems.Add "Payout Schedule", "gotoshed"
            .MenuItems.Add "Joker Wild Schedule", "jokertable"
            .MenuItems.Add "Probability Table", "gotoprob"
        End With
        With .Menus.Add("Statistics")
            .MenuItems.Add "Graphical", "gotograph"
            .MenuItems.Add "Hands Won", "gotohand"
        End With
        With .Menus.Add("Help")
            .MenuItems.Add "Contents", "contents"
            .MenuItems.Add "-"
            .MenuItems.Add "About XL Video Poker", "about"
        End With
    End With
    MenuBars("Video Poker").Activate
End Sub

Sub menureset()
    ActiveSheet.UnProtect
    With Worksheets("Display")
        If .DrawingObjects("menu").Characters.Text = "Menu Reset" Then
            MenuBars(xlWorksheet).Activate
            .DrawingObjects("menu").Characters.Text = "Menu Set"
        Else
            MenuBars("Video Poker").Activate
            .DrawingObjects("menu").Characters.Text = "Menu Reset"
        End If
    End With
End Sub
Sub screenset()
    ActiveSheet.UnProtect
    If ActiveSheet.DrawingObjects("window").Characters.Text = "Window Set" Then
        onstatus = False
        lft = 50
        Tp = 25
        Wdh = 533
        Ht = 360
        CommandBars.ExecuteMso "MinimizeRibbon"
        ActiveSheet.DrawingObjects("window").Characters.Text = "Window Reset"
    Else
        onstatus = True
        With Worksheets("Deck Data")
            lft = .Range("S2").Value
            Tp = .Range("S3").Value
            Wdh = .Range("S4").Value
            Ht = .Range("S5").Value
            Winstat = .Range("S6").Value
        End With
        CommandBars.ExecuteMso "MinimizeRibbon"
        ActiveSheet.DrawingObjects("window").Characters.Text = "Window Set"
    End If
    With ActiveWindow
        .DisplayGridlines = onstatus
        .DisplayHeadings = onstatus
        .DisplayHorizontalScrollBar = onstatus
        .DisplayVerticalScrollBar = onstatus
        .DisplayWorkbookTabs = onstatus
        .DisplayHeadings = onstatus
    End With
    With Application
        .DisplayFormulaBar = onstatus
        .DisplayStatusBar = onstatus
        .DisplayNoteIndicator = onstatus
    End With
    'If Winstat = 1 Then
    'Else
    'End If
    Application.WindowState = xlNormal
    '    ActiveWindow.WindowState = xlNormal
    '    With ActiveWindow
    '        .Width = 504
    '        .Height = 303
    '    End With
    'Application.WindowState = xlMaximized
    With Application
        .Left = lft
        .Top = Tp
        .Width = Wdh
        .Height = Ht
    End With
    Toolbars(1).Visible = onstatus
    Toolbars(2).Visible = onstatus
End Sub
Sub about()
    MsgBox "Excel Video Poker Simulator" + Chr$(13) + "(Dedicated to Cath, my wife)" + Chr$(13) + "Copyright 1998" + Chr$(13) + Chr$(13) + "E-mail: home@cyberpad.freeserve.co.uk"
End Sub
Sub closepoker()
    ActiveSheet.UnProtect
    Worksheets("Display").DrawingObjects("window").Characters.Text = "Window Reset"
    screenset
    Worksheets("Display").DrawingObjects("menu").Characters.Text = "Menu Reset"
    menureset
    'Worksheets("Display").DrawingObjects.Ungroup
    On Error Resume Next
    'Worksheets("Display").DrawingObjects.Group
    MenuBars("Video Poker").Delete
    On Error GoTo 0
    a$ = ActiveWorkbook.name
    Workbooks(a$).Close
End Sub

Sub savegame()
    ActiveWorkbook.Save
End Sub
