Sub BodyMacro()
' Khalid Alzuhair
' This is a Macro that edits the body part of a report
' BodyMacro Macro
' This macro sets the margins to be 0.75", 1", 0.63", and 0.63", the references to be Chicago,
' the spacing to be 0, 0, 0, 6, and formats the text to be two columns each with 3.5" of width
'
    With Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(0.75)
        .BottomMargin = InchesToPoints(1)
        .LeftMargin = InchesToPoints(0.63)
        .RightMargin = InchesToPoints(0.63)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.5)
        .FooterDistance = InchesToPoints(0.5)
        .PageWidth = InchesToPoints(8.5)
        .PageHeight = InchesToPoints(11)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .SectionDirection = wdSectionDirectionLtr
    End With
    
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    
    If ActiveWindow.ActivePane.View.Type <> wdPrintView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    
    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=1
        .EvenlySpaced = False
        .LineBetween = False
        .FlowDirection = wdFlowLtr
    End With
    
    Selection.PageSetup.TextColumns.Add Width:=InchesToPoints(3.5), Spacing:= _
        InchesToPoints(0.25), EvenlySpaced:=False
        
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    
    Selection.Font.Name = "Times New Roman"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    
End Sub
