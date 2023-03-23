Attribute VB_Name = "Reformat"
'**************************************************************************
' Runs all reformat macros
'**************************************************************************
Sub reformatEverything()

    reformatTables
    reformatNumberedLists
    restyleBulletLists
    reformatImages
    restyleSectionTitles
    addSpaceAfterHeadings
    addSpaceAfterImages
    
End Sub
'**************************************************************************
' Selects each table individually removes indent and bolds first row
'**************************************************************************
Sub reformatTables()

    Application.ScreenUpdating = False
    
    Dim unboldedTables As Integer
    unboldedTables = 0
    
    Dim oTable As Table
    For Each oTable In ActiveDocument.Tables
        oTable.Select
        
        'indents table if there is a table title
        If Selection.Previous(Unit:=wdParagraph, Count:=1).Style = "Caption" Then
            Selection.Previous(Unit:=wdParagraph, Count:=1).Select
            Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(4.01)
        End If
        
        oTable.Select
        Selection.Paragraphs.LeftIndent = 0
        
        On Error GoTo ExceptionHandler:
            Selection.Rows.Item(1).Select
            Selection.BoldRun
            
    GoTo NextTable
    
ExceptionHandler:
    unboldedTables = unboldedTables + 1
    Resume NextTable
    
'Looks useless, but Next oTable cannot be put in the ExceptionHandler label
NextTable:
    Next oTable

    If unboldedTables > 0 Then
        MsgBox (unboldedTables & " table header(s) have not been bolded.")
    End If

    Application.ScreenUpdating = True
End Sub

'**************************************************************************
' Reapplies the List Bullet styles
'**************************************************************************
Sub restyleBulletLists()

    Application.ScreenUpdating = False
    
    Dim search_style As String ' the style which apparently seem out of style
    Dim replace_style As String ' the desired style
    
    With Selection.Find
        .Text = ""
        .ClearFormatting
        .Style = "List Bullet 2"
        .Replacement.Text = ""
        .Replacement.ClearFormatting
        .Replacement.Style = "List Bullet 2"
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    
    With Selection.Find
        .Text = ""
        .ClearFormatting
        .Style = "List Bullet"
        .Replacement.Text = ""
        .Replacement.ClearFormatting
        .Replacement.Style = "List Bullet"
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True

End Sub

Sub reformatNumberedLists()

    Application.ScreenUpdating = False

    ' using this inefficient code because re applying styles doesn't properly reset numbering
    Dim LP As ListParagraphs
    Dim p As Paragraph
    Dim i As ListLevel
    Set LP = ActiveDocument.ListParagraphs
    For Each p In LP
        For Each i In p.Range.ListFormat.ListTemplate.ListLevels
            If i.Index = 1 And Not i.Index = 2 Then
                i.TrailingCharacter = wdTrailingTab
                i.NumberPosition = CentimetersToPoints(4) ' indent from left margin
                i.TextPosition = CentimetersToPoints(4.8) ' position from left margin of text
                i.TabPosition = CentimetersToPoints(4.8) ' position of tab stop
            ElseIf i.Index = 2 And Not i.Index = 1 Then
                i.TrailingCharacter = wdTrailingTab
                i.NumberPosition = CentimetersToPoints(4.8)
                i.TextPosition = CentimetersToPoints(5.6)
                i.TabPosition = CentimetersToPoints(5.6)
            End If
        Next i
    Next p
    
    Application.ScreenUpdating = True

End Sub
'**************************************************************************
' Goes over every image in the document. Ignores small icons. Ensures body graphics (and figure titles) are indented. Ensures full page graphics have full width.
'**************************************************************************
Sub reformatImages()
    Application.ScreenUpdating = False
    
    Const BODY_INDENT As Double = 113.6693 ' 4.01 cm
    Const MAX_FULL_WIDTH As Double = 481.0193 ' 16.969 cm
    Const MAX_BODY_WIDTH As Double = 367.35 'value derived from table width = 12.959 cm 'MAX_FULL_WIDTH - BODY_INDENT
    Const MIN_WIDTH As Double = 100 'arbitrary value used to exclude extremely small images (unlikely to be body graphics)
    
    'Selection.HomeKey Unit:=wdStory ' used only for testing purposes
    
    For Each oShape In ActiveDocument.InlineShapes
    
        oShape.Select
        If oShape.Width < MIN_WIDTH Or Selection.Previous(Unit:=wdParagraph, Count:=1).Style <> "Caption" Then ' skip current shape if it is too small or doesn't have a caption above it
            GoTo NextShape
        End If
        
        Set convertedShape = oShape.ConvertToShape ' convert shape to a floating shape
        
        'set the formatting of the floating shape to Top and Bottom
        With convertedShape.WrapFormat
        .Type = wdWrapTopBottom
        .AllowOverlap = False
        End With
        
        With convertedShape
        .LockAnchor = True
        .LockAspectRatio = True
        End With
            
        'If the graphic is a full body graphic, scale it down, indent it and the figure title above it
        If convertedShape.Width > MAX_BODY_WIDTH Then
            convertedShape.Width = MAX_BODY_WIDTH
        End If
        convertedShape.Left = BODY_INDENT
        
        'indents figure title
        convertedShape.Select
        Selection.Previous(Unit:=wdParagraph, Count:=1).Select 'wdParagraph and wdLine work
        Selection.ParagraphFormat.LeftIndent = BODY_INDENT
        'Selection.ParagraphFormat.SpaceAfter = 6 'doesn't work
        Selection.EndOf
        
'        For Each convertedShape In ActiveDocument.Shapes
'            convertedShape.Select
'            Selection.ShapeRange.ConvertToInlineShape
'        Next convertedShape
        
        'TODO spaces text below image
'        convertedShape.Select
'        Selection.Next(Unit:=wdLine, Count:=1).Select 'wdParagraph and wdLine work
'        Selection.ParagraphFormat.SpaceBeforeAuto = 6
'        Selection.EndOf
        GoTo NextShape
        
NextShape:
    Next oShape
    
    Application.ScreenUpdating = True

End Sub
'**************************************************************************
' Styles the section titles (from DITA) to body text
'**************************************************************************
Sub restyleSectionTitles()

Application.ScreenUpdating = False

    Dim oPara As Paragraph
    For Each oPara In ActiveDocument.Paragraphs
        If oPara.Style = "Subtitle" Then
            oPara.Style = "Body Text,Corpo del testo Carattere,Corpo del testo Carattere Carattere Carattere,Corpo del testo Carattere Carattere,Body Text Char1 Char,Body Text Char Char Char,Body Text Char2 Char Char Char,Body Text Char1 Char Char Char Char"
            oPara.SpaceAfter = 6
            oPara.Range.Bold = True
        End If
    Next oPara

Application.ScreenUpdating = True

End Sub

'**************************************************************************
' Sets space after all headings to 12 points
'**************************************************************************
Sub addSpaceAfterHeadings()
    Application.ScreenUpdating = False
    
    ' loop over all headings. if they don't have 12pts after, make it so
    Dim oHeading As Paragraph
    For Each oHeading In ActiveDocument.Paragraphs
        If (oHeading.Style = "Heading 1" Or oHeading.Style = "Heading 2" Or oHeading.Style = "Heading 3" Or oHeading.Style = "Heading 4" _
        Or oHeading.Style = "Heading 5" Or oHeading.Style = "Heading 6") And oHeading.SpaceAfter <> 12 Then
            oHeading.SpaceAfter = 12
        End If
    Next oHeading

    Application.ScreenUpdating = True
End Sub

'**************************************************************************
' Sets space after all images to 6 points
'**************************************************************************
'Note: uses normal style to identify images because using inline shapes doesn't work properly
Sub addSpaceAfterImages()
    Application.ScreenUpdating = False
    
    Dim oImage As Paragraph
    For Each oImage In ActiveDocument.Paragraphs
        If oImage.Style = "Normal" And oImage.SpaceAfter <> 6 Then
            oImage.SpaceAfter = 6
        End If
    Next oImage

    Application.ScreenUpdating = True
End Sub

'**************************************************************************
' Used to change and test code from other procedures
'**************************************************************************
Sub experimenter()
    Application.ScreenUpdating = False

    Application.ScreenUpdating = True
End Sub
