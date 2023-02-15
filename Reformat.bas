Attribute VB_Name = "Reformat"
'**************************************************************************
' Runs all reformat macros
'**************************************************************************
Sub reformatEverything()
    reformatBulletedLists
    reformatTables
    'reformatNumberedLists 'DEPRECATED
    reformatImages
    
End Sub
'**************************************************************************
' Selects each table individually removes indent and bolds first row
'**************************************************************************
Sub reformatTables()

    Application.ScreenUpdating = False
    
    Dim oTable As Table
    For Each oTable In ActiveDocument.Tables
        
        oTable.Select
        
        ' removes indent
        'Selection.Paragraphs.Alignment = wdAlignParagraphLeft
        Selection.Paragraphs.LeftIndent = 0
        
        ' bolds header
        Selection.Rows.Item(1).Select
        Selection.BoldRun
        
        ' indents table title
        'If Selection.Previous(Unit:=wdLine, Count:=1) = WdListType.wdListNoNumbering Then
           ' Selection.Previous(Unit:=wdParagraph, Count:=1).Select 'wdParagraph and wdLine work
            'Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(4.01)
        'End If
    Next
    
    Application.ScreenUpdating = True
End Sub
'**************************************************************************
' Iterates through each bulleted/numbered list and fixes its formatting. Currently used to reformat both bulleted and numbered lists
'**************************************************************************
Sub reformatBulletedLists()

    Application.ScreenUpdating = False
    
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
    i.NumberPosition = CentimetersToPoints(4.8) ' indent from left margin
    i.TextPosition = CentimetersToPoints(5.6) ' position from left margin of text
    i.TabPosition = CentimetersToPoints(5.6) ' position of tab stop
    
    End If
Next i
Next p
    
    Application.ScreenUpdating = True

End Sub
'**************************************************************************
' DEPRECATED since reformatBulletedLists does this and is more reliable. Iterates through each numbered list and outdents it.
'**************************************************************************
Sub reformatNumberedLists()

    Dim oPara As Word.Paragraph
    For Each oPara In ActiveDocument.Paragraphs
       If oPara.Range.ListFormat.ListType = WdListType.wdListSimpleNumbering Then
        oPara.Outdent
        'oPara.Range.ListFormat.ApplyListTemplateWithLevel
       End If
    Next

End Sub
'**************************************************************************
' DOESN'T WORK
'**************************************************************************
Sub reformatFigTitles()
    Selection.HomeKey Unit:=wdStory
    For Each oShape In ActiveDocument.Shapes
        If oShape.Width <> 507.4016 Or oShape.Width = 391.46457 Then ' if the shape's width isn't full width
        'Alternate condition: Left <> CentimetersToPoints(2.4) Then
            oShape.Select
            Selection.Previous(Unit:=wdLine, Count:=1).Select
            Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(4.01)
        End If
    Next
End Sub
'**************************************************************************
' DOESN'T WORK
'**************************************************************************
Sub restyleBulletedLists()
'Iterates through each bulleted list and styles it
' comment more + give source

    Application.ScreenUpdating = False
    
    Dim oPara As Word.Paragraph
    For Each oPara In ActiveDocument.Paragraphs
       If oPara.Range.ListFormat.ListType = _
             WdListType.wdListBullet Then
             'oPara.Style = "BulletStyle" 'Applies custom style. caveat: won't work on level 2 bullets. another caveat: need to save the style
             Application.Run "TemplateProject.Styles.kbdListBullet" ' problem: runs only once
       End If
    Next
    
    Application.ScreenUpdating = True

End Sub
'**************************************************************************
' DOESN'T WORK
'**************************************************************************
Sub restyleEverything()
    'Application.Run "TemplateProject.Styles"
    Call kdbListBullet
End Sub
'**************************************************************************
' Goes over every image in the document. Ignores small icons. Ensures body graphics (and figure titles) are indented. Ensures full page graphics have full width.
'**************************************************************************
Sub reformatImages()
    Application.ScreenUpdating = False
    
    Const BODY_INDENT As Double = 113.6693 ' 4.01 cm
    Const MAX_FULL_WIDTH As Double = 507.4016 ' 17.90 cm 'unused
    Const MAX_BODY_WIDTH As Double = MAX_FULL_WIDTH - BODY_INDENT '476.2205 = 16.80 cm
    Const MIN_WIDTH As Double = 80 'arbitrary value used to exclude extremely small images (unlikely to be body graphics)
    
    Selection.HomeKey Unit:=wdStory
    
    For Each oShape In ActiveDocument.InlineShapes
    
    If oShape.Width < MIN_WIDTH Then ' skip current shape if it is too small
        GoTo ExitLine
    End If
    
    Set convertedShape = oShape.ConvertToShape ' convert shape to a floating shape
    
    'set the formatting of the floating shape to Top and Bottom
    With convertedShape.WrapFormat
    .Type = wdWrapTopBottom
    .AllowOverlap = False
    End With
    
    With convertedShape
    .LockAnchor = True ' seems to have no effect
    .LockAspectRatio = True
    End With
    
    'If the graphic is a body grpahic, indent it and the figure title above it
    If convertedShape.Width < MAX_BODY_WIDTH Then
        convertedShape.Left = BODY_INDENT
        convertedShape.Select
        Selection.Previous(Unit:=wdParagraph, Count:=1).Select 'wdParagraph and wdLine work
        Selection.ParagraphFormat.LeftIndent = BODY_INDENT
        
    'If the graphic is a full body graphic, scale it down, indent it and the figure title above it
    ElseIf convertedShape.Width > MAX_BODY_WIDTH Then
        convertedShape.Width = MAX_BODY_WIDTH
        convertedShape.Left = BODY_INDENT
        convertedShape.Select
        Selection.Previous(Unit:=wdParagraph, Count:=1).Select 'wdParagraph and wdLine work
        Selection.ParagraphFormat.LeftIndent = BODY_INDENT
    
    End If
ExitLine:
    Next oShape
    
    Application.ScreenUpdating = True

End Sub
'**************************************************************************
' DOESN'T WORK. WORKS FOR OASD COPY only.
'**************************************************************************
Sub OLDreformatImages()
    Application.ScreenUpdating = False
    
    Const MAX_BODY_WIDTH As Double = 391.46457
    Const MAX_FULL_WIDTH As Double = 507.4016
    
   
    
    Selection.HomeKey Unit:=wdStory
    
    For Each oShape In ActiveDocument.InlineShapes
    
    Set convertedShape = oShape.ConvertToShape ' convert shape to a floating shape
    
    'set the formatting of the floating shape to Top and Bottom
    With convertedShape.WrapFormat
    .Type = wdWrapTopBottom
    .AllowOverlap = True ' False would result in images bunching up together
    End With
    
    With convertedShape
    .LockAnchor = True ' seems to have no effect
    .LockAspectRatio = True
    End With
    
    'If a body graphic/screenshot is too big, then convert it into a full page graphic/screenshot
    If convertedShape.Width > MAX_BODY_WIDTH Then
        convertedShape.Width = MAX_FULL_WIDTH
        convertedShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        convertedShape.Left = CentimetersToPoints(2.4) 'ActiveDocument.PageSetup.PageWidth - Application.CentimetersToPoints(4.01) - .Width
        'Alternate solutions: using Shape.IncrementLeft
    Else
        convertedShape.Select
        Selection.Previous(Unit:=wdParagraph, Count:=1).Select 'wdParagraph and wdLine work
        Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(4.01)
    End If
    ' can replace reformatFigTitles by putting the three lines of code and replacing the oShape with convertedShape
    ' make this code into another function. because that is best practice. a function should do only ONE thing.
    
    
        'TODO: add figure reformatting call to that function here (since i should only indent it when the figure itself is indented)
        'set figtitlereformatting to private if using it as a helper method.
    Next oShape
    Application.ScreenUpdating = True

End Sub
'**************************************************************************
' Removes extra wording from figure titles (like Figure 1: Figure1Apple to Figure 1: Apple) Currently useless as there is no redundant wording.
'**************************************************************************
Sub removeFigWording()

Set myRange = ActiveDocument.Range(Start:=0, End:=0)
For i = 200 To 1 Step -1
    With myRange.Find
     .ClearFormatting
     .Text = ": Figure " & i
     With .Replacement
     .ClearFormatting
     .Text = ": "
     End With
     .Execute _
      Replace:=wdReplaceAll, _
      Format:=True, _
      MatchCase:=True, _
      MatchWholeWord:=True
    End With
Next i

End Sub
'**************************************************************************
' Used to test code which could possible be implemented in other Sub procedures
'**************************************************************************
Sub experimenter()

'**************************************************************************
' Indents all figure and table titles. May not be neccessary as this is already done in other procedures and we may misdiagnose
'**************************************************************************

Selection.Find.ClearFormatting
Selection.Find.Style = ActiveDocument.Styles("Caption")
With Selection.Find
    .Text = ""
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute
WordBasic.SelectSimilarFormatting
Selection.Paragraphs.LeftIndent = CentimetersToPoints(4.01)

End Sub
