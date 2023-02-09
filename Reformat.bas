Attribute VB_Name = "Reformat"
Sub reformatEverything()
    reformatTables
    reformatBulletedLists
    reformatNumberedLists
    removeFigWording
    reformatImages
    
End Sub

Sub reformatTables()
'Selects each table individually in the document and then removes indent
    Application.ScreenUpdating = False
    
    Dim oTable As Table
    For Each oTable In ActiveDocument.Tables
        oTable.Select
        Selection.Paragraphs.LeftIndent = 0
    Next
    
    Application.ScreenUpdating = True
End Sub

Sub reformatBulletedLists()
' should only be called ONCE
'Iterates through each bulleted list and outdents it
' comment more + give source

    'Application.ScreenUpdating = False
    
    Dim oPara As Word.Paragraph
    For Each oPara In ActiveDocument.Paragraphs
       If oPara.Range.ListFormat.ListType = WdListType.wdListBullet Or oPara.Range.ListFormat.ListType = WdListType.wdListPictureBullet Then
        oPara.Outdent
        'oPara.Range.ListFormat.ListOutdent
        ' outdents yet turns everything to level 1
       End If
    Next
    
    'Application.ScreenUpdating = True

End Sub

Sub reformatNumberedLists()
'Iterates through each numbered list and outdents it
' comment more + give source
' should only be called once

    Application.ScreenUpdating = False
    
    Dim oPara As Word.Paragraph
    For Each oPara In ActiveDocument.Paragraphs
       If oPara.Range.ListFormat.ListType = _
             WdListType.wdListSimpleNumbering Then 'TODO: investigate difference between wdListSimpleNumbering, wdListMixedNumbering, etc.
             oPara.Outdent
       End If
    Next
    
    Application.ScreenUpdating = True

End Sub
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

Sub restyleBulletedLists()
'Iterates through each bulleted list and styles it
' comment more + give source

    Application.ScreenUpdating = False
    
    Dim oPara As Word.Paragraph
    For Each oPara In ActiveDocument.Paragraphs
       If oPara.Range.ListFormat.ListType = _
             WdListType.wdListBullet Then
             'oPara.Style = "BulletStyle" 'Applies custom style. caveat: won't work on level 2 bullets. another caveat: need to save the style
             ' this is a very important concept that shoudld be applied to figure reformatting.
             Application.Run "TemplateProject.Styles.kbdListBullet" ' problem: runs only once
             ' interesting thing: can access macros that i can't edit
       End If
    Next
    
    Application.ScreenUpdating = True

End Sub

Sub restyleEverything()
    'Application.Run "TemplateProject.Styles"
    Call kdbListBullet
End Sub

Sub reformatImages()
    Application.ScreenUpdating = False
    
    Const MAX_BODY_WIDTH As Double = 391.46457
    Const MAX_FULL_WIDTH As Double = 507.4016
    
    Selection.HomeKey Unit:=wdStory
    
    For Each oShape In ActiveDocument.InlineShapes
    
    Set convertedShape = oShape.ConvertToShape ' convert shape to a floating shape
    
    'set the formatting of the floating shape to Top and Bottom
    With convertedShape.WrapFormat
    .Type = wdWrapTopBottom
    .AllowOverlap = False
    .DistanceTop = 0 ' don't this and DistanceBottom are necessary
    .DistanceBottom = 0
    End With
    
    With convertedShape
    .LockAnchor = True
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

Sub experimenter()

Application.ScreenUpdating = False
    Selection.HomeKey Unit:=wdStory
    Dim oStr As String
    Dim oFig As InlineShape
    
    ' Make new style
    AddNewStyle "FigT"
     
    ' Loop through all available inline shapes
    For Each oFig In ActiveDocument.InlineShapes
        On Error GoTo FigErr:
        oFig.Select
        On Error GoTo FigErr:
        Selection.Previous(Unit:=wdLine, Count:=1).Select
        On Error GoTo FigErr:
        oStr = Selection.Range.Text
        If oStr Like "*Figure*" Then
            Selection.Range.FormattedText
            Selection.Range.Font.Position
            
            Selection.Range.Style = "FigT"
        End If
FigContinue:
    Next oFig

Application.ScreenUpdating = False
End Sub

