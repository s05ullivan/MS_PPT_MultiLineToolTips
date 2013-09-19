MS_PPT_MultiLineToolTips
========================

How to create multi-line screen tips in PowerPoint (VBA)

The VBA macro below enables Microsoft PowerPoint presentations to provide the interactive feature of displaying Screen Tips (Popup / Hover Text) when the mouse is moved over shapes during presentation mode.  The macro is only needed to apply edits – it is not needed to make the presentation work.

To set the presentation mode hover text for a PowerPoint shape:

Enter the desired text in the Alternative Text area with the “Size and Position…” option from the pop-up menu for the shape.
Use the macro below to copy all the Alternative Text entries into the shape’s Screen Tips.
 
--~--

Public Sub CopyAllShapesAlternateTxtToScreenTip

‘Sean O’Sullivan — 2012

 

Dim mySlide As Slide

Dim myShapes As Shapes

 

‘Count is useful for debugging

Dim count As Integer

count = 0

 
For Each mySlide In ActivePresentation.Slides

    Set myShapes = mySlides.Shapes

 
    Dim xShape as Shape


    For Each xShape In myShapes

        count = count + 1

        If xShape.AlternativeText <> “” Then

            If xShape.ActionSettings(ppMouseClick).Action <> ppActionHyperlink Then

                ‘By default set the shape hyperlink back to the current slide – otherwise don’t overwrite the hyperlink if one already exists.

                xShape.ActionSettings(ppMouseClick).Hyperlink.SubAddress = mySlide.SlideNumber

            End If


        xShape.ActionSettings(ppMouseClick).Hyperlink.ScreenTip = xShape.AlternativeText

        End If


     Next ‘Shape

    Next ‘Slide

End Sub ‘CopyAllShapesAlternateTxtToScreenTip

