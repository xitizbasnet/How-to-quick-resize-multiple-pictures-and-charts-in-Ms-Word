# How-to-quick-resize-multiple-pictures-and-charts-in-Ms-Word
How to quick resize multiple pictures and charts in Ms Word


Are you tired of your office document work containing numerous pictures and charts?

Think easy and innovative.
There is a way you can change and resize your pictures and charts in just a click.
Have you heard about VBA coding? If yes just go straight for code below, if not lets learn just in a minute.
First open your document in which you have pictures and charts that needs to be resized at once.

.......................................................................

Then, if you don't have developer option tab in menu section, follow this procedure.
-Right click on menu tab.

![2](https://user-images.githubusercontent.com/22270483/151377077-af026e15-f0e7-4345-9ea0-cb9f38ba7d4e.png)

.......................................................................

-click on customize quick access toolbar.

-select Customize ribbon tab and in choose command from drop down list to click for Main tabs.

![3](https://user-images.githubusercontent.com/22270483/151377115-720c0d3e-8233-480f-b8b9-2a1a995c89b3.png)

.......................................................................

-Then click Developer tab and click ADD button and press OK.

![4](https://user-images.githubusercontent.com/22270483/151377259-4200ca9c-ea95-41ea-882a-17f23f75e0ba.png)

.......................................................................

-Now goto Developer menu tab and Click Visual Basic editor.

![5](https://user-images.githubusercontent.com/22270483/151377337-3d05c778-0e3f-4c87-b18c-b1ae023568b7.png)

.......................................................................

-Now Add a module and Type the code on the module window.

![6](https://user-images.githubusercontent.com/22270483/151377409-75cd7a2d-ffef-42d8-9f87-d4b0be0774ba.png)

![7](https://user-images.githubusercontent.com/22270483/151377461-65f92779-4ea8-4c81-8267-7af36607773f.png)

.......................................................................

CODE FOR RESIZING THE PICTURES:

Sub resize()
Dim x As Long
With ActiveDocument
For x = 1 To .InlineShapes.Count
With .InlineShapes(x)
.Height = InchesToPoints(2)
.Width = InchesToPoints(3)
End With
Next x
End With
End Sub

Hope you find a easy method for your resizing mass pictures. This can be done in every office packages like in excel,powerpoint as well. 

Thank you.

.......................................................................
