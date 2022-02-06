Attribute VB_Name = "Module1"
Option Explicit

Sub getPosition()

Dim shp As Shape

Set shp = ActivePresentation.Slides(1).Shapes(1)

Debug.Print (shp.Top)
Debug.Print (shp.Left)
Debug.Print (shp.Width)
Debug.Print (shp.Height)

End Sub
