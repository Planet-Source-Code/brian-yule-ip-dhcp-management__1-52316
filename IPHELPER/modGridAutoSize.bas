Attribute VB_Name = "modGridAutoSize"
Option Explicit

Private Const FLEXGRIDBORDERSIZE = 7

Public Function AutoSizeFlexGrid(ByRef fGrid As MSFlexGridLib.MSFlexGrid)
' This auto sizes the grid columns to the size of the data in the grid
    
    Dim prevRedrawState As Boolean
    
    With fGrid
        prevRedrawState = .Redraw ' Save the current redraw state
        If prevRedrawState Then .Redraw = False ' Set the grid not to redraw
        If .Cols > 0 And .Rows > 0 Then ' If there are more than 0 columns and rows
            Dim i As Integer, j As Integer
            For j = 0 To fGrid.Cols - 1 ' For each column
                Dim max As Long
                fGrid.Col = j
                For i = 0 To fGrid.Rows - 1 ' For each row
                    fGrid.Row = i
                    Dim txtlength As Integer
                    txtlength = fGrid.Parent.TextWidth(fGrid.Text) ' Get the text width of the data
                    If txtlength > max Then max = txtlength ' Select the max width
                Next i
                fGrid.ColWidth(j) = max + (FLEXGRIDBORDERSIZE * Screen.TwipsPerPixelY) ' Se the width of the column to the max width
                max = 0 ' reset the max width
            Next j
        End If
        If prevRedrawState Then .Redraw = True ' Set the redraw to its initial state
    End With
End Function
