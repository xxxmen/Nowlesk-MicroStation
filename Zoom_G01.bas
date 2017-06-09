Attribute VB_Name = "Zoom_G01"

Sub getRange()
 On Error GoTo errhnd
 Dim lngDspPrty As Long
  
 Dim ele As CellElement
 Dim success As Boolean
 success = False
 Dim rng As Range3d
 Dim BorderName As String
 
 Set ele = F2_G01.GetBorder(False)
 rng = ele.Range
 BorderName = ele.Name
 
 success = ZoomToTitle(rng, BorderName, 1)
 
 MsgBox success
errhnd:
   Select Case Err.Number
     Case 91 'Get Border didn't find any Borders
             'Could be a raster file a raster Title Block
       MsgBox "Program ended! No title block on this drawing."
       Err.Clear
     End Select
End Sub

Function ZoomToTitle(Rngr As Range3d, BDR_X1X As String, viewNmbr As Integer) As Boolean
 Dim dblFactor As Double
 Dim DeltaY As Double
 Dim DeltaX As Double
 Dim oView As View
 Set oView = ActiveDesignFile.Views(viewNmbr)
 Dim pntOrigin As Point3d
 Dim rngExtents As Range3d
 Dim pntExtents As Point3d
 Dim myLine As LineElement
 Dim pntZoom As Point3d
 
 'Establish extents just around the Title Block area
 'This allows for extra elements outside of the Title Block
 'area to not affect the zoom into the title area of the Title
 'Block
 '*******************************************
 rngExtents = Rngr
 
 oView.Origin = rngExtents.Low
 
 pntExtents.x = rngExtents.High.x - rngExtents.Low.x
 pntExtents.Y = rngExtents.High.Y - rngExtents.Low.Y
 
 oView.Extents = pntExtents
 'oView.Redraw
 'oView.Redraw
 '********************************************
 
 pntOrigin.x = Rngr.Low.x
 pntOrigin.Y = Rngr.Low.Y
 pntOrigin.Z = 0
 
 DeltaX = Rngr.High.x - Rngr.Low.x
 DeltaY = Rngr.High.Y - Rngr.Low.Y
 
 Select Case BDR_X1X
    Case "BDR-D10"
        With Rngr
          pntZoom.x = .Low.x + (1.488623 * DeltaY)
          pntZoom.Y = .Low.Y + (0.2227318 * DeltaY)
        End With
        dblFactor = 0.43
     Case "BDR-E10"
        With Rngr
          pntZoom.x = .Low.x + (1.358326 * DeltaY)
          pntZoom.Y = .Low.Y + (0.163336 * DeltaY)
        End With
        dblFactor = 0.32
     Case "BDR-D12"
        With Rngr
          pntZoom.x = .Low.x + (1.488623 * DeltaY)
          pntZoom.Y = .Low.Y + (0.2227318 * DeltaY)
        End With
        dblFactor = 0.43
     Case "BDR-E12"
        With Rngr
          pntZoom.x = .Low.x + (1.358323 * DeltaY)
          pntZoom.Y = .Low.Y + (0.163336 * DeltaY)
        End With
        dblFactor = 0.32
     Case "BDR-T10"
        With Rngr
          pntZoom.x = .Low.x + (0.9702261904 * DeltaX)
          pntZoom.Y = .Low.Y + (0.16335 * DeltaY)
        End With
        dblFactor = 0.32
     Case "BDR-T12"
        With Rngr
          pntZoom.x = .Low.x + (0.9702214 * DeltaX)
          pntZoom.Y = .Low.Y + (0.16333 * DeltaY)
        End With
        dblFactor = 0.32
     Case Else1
       ZoomToTitle = False
       Exit Function
 End Select
 
         
      With Application
        Set myLine = .CreateLineElement2(Nothing, pntOrigin, pntZoom)
        .ActiveModelReference.AddElement myLine
      End With
 'Zoom about the center of the range.
 oView.ZoomAboutPoint pntZoom, dblFactor
 oView.Redraw
 oView.Redraw
 ZoomToTitle = True
 

End Function



