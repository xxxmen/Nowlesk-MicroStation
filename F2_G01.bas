Attribute VB_Name = "F2_G01"
Option Explicit

Sub testUserForm1()
  UserForm1.show vbModeless
End Sub

Sub test_SeeAttachment()
 Dim message As String
 message = SeeAttachment
 MsgBox message
End Sub

Function SeeAttachment() As String
  Dim RasterFullName As String
  Dim RasterPath As String
  Dim DesignFileName As String
  Dim DesignFilePath As String
  Dim Icount As Integer
  
  Dim RasterCount As Integer
        'gets the name of the first attached raster and the number of rasters attached
        'need full path since raster could be in another folder
   Dim att As Rasters
   Set att = Application.RasterManager.Rasters
        'are there any attachments
    If att.count = 0 Then
          'do nothing --> No raster images are attached
          SeeAttachment = ""
          Exit Function
       Else
          If RasterCount > 1 Then
            SeeAttachment = "Too Many Rasters"
            Exit Function
          End If
    End If
    
  MsgBox att.Item(1).RasterInformation.FullName & vbLf & "count: " & att.count
  MsgBox att.Item(1).RasterInformation.path
  
End Function


Sub SeeFileName()
  Dim fna As DesignFile
  Set fna = Application.ActiveDesignFile
  MsgBox fna.Name
  MsgBox fna.path
End Sub

Sub FileAttributes()
  Dim message As String
  Dim SnapE As Boolean
  Dim UnitL As Boolean
  Dim graphG As Boolean
  Dim activeR As Boolean
  Dim ActRefMod As ModelReference
  
  With Application.ActiveSettings
      .SnapLockEnabled = True
      .UnitLockEnabled = True
      .GraphicGroupLockEnabled = True
        '    .GridUnits
        '    .GridReference
       .AxisLockEnabled = False
       .GridLockEnabled = False
       
  End With
  
  CadInputQueue.SendKeyin "LOCK SNAP KEYpoint"
  
 If activeR = Application.HasActiveModelReference Then
   Set ActRefMod = Application.ActiveModelReference
 End If
      
   MsgBox SnapE & UnitL
    
End Sub

 
Sub F2_G01()  'G --> retive information and reset settings before Getting Out Of Drawing
    'by Keith Knowles 12/10/2013
    
    Dim message As String
    Dim bdrElement As CellElement
    Set bdrElement = GetBorder(False)

    'MsgBox bdrElement.Name & "made it"
    
    'put on the fence
    message = GetThePoints(bdrElement)
    message = TableColor_G01.ColorTable
     
    If ActiveDesignFile.Views(5).IsOpen Then
        CadInputQueue.SendCommand "FIT VIEW EXTENDED 5"
      Else
        message = message & "FIX VIEW 5!"
    End If

     If message <> "" Then
       ShowStatus message
     End If
     
End Sub
Function GetThePoints(BDR As CellElement) As String
     Dim delta_Y As Variant
     Dim delta_X As Variant
     Dim D12_Ratio As Double
     D12_Ratio = 1.54545454545455
     Const E12_Ratio As Double = 1.4
     Dim FortyTwo As Variant
     FortyTwo = 42#
     Dim ThirtyFour As Variant
     ThirtyFour = 34#
     Dim curElem As Element
     
     Dim pts(1 To 4) As Point3d
        'BDR.Origin.x
        
        'If non-scaled drawings
      If BDR.Name = "BDR-D10" Then
        pts(1).x = BDR.Range.Low.x
        pts(1).Y = BDR.Range.Low.Y
        pts(2).x = BDR.Range.Low.x + ThirtyFour
        pts(2).Y = BDR.Range.Low.Y
        pts(3).x = BDR.Range.Low.x + ThirtyFour
        pts(3).Y = BDR.Range.High.Y
        pts(4).x = BDR.Range.Low.x
        pts(4).Y = BDR.Range.High.Y
        
       ElseIf BDR.Name = "BDR-E10" Then
        pts(1).x = BDR.Range.Low.x
        pts(1).Y = BDR.Range.Low.Y
        pts(2).x = BDR.Range.Low.x + FortyTwo
        pts(2).Y = BDR.Range.Low.Y
        pts(3).x = BDR.Range.Low.x + FortyTwo
        pts(3).Y = BDR.Range.High.Y
        pts(4).x = BDR.Range.Low.x
        pts(4).Y = BDR.Range.High.Y
       
       
       ElseIf BDR.Name = "BDR-T10" Or BDR.Name = "BDR-T12" Then
        pts(1).x = BDR.Range.Low.x
        pts(1).Y = BDR.Range.Low.Y
        pts(2).x = BDR.Range.High.x
        pts(2).Y = BDR.Range.Low.Y
        pts(3).x = BDR.Range.High.x
        pts(3).Y = BDR.Range.High.Y
        pts(4).x = BDR.Range.Low.x
        pts(4).Y = BDR.Range.High.Y
       
       ElseIf BDR.Name = "BDR-D12" Then
        delta_Y = BDR.Range.High.Y - BDR.Range.Low.Y
        'deduce delta_X by  Ratio 17/11
        delta_X = D12_Ratio * delta_Y
        pts(1).x = BDR.Range.Low.x
        pts(1).Y = BDR.Range.Low.Y
        pts(2).x = BDR.Range.Low.x + delta_X
        pts(2).Y = BDR.Range.Low.Y
        pts(3).x = BDR.Range.Low.x + delta_X
        pts(3).Y = BDR.Range.High.Y
        pts(4).x = BDR.Range.Low.x
        pts(4).Y = BDR.Range.High.Y
        
       ElseIf BDR.Name = "BDR-E12" Then
        delta_Y = BDR.Range.High.Y - BDR.Range.Low.Y
        'deduce delta_X by ratio 14/10
        delta_X = E12_Ratio * delta_Y
        pts(1).x = BDR.Range.Low.x
        pts(1).Y = BDR.Range.Low.Y
        pts(2).x = BDR.Range.Low.x + delta_X
        pts(2).Y = BDR.Range.Low.Y
        pts(3).x = BDR.Range.Low.x + delta_X
        pts(3).Y = BDR.Range.High.Y
        pts(4).x = BDR.Range.Low.x
        pts(4).Y = BDR.Range.High.Y
       Else
         MsgBox "No border on drawing!"
       End If
              
    ' if fence happens
    ' GetThePoints = True
    With ActiveDesignFile.Fence
       .DefineFromModelPoints 1, pts()
       .Draw msdDrawingModeHilite
    End With
    
    
    If BDR.Range.Low.x <> 0 Or BDR.Range.Low.Y <> 0 Then
      GetThePoints = "BORDER OFF 0,0! "
    End If
    
End Function
Function GetBorder(ignoreT As Boolean) As Element

  Dim rngBDR As Range3d
  Dim pntBDRs As Point3d
  Dim pntBDRe As Point3d
  Dim rngTBDR As Range3d
  Dim pntTBDRs As Point3d
  Dim pntTBDRe As Point3d
  Dim dblScale As Double
  
  Dim BorderName As String
  Dim oElem As Element
  Dim oCellElem As CellElement
  Dim BdrObject As CellElement
  Dim TbdrObject As CellElement
  Dim oEnum As ElementEnumerator
  Dim ElementCounter As Long
  Dim BorderType As String
  
  Dim BorderD10 As Boolean
  Dim BorderE10 As Boolean
  Dim BorderT10 As Boolean
  
  Dim BorderD12 As Boolean
  Dim BorderE12 As Boolean
  Dim BorderT12 As Boolean
  
  BorderD10 = False
  BorderE10 = False
  BorderT10 = False
  BorderD12 = False
  BorderE12 = False
  BorderT12 = False
  
  Set oEnum = ActiveModelReference.Scan()

  While oEnum.MoveNext
        ElementCounter = ElementCounter + 1
        Set oElem = oEnum.Current
        If oElem.IsCellElement Then
          Set oCellElem = oElem
          Select Case oCellElem.Name
            Case "BDR-D10"
              'MsgBox "D10"
              Set BdrObject = oCellElem
              BorderD10 = True
            Case "BDR-D12"
              'MsgBox "D12"
              Set BdrObject = oCellElem
              BorderD12 = True
            Case "BDR-E10"
              'MsgBox "E10"
              Set BdrObject = oCellElem
              BorderE10 = True
            Case "BDR-E12"
              'MsgBox "E12"
              Set BdrObject = oCellElem
              BorderE12 = True
            Case "BDR-T10"
              'MsgBox "T10"
              Set TbdrObject = oCellElem
              BorderT10 = True
            Case "BDR-T12"
              'MsgBox "T12"
              Set TbdrObject = oCellElem
              BorderT12 = True
            Case Else
              'do nothing
           End Select
        End If
   Wend
   
   If ignoreT = True Then
     'MsgBox "ignore T"
       If BorderE10 Or BorderD10 Or BorderD12 Or BorderE12 Then
          Set GetBorder = BdrObject
          Exit Function
         Else
          MsgBox "No D or E Borders in this file"
       End If
   End If
   
   If ignoreT = False Then
       If BorderT10 Or BorderT12 Then
          Set GetBorder = TbdrObject
          'MsgBox "T border takes priority"
         ElseIf BorderE10 Or BorderD10 Or BorderD12 Or BorderE12 Then
          Set GetBorder = BdrObject
          Exit Function
         Else
          MsgBox "No D or E Borders in this file"
       End If
   End If
 
End Function

