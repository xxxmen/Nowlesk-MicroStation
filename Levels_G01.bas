Attribute VB_Name = "Levels_G01"
Option Base 1
Option Explicit

Sub test_FoundAllLevels()
  Dim GotAllLevels As Boolean
  GotAllLevels = FoundAllLevels
  
    If GotAllLevels = False Then
       MsgBox "LEVELS"  'LEVELS PROBLEM
      Else
       MsgBox ""  'ALL LEVELS IN SUBSTATION.LEVELS WERE FOUND, SO RETURN NOTHING
    End If
  
End Sub
Function FoundAllLevels() As Boolean  'return True if all levels found

  Dim myLevel As Level
  Dim LevelCounter As Integer
  Dim ICountLevels As Integer
  Dim Icount As Integer
  Dim ICountChange As Boolean
  
  
  Dim strLevel(23) As String
  strLevel(1) = "Text"
  strLevel(2) = "Property line"
  strLevel(3) = "Backcircle"
  strLevel(4) = "Border-titleblock"
  strLevel(5) = "DIMENSIONS"
  strLevel(6) = "New or Revisions"
  strLevel(7) = "Baselines"
  strLevel(8) = "Fence"
  strLevel(9) = "Removal or Abandoned"
  strLevel(10) = "Contours 1 ft"
  strLevel(11) = "Contours 5 ft"
  strLevel(12) = "Liner Seal to Concrete"
  strLevel(13) = "Liner Extent"
  strLevel(14) = "Notes and References"
  strLevel(15) = "Material Item"
  strLevel(16) = "Vendor"
  strLevel(17) = "Design Master(Red)"
  strLevel(18) = "Existing"
  strLevel(19) = "Mark List"
  strLevel(20) = "Default"
  strLevel(21) = "Fence Corners"
  strLevel(22) = "Centerlines"
  strLevel(23) = "Foundations"
  
  Icount = 1
  ICountLevels = 0
  'ICountChange = False
  

  For Each myLevel In ActiveDesignFile.Levels
    
        For Icount = 1 To UBound(strLevel)
          If myLevel.Name = strLevel(Icount) Then
             Debug.Print myLevel.Name
             ICountLevels = ICountLevels + 1
             'ICountChange = True
             Exit For
          Else
             'do nothing
             'ICountChange = False
          End If
        Next Icount
            
                'test to see if ICountLevels remained the same, if so then
                'the level tested wasn't a Good level, so try to delete it
                'Test to see if level might still be in use
                'If ICountChange = False Then
                
                   ' If myLevel.IsInUse Then
                          'do nothing, because level is being used
                       ' Else        'delete level
                           '*****the following code gave error "Level Id is invalid"
                           ' ActiveDesignFile.DeleteLevel myLevel
                           ' ActiveDesignFile.Levels.Rewrite
                    ' End If
                 'End If
  Next
   
  If ICountLevels = UBound(strLevel) Then
    FoundAllLevels = True
  Else
    FoundAllLevels = False
  End If
 
End Function

