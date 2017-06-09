Attribute VB_Name = "testCodeNotUsed"
Sub TestGetFiles()
   ' Call to test GetFiles function.
   Dim PPath As String
   Dim BellPath As String
   Dim dctDict As Dictionary
   Dim varItem As Variant
   Dim GetTempDir As String
   'GetTempDir = "C:\Users\knowles_keith\Desktop\Microstation_test_Folder"
   PPath = "P:\Active Projects\PGE\Substation\"
   BellPath = "6446 BELL\2000 Substation\2300 Engineering\2310 Electrical\2311 Drawings\Indoor"
   GetTempDir = PPath & BellPath
   'Create new dictionary.
   Set dctDict = New Dictionary
   ' Call recursively, return files into Dictionary object.
   If GetFiles(GetTempDir, dctDict, False) Then
      ' Print items in dictionary.
      For Each varItem In dctDict
         Debug.Print varItem
      Next
   End If
End Sub
Function GetFiles(strPath As String, _
                dctDict As Dictionary, _
                Optional blnRecursive As Boolean) As Boolean
            
   ' This procedure returns all the files in a directory into
   ' a Dictionary object. If called recursively, it also returns
   ' all files in subfolders.
   
   Dim fsoSysObj      As FileSystemObject
   Dim fdrFolder      As Folder
   Dim fdrSubFolder   As Folder
   Dim filFile        As File
   
   ' Return new FileSystemObject.
   Set fsoSysObj = New FileSystemObject
   
   On Error Resume Next
   ' Get folder.
   Set fdrFolder = fsoSysObj.GetFolder(strPath)
   If Err <> 0 Then
      ' Incorrect path.
      GetFiles = False
      GoTo GetFiles_End
   End If
   On Error GoTo 0
   
   ' Loop through Files collection, adding to dictionary.
   For Each filFile In fdrFolder.Files
      dctDict.Add filFile.path, filFile.path
   Next filFile

   ' If Recursive flag is true, call recursively.
   If blnRecursive Then
      For Each fdrSubFolder In fdrFolder.SubFolders
         GetFiles fdrSubFolder.path, dctDict, True
      Next fdrSubFolder
   End If

   ' Return True if no error occurred.
   GetFiles = True
   
GetFiles_End:
   Exit Function
End Function


Sub KJK()
 Dim ob As Application
  

  Application.ActiveDesignFile.TotalEditingTime
  
  
End Sub

Sub testScanFilter()
  Dim rng As Range3d
  Dim pnt3D As Point3d
  
  Dim mycell As CellInformation
  Dim myCellEnum As CellInformationEnumerator
  
  Dim myElem As Element
  Dim myEnum As ElementEnumerator
  Dim myFilter As New ElementScanCriteria
  Dim ElementCounter As Long
  Dim myCollection As New Collection
  'myFilter.ExcludeAllTypes
  myFilter.ExcludeAllLevels
  'myFilter.ExcludeAllColors
  'myFilter.IncludeType msdElementTypeText
  'myFilter.IncludeType msdElementTypeTextNode
  
  myFilter.IncludeLevel ActiveDesignFile.Levels("Border-titleblock")
  myFilter.IncludeLevel ActiveDesignFile.Levels("Border and Titleblock")
  'myFilter.IncludeLevel ActiveDesignFile.Levels("Level 1")
  'myFilter.IncludeLevel ActiveDesignFile.Levels("Existing")
  
  'myFilter.IncludeOnlyCell "BDR-D10"
  'myFilter.IncludeColor 4
  Set myEnum = ActiveModelReference.Scan(myFilter)

  While myEnum.MoveNext
    ElementCounter = ElementCounter + 1
    Set myElem = myEnum.Current
    myCollection.Add myElem
    MsgBox myElem.AsCellElement.Name & " " & "origin: " & vbLf & _
                  myElem.AsCellElement.Origin.x & ", " & myElem.AsCellElement.Origin.Y
       
    'MsgBox myElem.AsCellElement.Origin.x & " " & myElem.AsCellElement.Origin.Y
    'MsgBox myElem.AsCellElement.IsGraphical
    rng = myElem.AsCellElement.Range
    
    'MsgBox "x: " & Str(pnt3D.x = rng.High.x)
    pnt3D.x = rng.High.x
    pnt3D.Y = rng.High.Y
    MsgBox "High X: " & Str(pnt3D.x) & "High Y: " & Str(pnt3D.Y)
     pnt3D.x = rng.Low.x
    pnt3D.Y = rng.Low.Y
    MsgBox "Low X: " & Str(pnt3D.x) & "Low Y: " & Str(pnt3D.Y)
  Wend
  MsgBox ElementCounter & " elements found."
End Sub

