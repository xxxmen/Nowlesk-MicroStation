Attribute VB_Name = "modNextFile"
Option Base 1
Declare Function mdlDialog_fileOpen Lib _
"stdmdlbltin.dll" (ByVal _
 FileName As String, ByVal rFileH As Long, ByVal _
 resourceId As Long, ByVal suggestedFileName As String, _
 ByVal filterString As String, _
 ByVal defaultDirectory As String, _
 ByVal titleString As String) As Long
Sub NextFile()
  frmNextFile.show vbModeless
End Sub
Sub TestInsertFileLinesToArray()
   Dim fpath As String
   fpath = GetFileListPath
   InsertFileLinesToArray (fpath)
End Sub
Sub InsertFileLinesToArray(FilePath As String)
  'this still needs to be updated
  'the routine is to be used to load an array and then to load a list box
  'will use a two-dimensional array so that a number is associated with each
  'file name, so that the file can be opened and the current file can be saved to
  'a text file ... this has yet to be coded.

  Dim myArray() As String    ' Declare dynamic array.
  Dim FileToOpen As String
  Dim I As Integer
  Dim arraySize As Integer
  Dim x As Integer
  Dim a As Integer
  Dim batchfile As String
  
  'clear ListBox1 before adding items to it
  frmNextFile.ListBox1.Clear
    
  Dim FFile As Long
  FFile = FreeFile
  'BatchFile = "C:\filelist.txt"
  batchfile = FilePath
  Open batchfile For Input As #FFile
  I = 1
  While EOF(FFile) = False
     Line Input #FFile, FileToOpen
     'MsgBox FileToOpen
     'insert the files to an array
        ReDim Preserve myArray(I)      ' Re-allocate
        myArray(I) = FileToOpen    ' Initialize array.
        'MsgBox "array has:" & myArray(i)
        'frmNextFile.ListBox1.AddItem (myArray(i))
        I = I + 1
  Wend
     
     'loads all files
     frmNextFile.ListBox1.List() = myArray
     
     
    ' x = UBound(myArray)
     'a = LBound(myArray)
     'MsgBox Str(x) & " " & Str(a)
     'close the file
     Close FFile
End Sub



Function OpenFile(FileName As String, EditMode As Boolean) As Boolean
  Application.OpenDesignFile FileName, EditMode
End Function

Sub TestPickAFolder()
  Dim Folder As String
  Dim path As String
  Folder = PickAFolder
  MsgBox RootFolder
End Sub

Sub TestDesktopPathFunction()
   Dim strLine As String
   Dim strPath As String
   strPath = DesktopPath
   MsgBox strPath
End Sub


Function DesktopPath() As String
   Dim objFolders As Object
   Set objFolders = CreateObject("wScript.Shell").specialfolders
   DesktopPath = objFolders("desktop")
End Function


Function fileNamesInTextFile()
    Dim FilePath As String
    Dim FileName As String
    Dim First As Boolean
    Dim count As Integer
        
    Dim Folderpath As String
    Dim myFSO As New Scripting.FileSystemObject
    Dim myFolder As Scripting.Folder
    Dim myFile As Scripting.File
    Dim RootFolder As String
    'RootFolder = InputBox("Enter Root Folder:")
    RootFolder = PickAFolder(Folderpath)
    
    Set myFolder = myFSO.GetFolder(RootFolder)
    First = True
    count = 1
    For Each myFile In myFolder.Files
       Select Case UCase(Right(myFile.Name, 3))
         Case "DGN"
           If First = True Then
             Open FilePath For Output As #1
               Print #1, Str(count) & myFile.path
               First = False
             Close #1
           Else
             Open FilePath For Append As #1
               Print #1, Str(count) & " " & Right(myFile.path, 14)
             Close #1
           Else
             Open FilePath For Append As #1
               Print #1, Str(count) & " " & Right(myFile.path, 14)  'myFile.Path
             Close #1
           End If
           count = count + 1
    End Select
    Next
       
End Function

Sub SelectDTopFile()

  Dim fname As String
  Dim strPath As String
  Dim FilesFolder As String
  Dim objFolders As Object
  Set objFolders = CreateObject("wScript.Shell").specialfolders
   
  DTopPath = objFolders("desktop")
  
  strPath = DTopPath & "\"      'uses function in this module, string path could be hard coded
  
  strPath = strPath & "Filelists"
  If (Dir(strPath, vbDirectory) <> "") Then
     'do nothing
    Else
      MkDir (strPath)
  End If
  strPath = strPath & "\"
    
  'fname = SelectFile(strPath, "*.txt", "filelist.txt", "Select the file names file")
  'fname = SelectFile(strPath, "*.txt", "key-ins.txt")
  'MsgBox strPath
End Sub

Sub testselectFolder()
  Dim T As String
  T = SelectDGNFolder
  T = T & "did it show up"
  MsgBox T
End Sub

Function SelectDGNFolder() As String
  'This function calls the SelectFile function listed below
  
  Dim fname As String
  Dim strPath As String
  Dim PPath As String
  Dim TitleInfo As String
  Dim AFolderName As String
  Dim BSlash As Integer
  
  TitleInfo = "To Select a Folder, Select a File inside of a Folder!"
    
  PPath = "P:\Active Projects\PGE\Substation\"
  
  strPath = PPath      'uses function in this module, string path could be hard coded
  
  fname = SelectFile(strPath, "*.dgn", "To select a folder, select a file inside a folder", TitleInfo)
  'MsgBox fname
  
  If fname = "" Then
    MsgBox "you did not select a file!"
    'clear the list box
    frmNextFile.ListBox1.Clear
    Exit Function
  End If
  
  BSlash = RightMostBackSlash(fname)
  
  AFolderName = FolderName(BSlash, fname)
  SelectDGNFolder = AFolderName
    
End Function
Function SelectFile(strStartingPath As String, strFilter As String, strSuggFName As String, TitleText As String) As String
  'this subroutine requires the declaration statement at the top of this module
  'This routine uses the function SelectDGNFolder listed above
  Dim strFName As String
  Dim lngfhandle As Long
  Dim lngrid As Long
  Dim retVal As Long
  Dim strPath As String
  strFName = Space(255)
  retVal = mdlDialog_fileOpen(FileName:=strFName, rFileH:=lngfhandle, resourceId:=lngrid, _
                              suggestedFileName:=strSuggFName, filterString:=strFilter, defaultDirectory:=strStartingPath, _
                              titleString:=TitleText)
                              
  Select Case retVal
     Case 0  'Open
       strFName = Left(strFName, InStr(1, strFName, Chr(0)) - 0.1)
       'MsgBox "File Selected:" & vbCr & strFName
     Case 1  'Cancel
       MsgBox "No File Selected."
       strFName = ""
     End Select
  SelectFile = strFName
End Function
Function RightMostBackSlash(strPath As String) As Integer
  'This function is used by the SelectDGNFolder listed above
  Dim count As Integer
  Dim LeftPart As String
  Dim RightPart As String
  count = 0
  
  'find the right most "\" backslash
  While LeftPart <> "\"
    count = count + 1
    RightPart = Right(strPath, count)
    'Debug.Print RightPart
    LeftPart = Left(RightPart, 1)
    'Debug.Print LeftPart
  Wend
  'MsgBox Str(count)
  RightMostBackSlash = count
     
End Function
Function FolderName(BackSlashPos As Integer, PathAndFileName As String) As String
  'I don't think this function will be used
  Dim Folderpath As String
  Dim FileName As String
  Dim intBSlashLoc As Integer
  
  Trim (PathAndFileName)
  slen = Len(PathAndFileName)
  
  Folderpath = Left(PathAndFileName, slen - BackSlashPos + 1)
  'FileName = Right(PathAndFileName, intBSlashLoc - 1)
  'MsgBox "Folder Path:" & " " & folderPath & vbLf & "File Name:" & " " & FileName
  
  FolderName = Folderpath
End Function

Sub test_FileCreate()
   Call FileCreate("C:\Users\knowles_keith\Desktop\Microstation_test_Folder\")
End Sub

Function GetFileListPath() As String
   Dim DesktopPath As String
   Dim objFolders As Object
  
   Set objFolders = CreateObject("wScript.Shell").specialfolders
   DesktopPath = objFolders("desktop")
   GetFileListPath = DesktopPath & "\Filelists\filelist.txt"
   
End Function

Sub FileCreate(Folder As String)
  Dim textfile As String
  Dim Folderpath As String
  Dim DesktopPath As String
  Dim objFolders As Object
  
   Set objFolders = CreateObject("wScript.Shell").specialfolders
   DesktopPath = objFolders("desktop")
    
   textfile = GetFileListPath
   
   'folderPath = "C:\Users\knowles_keith\Desktop\Microstation_test_Folder\"
   'Folderpath = "C:\Users\knowles_keith\Desktop\Microstation_test_Folder\"
   
   Folderpath = Folder
   
   'folderPath = "P:\Active Projects\PGE\Substation\6446 BELL\2000 Substation\2300 Engineering\2310 Electrical\2311 Drawings\Indoor\"
   Call TextFileCreate(textfile, Folderpath)
  
End Sub

Sub TextFileCreate(textfile As String, Folderpath As String)
'*************************************************************
'Good - code is used in button, "Change Folder Path"

  Dim N As Integer
  Dim I As Integer
  Dim MyPath As String
  Dim MyName As String
  Dim FileNames() As String
  Dim counter As Integer
  
  MyPath = Folderpath
    
  'Display the names in the directory
    MyName = Dir(MyPath)    'Retrieve the first entry.
    I = 0
    Do While MyName <> ""    ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        If MyName <> "." And MyName <> ".." Then
          If Right(MyName, 4) = ".dgn" Then
            'Debug.Print MyName    ' Display entry only if it
             I = I + 1
             ReDim Preserve FileNames(I)
             FileNames(I) = MyName
             Debug.Print FileNames(I)
          End If
        End If    ' it represents a directory.
        MyName = Dir    ' Get next entry.
    Loop
    
    'create textfile or over-write the existing file
    'textfile

    Open textfile For Output As #1  'over-writes and/or creates new file
    Print #1, FileNames(1)
    Close #1
    
    If UBound(FileNames) <= 1 Then
        Exit Sub
    End If
    counter = 2
      Open textfile For Append As #1
    For counter = 2 To UBound(FileNames)
      Print #1, FileNames(counter)
    Next
    Close #1
    
 '*****************************************************************************
End Sub

Sub filepathsTxtFile()  'needs to take in the file path name, may need to
  Dim textfile As String
  Dim N As Integer
  Dim I As Integer
  Dim MyPath As String
  Dim MyName As String
  Dim FolderString() As String
  
  'This folder should be passed to the routine by arguments
  Folderpath = "C:\Users\knowles_keith\Desktop\Microstation_test_Folder\"
  
  'get file --- usually located on the desktop folder with the name filelist.txt
  'textfile = "c:\filelist.txt"

  MyPath = Folderpath
    
  'Display the names in the directory
    MyName = Dir(MyPath)    ' Retrieve the first entry
    Do While MyName <> ""    ' Start the loop
        ' Ignore the current directory and the encompassing directory
        If MyName <> "." And MyName <> ".." Then
          If Right(MyName, 4) = ".dgn" Then
            Debug.Print MyName    ' Display entry only if it
          End If
        End If    ' it represents a directory.
        MyName = Dir    ' Get next entry.
    Loop
End Sub

Sub test_SplitFolderFromFileName()

  Dim Folderpath As String
  Dim FileName As String
  Dim strPath As String
  Dim intBSlashLoc As Integer
  
  strPath = "C:\Indoor\123.dgn"
  
  Trim (strPath)
  slen = Len(strPath)
  
  'Get the position of the right most back slash in file path
  intBSlashLoc = RightMostBackSlash(strPath)
  
  Folderpath = Left(strPath, slen - intBSlashLoc + 1)
  FileName = Right(strPath, intBSlashLoc - 1)
  MsgBox "Folder Path:" & " " & Folderpath & vbLf & "File Name:" & " " & FileName
  
End Sub

Function WritePathToFile(FullFolderPath As String) As Boolean
   Dim textfilepath As String
   Dim path As String
   Dim IsFolderThere As Boolean
      
   'get the desktop folder path
   path = DesktopPath
   path = path & "\Filelists"
   
   'check to see if folder on desktop
   If (Dir(path, vbDirectory) <> "") Then
        IsFolderThere = True
        MsgBox "folder there"
     Else
        IsFolderThere = False
        MsgBox "Folder not there"
        MkDir (path)
   End If
   
   path = path & "\path.txt"
            
   Open path For Output As #1
      Print #1, FullFolderPath
   Close #1
   
End Function

Function GetPathToFile() As String
   Dim textfilepath As String
   Dim path As String
   Dim IsFolderThere As Boolean
      
   'get the desktop folder path
   path = DesktopPath
   path = path & "\Filelists"
   
   'check to see if folder on desktop
   If (Dir(path, vbDirectory) <> "") Then
        IsFolderThere = True
        'MsgBox "folder there"
     Else
        IsFolderThere = False
        'MsgBox "Folder not there"
        MkDir (path)
   End If
   
   path = path & "\path.txt"
            
   Open path For Input As #1
      Line Input #1, FullFolderPath
   Close #1
   
   GetPathToFile = FullFolderPath
End Function
