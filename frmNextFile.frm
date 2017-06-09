VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNextFile 
   Caption         =   "Next File"
   ClientHeight    =   7695
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   2460
   OleObjectBlob   =   "frmNextFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFolderPath_Click()
    'Get the folder of drawings and insert the path into the form
    lblFolderPath.Caption = modNextFile.SelectDGNFolder
    
    If lblFolderPath.Caption = "" Then
      MsgBox "You pressed Cancel, or you did not select a file inside of the folder."
      Exit Sub
    End If
        
    'save path to desktop\Filelists\path.txt
    modNextFile.WritePathToFile FullFolderPath:=lblFolderPath.Caption
            
    'create a text file of all of the DGN files in the folder
    Call modNextFile.FileCreate(lblFolderPath)
    
    'Create an array from the text file and insert the drawing names into
    'the listbox
    Call modNextFile.InsertFileLinesToArray(modNextFile.GetFileListPath)
    
End Sub

Private Sub cmdLoadPrevious_Click()
    Dim path As String
   'get path from path.txt file
    path = GetPathToFile
    
   'put path on label
   lblFolderPath.Caption = path
   
    'create a text file of all of the DGN files in the folder
    Call modNextFile.FileCreate(path)
    
    'Create an array from the text file and insert the drawing names into
    'the listbox
    Call modNextFile.InsertFileLinesToArray(modNextFile.GetFileListPath)
    
End Sub

Private Sub ListBox1_Click()

    Dim path As String
    Dim FileName As String
    Dim ReadOnly As Boolean
    
    path = lblFolderPath.Caption
    'filename = ListBox1.Selected(pvargindex)
    FileName = ListBox1.Value
    'MsgBox filename
    FileName = path & FileName
    
    If togReadOnly.Caption = "Read Only" Then
      ReadOnly = True
     Else
      ReadOnly = False
    End If
    
    modNextFile.OpenFile FileName:=FileName, EditMode:=ReadOnly
    
    'more code to make read only open with extents in view 1
    
     If ReadOnly Then
       'MsgBox "This is read only mode"
     End If
End Sub

Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  MsgBox "a key was pressed"
End Sub

Private Sub togReadOnly_Click()
    If togReadOnly.Caption = "Read Only" Then
      togReadOnly.Caption = "Edit Mode"
     Else
      togReadOnly.Caption = "Read Only"
   End If
End Sub

Private Sub UserForm_Initialize()
  togReadOnly.Caption = "Read Only"
End Sub


Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case vbKeyF2
      MsgBox "f2 pressed"
    Case Else
      KeyAscii = 0
    End Select
End Sub
