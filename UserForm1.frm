VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2484
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5985
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
  
End Sub

Private Sub CommandButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Call custom_KeyDown(KeyCode, Shift)
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub CommandButton2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Call custom_KeyDown(KeyCode, Shift)
End Sub

Private Sub custom_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Dim Icount As Integer
  Dim TooHigh As Integer
  Const TooLow = 0
  Const No_Selection = -1
  Const EnterKey = 13
  
  Select Case KeyCode
     Case EnterKey
       Label1.Caption = "Pressed ENTER"
     Case vbKeyF2
       Label1.Caption = "F2 pressed"
     Case vbKeyUp
       Label1.Caption = "Up arrow"
       Icount = ListBox1.ListIndex
       TooHigh = ListBox1.ListCount - 1
       
       'MsgBox str(icount)
        Select Case Icount
          Case No_Selection
            'Select first item
            ListBox1.ListIndex = 0
          Case TooLow
            'Wrap back to first item
            ListBox1.ListIndex = ListBox1.ListCount - 1
          Case Else
            ListBox1.ListIndex = Icount - 1
        End Select
    Case vbKeyDown
     Label1.Caption = "down arrow"
     Icount = ListBox1.ListIndex
     TooHigh = ListBox1.ListCount - 1
     'MsgBox str(icount)
        Select Case Icount
          Case No_Selection
            'Select first item
            ListBox1.ListIndex = 0
          Case TooHigh
            'Wrap back to first item
            ListBox1.ListIndex = 0
          Case Else
            ListBox1.ListIndex = Icount + 1
        End Select
   Case Else
    'do nothing
  End Select
End Sub

Private Sub ListBox1_Click()
  'Label1.Caption = "you clicked" & " " & ListBox1.Value & "listed at:" & str(ListBox1.ListIndex)
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Call custom_KeyDown(KeyCode, Shift)
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
  ListBox1.AddItem "first"
  ListBox1.AddItem "second"
  ListBox1.AddItem "third"
  ListBox1.AddItem "fourth"
  ListBox1.AddItem "fifth"
  ListBox1.AddItem "sixth"
  ListBox1.AddItem "seventh"
  
End Sub
