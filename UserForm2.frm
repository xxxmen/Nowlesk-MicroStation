VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6135
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5520
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim v2 As String
    
    TextBox2.Text = TextBox1.CurLine
    TextBox3.Text = TextBox1.CurX
    TextBox4.Text = TextBox1.CurTargetX
    TextBox5.Text = TextBox1.TabKeyBehavior
    
    v2 = KeyCode
    
    
    MsgBox v2
    
    
    
    
End Sub
Private Sub UserForm_Initialize()
    TextBox1.MultiLine = True

    TextBox1.Text = "Type your text here. User CTRL + ENTER to start a new line."
End Sub

