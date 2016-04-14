VERSION 5.00
Begin VB.Form AddUser 
   Caption         =   "Add User"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form3"
   ScaleHeight     =   4155
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   3120
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Text            =   "Re enter Password"
      Top             =   2280
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   1080
      TabIndex        =   1
      Text            =   "Password"
      Top             =   1440
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Text            =   "User Name"
      Top             =   600
      Width           =   5295
   End
End
Attribute VB_Name = "AddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub
Private Sub Text2_Change()
Text2.PasswordChar = "*"
End Sub


Private Sub Text3_Click()
Text3.Text = ""
End Sub
Private Sub Text3_Change()
Text3.PasswordChar = "*"
End Sub
Private Sub Text1_LostFocus()
If Text1.Text = "" Then
    Text1.Text = "Username"
End If
End Sub
Private Sub Text2_LostFocus()
If Text2.Text = "" Then
    Text2.PasswordChar = ""
    Text2.Text = "Password"
End If
End Sub
Private Sub Text3_LostFocus()
If Text3.Text = "" Then
    Text3.PasswordChar = ""
    Text3.Text = "Re enter Password"
End If
End Sub
