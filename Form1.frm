VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Display"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Menu a 
      Caption         =   "&Search"
   End
   Begin VB.Menu b 
      Caption         =   "Add User"
   End
   Begin VB.Menu c 
      Caption         =   "Edit Info"
   End
   Begin VB.Menu d 
      Caption         =   "Log In"
   End
   Begin VB.Menu e 
      Caption         =   "My Plan"
   End
   Begin VB.Menu f 
      Caption         =   "Tour Planner"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click()
Form2.Show
End Sub

Private Sub b_Click()
AddUser.Show
End Sub

Private Sub Command1_Click()
Display.Show
End Sub

Private Sub e_Click()

End Sub
