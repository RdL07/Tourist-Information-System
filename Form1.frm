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
   Begin VB.Menu b 
      Caption         =   "&Add User"
   End
   Begin VB.Menu c 
      Caption         =   "Admin &Control"
   End
   Begin VB.Menu d 
      Caption         =   "Log &In"
   End
   Begin VB.Menu g 
      Caption         =   "Log &Out"
   End
   Begin VB.Menu a 
      Caption         =   "&Search"
   End
   Begin VB.Menu e 
      Caption         =   "&My Plan"
   End
   Begin VB.Menu f 
      Caption         =   "&Tour Planner"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public username As String
Public admin As Integer
Public login_status As Integer
Private Sub a_Click()
Form2.Show
End Sub

Private Sub b_Click()
AddUser.Show
End Sub

Private Sub c_Click()

'If admin = 0 Then
'MsgBox "You Are Not An Admin!", vbCritical, "Admin Status"
'Else
Adminp.Show
Form1.Hide
'End If

End Sub

Private Sub Command1_Click()
Display.Show
End Sub

Private Sub d_Click()

SignIn.Show
c.Enabled = True
f.Enabled = True

End Sub

Private Sub e_Click()

MyPlan.Show
Form1.Hide

End Sub

Private Sub Form_Load()

'c.Enabled = False
g.Enabled = False
f.Enabled = False
login_status = 0
username = ""
d.Enabled = True
    
End Sub

Private Sub g_Click()
    
    MsgBox "Successfully Logged Out!", vbInformation, "Log Out"
    login_status = 0
    username = ""
    d.Enabled = True
    g.Enabled = False
    'c.Enabled = False
    f.Enabled = False
    
End Sub
