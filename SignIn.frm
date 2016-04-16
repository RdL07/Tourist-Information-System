VERSION 5.00
Begin VB.Form SignIn 
   Caption         =   "Form3"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form3"
   ScaleHeight     =   4575
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log In"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "SignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim strsql1 As String
Dim pid As Integer



Private Sub Command1_Click()
str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\SOURISH\Desktop\Tourist-Information-System\Database4.mdb;Persist Security Info=False"
cn.ConnectionString = str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "Table2"
strsql1 = "select * from Table2 where UNAME LIKE '%" & Text1.Text & "%'"
rs.Open strsql1
With rs
    On Error GoTo eh1246:
    If !Password = Text2.Text Then
        Form1.username = Text1.Text
        
        Form1.login_status = 1
        rs.Close
        cn.Close
        Form1.d.Enabled = False
        Form1.g.Enabled = True
        SignIn.Hide
    End If
End With
If 1 = 2 Then
eh1246: Text1.Text = "Failed"
End If
End Sub

