VERSION 5.00
Begin VB.Form Display 
   Caption         =   "Display"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form3"
   ScaleHeight     =   4665
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Region"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p_id As Integer
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim pid As Integer
Private Sub Form_Load()
str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\ABC\Desktop\Tourist-Information-System\Database4.mdb;Persist Security Info=False"
cn.ConnectionString = str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "Table1"
rs.Open


End Sub

Private Sub Label1_Click()
p_id = 1
rs.Find ("P_ID= pid  ")
rs.Find
Label1.Caption = rs.Fields(1).Value
End Sub
