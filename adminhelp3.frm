VERSION 5.00
Begin VB.Form adminhelp3 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Tourist Spots"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11175
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   7320
      TabIndex        =   15
      Text            =   "Temperature in Summer..."
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Text            =   "Temperature in Monsoon..."
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Text            =   "Temperature in Winter..."
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Text            =   "Airport Nearby..."
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Text            =   "Bus Terminus Nearby..."
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      Text            =   "Place Name..."
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Text            =   "Region Name..."
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Text            =   "Hotels..."
      Top             =   4200
      Width           =   7935
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Text            =   "State Name..."
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Text            =   "Type..."
      Top             =   2040
      Width           =   7935
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Text            =   "Railway Station Nearby..."
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Place ID..."
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Star (*) Marked Are Required Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   20
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   19
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   17
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "adminhelp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p_id As Integer
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub Command1_Click()

If Text1.Text = "Place ID..." Or Text1.Text = "" Then
MsgBox "Place ID is a Required Field", vbCritical, "Place ID"
ElseIf Text7.Text = "Region Name..." Or Text7.Text = "" Then
MsgBox "Region Name is a Required Field", vbCritical, "Region Name"
ElseIf Text6.Text = "State Name..." Or Text6.Text = "" Then
MsgBox "State Name is a Required Field", vbCritical, "State Name"
ElseIf Text8.Text = "Place Name..." Or Text8.Text = "" Then
MsgBox "Place Name is a Required Field", vbCritical, "Place Name"
ElseIf Text5.Text = "Type..." Or Text5.Text = "" Then
MsgBox "Type is a Required Field", vbCritical, "Type"
Else
rs.Find ("P_ID= '" & Text1.Text & "'")
If rs.EOF Then
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text7.Text
rs.Fields(2) = Text6.Text
rs.Fields(3) = Text8.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Text4.Text
rs.Fields(6) = Text10.Text
rs.Fields(7) = Text9.Text
rs.Fields(8) = Text3.Text
rs.Fields(9) = Text11.Text
rs.Fields(10) = Text14.Text
rs.Fields(11) = Text13.Text
rs.Save
Else
MsgBox "Enter Different Place ID", vbCritical, "Place ID Already Exists"
End If
End If

End Sub

Private Sub Command2_Click()

adminhelp3.Hide
Adminp.Show

End Sub

Private Sub Command3_Click()

If Text1.Text = "" Or Text1.Text = "Place ID..." Then
MsgBox "Place ID is a Required Field", vbCritical, "Place ID"
Else
rs.Find ("P_ID= '" & Text1.Text & "'")
If rs.EOF Then
MsgBox "Place ID Not Found!", vbCritical, "Not Found"
rs.MoveFirst
Else
If Text7.Text = "Region Name..." Or Text7.Text = "" Then
MsgBox "Region Name is a Required Field", vbCritical, "Region Name"
ElseIf Text6.Text = "State Name..." Or Text6.Text = "" Then
MsgBox "State Name is a Required Field", vbCritical, "State Name"
ElseIf Text8.Text = "Place Name..." Or Text8.Text = "" Then
MsgBox "Place Name is a Required Field", vbCritical, "Place Name"
ElseIf Text5.Text = "Type..." Or Text5.Text = "" Then
MsgBox "Type is a Required Field", vbCritical, "Type"
Else
rs.Update
rs.Fields(1) = Text7.Text
rs.Fields(2) = Text6.Text
rs.Fields(3) = Text8.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Text4.Text
rs.Fields(6) = Text10.Text
rs.Fields(7) = Text9.Text
rs.Fields(8) = Text3.Text
rs.Fields(9) = Text11.Text
rs.Fields(10) = Text14.Text
rs.Fields(11) = Text13.Text
rs.Save
End If
End If
End If

End Sub

Private Sub Command4_Click()

If Text1.Text = "Place ID..." Or Text1.Text = "" Then
MsgBox "Place ID is a Required Field", vbCritical, "Place ID"
Else
rs.Find ("P_ID= '" & Text1.Text & "'")
If rs.EOF Then
MsgBox "Place ID Not Found!", vbCritical, "Not Found"
rs.MoveFirst
Else
rs.Delete
MsgBox "Deletion Done Successfully", vbInformation, "Remove"
rs.Save
End If
End If

End Sub

Private Sub Form_Load()
str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\SOURISH\Desktop\Tourist-Information-System\Database4.mdb;Persist Security Info=False"
cn.ConnectionString = str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "Table1"
rs.Open

End Sub


Private Sub Text14_Click()

Text14.Text = ""

End Sub

Private Sub Text14_LostFocus()
If Text14.Text = "" Then
Text14.Text = "Temperature in Summer..."
End If
End Sub

Private Sub Text13_Click()

Text13.Text = ""

End Sub

Private Sub Text13_LostFocus()
If Text13.Text = "" Then
Text13.Text = "Hotels..."
End If
End Sub
Private Sub Text3_Click()

Text3.Text = ""

End Sub

Private Sub Text3_LostFocus()
If Text3.Text = "" Then
Text3.Text = "Temperature in Winter..."
End If
End Sub

Private Sub Text11_Click()

Text11.Text = ""

End Sub

Private Sub Text11_LostFocus()
If Text11.Text = "" Then
Text11.Text = "Temperature in Monsoon..."
End If
End Sub
Private Sub Text10_Click()

Text10.Text = ""

End Sub

Private Sub Text10_LostFocus()
If Text10.Text = "" Then
Text10.Text = "Bus Terminus Nearby..."
End If
End Sub

Private Sub Text9_Click()

Text9.Text = ""

End Sub

Private Sub Text9_LostFocus()
If Text9.Text = "" Then
Text9.Text = "Airport Nearby..."
End If
End Sub
Private Sub Text5_Click()

Text5.Text = ""

End Sub

Private Sub Text5_LostFocus()
If Text5.Text = "" Then
Text5.Text = "Type..."
End If
End Sub

Private Sub Text4_Click()

Text4.Text = ""

End Sub

Private Sub Text4_LostFocus()
If Text4.Text = "" Then
Text4.Text = "Railway Station Nearby..."
End If
End Sub

Private Sub Text6_Click()

Text6.Text = ""

End Sub

Private Sub Text6_LostFocus()
If Text6.Text = "" Then
Text6.Text = "State Name..."
End If
End Sub

Private Sub Text8_Click()

Text8.Text = ""

End Sub

Private Sub Text8_LostFocus()
If Text8.Text = "" Then
Text8.Text = "Place Name..."
End If
End Sub

Private Sub Text1_Click()

Text1.Text = ""

End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then
Text1.Text = "Place ID..."
End If
End Sub

Private Sub Text7_Click()

Text7.Text = ""

End Sub

Private Sub Text7_LostFocus()
If Text7.Text = "" Then
Text7.Text = "Region Name..."
End If
End Sub



