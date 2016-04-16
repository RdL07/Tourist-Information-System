VERSION 5.00
Begin VB.Form Display 
   Caption         =   "Display"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form3"
   ScaleHeight     =   5400
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label14 
      Caption         =   "Tourist Attractions"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   4680
      Width           =   4815
   End
   Begin VB.Label Label13 
      Caption         =   "Best Time to Visit "
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label Label12 
      Caption         =   "Places  to eat"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label Label11 
      Caption         =   "Hotels"
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   2520
      Width           =   5415
   End
   Begin VB.Label Label10 
      Caption         =   "Average temaparature"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label9 
      Caption         =   "Temp_S"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "Temparature during winters"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label7 
      Caption         =   "Nearest Airport"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Nearest Bus Terminus"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Nearest Rail Station"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Type"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Places to visit nearby"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "State"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Region"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim pid As Integer
Private Sub Form_Load()
str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\SOURISH\Desktop\Tourist-Information-System\Database4.mdb;Persist Security Info=False"
cn.ConnectionString = str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "Table1"
p_id = 7
strsql1 = "select * from Table1 where P_ID = " + CStr(p_id)
rs.Open strsql1
    With rs

       Label1.Caption = Label1.Caption + " = " + !R_NAME
       Label2.Caption = Label2.Caption + " = " + !STATE_NAME
       Label3.Caption = Label3.Caption + " = " + !P_NAME
       Label4.Caption = Label4.Caption + " = " + !TYPE_NAME
       Label5.Caption = Label5.Caption + " = " + !RAIL
       Label6.Caption = Label6.Caption + " = " + !BUS
       Label7.Caption = Label7.Caption + " = " + !AIRPORT
       Label8.Caption = Label8.Caption + " = " + !TEMP_W
       Label9.Caption = Label9.Caption + " = " + !TEMP_S
       Label10.Caption = Label10.Caption + " = " + !TEMP_M
       Label11.Caption = Label11.Caption + " = " + !HOTEL
       Label12.Caption = Label12.Caption + " = " + !RESTAURANT
       Label13.Caption = Label13.Caption + " = " + !VISIT
       Label14.Caption = Label14.Caption + " = " + !PLACE
       
End With
rs.Close
End Sub

