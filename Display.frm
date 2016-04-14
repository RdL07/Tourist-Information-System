VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   4560
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Display.frx":0000
      OLEDBString     =   $"Display.frx":008E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
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
str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\SOURISH\Desktop\Tourist-Information-System\Database4.mdb;Persist Security Info=False"
cn.ConnectionString = str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "Table1"


End Sub

Private Sub Label1_Click()
p_id = 9
strsql1 = "select * from Table1 where P_ID = " + CStr(p_id)
rs.Open strsql1
    With rs

       Label1.Caption = !R_NAME

    End With
End Sub
