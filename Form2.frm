VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form2"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form2"
   ScaleHeight     =   5700
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   960
      List            =   "Form2.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4680
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form2.frx":0017
      Left            =   960
      List            =   "Form2.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form2.frx":002B
      Left            =   960
      List            =   "Form2.frx":0032
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0043
      Left            =   960
      List            =   "Form2.frx":0053
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080C0FF&
      Caption         =   "SELECT      TYPE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "SELECT        CITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "SELECT STATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "SELECT  REGION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Change()
If Combo2.Text = "WEST BENGAL" Then
Combo1.Text = "EAST"
End If
End Sub
