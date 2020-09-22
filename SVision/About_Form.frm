VERSION 5.00
Begin VB.Form About_Form 
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "About_Form.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Written in Visual Basic 6.0"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Legal 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Written By:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Programmer 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Version 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Program_Name 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "About_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
'----------------------
' Close About Form
'----------------------
 Unload Me
End Sub

Private Sub Form_Load()
'---------------------------------
' Set About Form Text
'---------------------------------
 Program_Name.Caption = "Stereo Vision"
 Version.Caption = "Version 0.1"
 Programmer.Caption = "Dennis C. Urie"
 Legal.Caption = "This program is FreeWare"
 
 Me.Caption = "About " & Program_Name.Caption
 
End Sub

