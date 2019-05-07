VERSION 5.00
Begin VB.Form locform 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose A New .DS File Location..."
   ClientHeight    =   2955
   ClientLeft      =   3495
   ClientTop       =   3015
   ClientWidth     =   3870
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   555
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   90
      Width           =   915
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   1
      Top             =   2250
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "Note: Select a folder that contains folders of .DS files"
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   2700
      Width           =   4110
   End
End
Attribute VB_Name = "locform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  DS_PATH = Dir1.Path
  If Right(DS_PATH, 1) = "\" Then DS_PATH = Left(DS_PATH, Len(DS_PATH) - 1)
  Unload locform
End Sub


Private Sub Command2_Click()
  Unload locform
End Sub


Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub


Private Sub Form_Load()
  On Error Resume Next
  
  Drive1.Drive = Left(DS_PATH, 1)
  Dir1.Path = DS_PATH
End Sub



