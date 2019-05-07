VERSION 5.00
Begin VB.Form fileform 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save As..."
   ClientHeight    =   1125
   ClientLeft      =   3135
   ClientTop       =   2310
   ClientWidth     =   3315
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1125
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSaveWAV 
      Caption         =   "Save .WAV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   3
      Top             =   720
      Value           =   1  'Checked
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2445
      TabIndex        =   2
      Top             =   660
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   1620
      TabIndex        =   1
      Top             =   660
      Width           =   720
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   2985
   End
End
Attribute VB_Name = "fileform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  
  AskFileName = LTrim(RTrim(txt))
  SaveWAV = chkSaveWAV
  Unload fileform
  
End Sub

Private Sub Command2_Click()
  AskFileName = ""
  SaveWAV = chkSaveWAV
  Unload fileform
End Sub

Private Sub Form_Load()
  chkSaveWAV = SaveWAV
  fileform.Top = mainform.Top + 0.6 * mainform.Height
  fileform.Left = mainform.Left + 0.1 * mainform.Width
  SendKeys "+{End}"
End Sub

Private Sub txt_KeyPress(keyAscii As Integer)
  
  Select Case keyAscii
    'Case Is < 32: keyAscii = 0
    Case Asc("."): keyAscii = 0
    Case Asc("\"): keyAscii = 0
    Case Asc("/"): keyAscii = 0
    Case Asc(":"): keyAscii = 0
    Case Asc("*"): keyAscii = 0
    Case Asc("?"): keyAscii = 0
    Case Asc(""""): keyAscii = 0
    Case Asc("<"): keyAscii = 0
    Case Asc(">"): keyAscii = 0
    Case Asc("|"): keyAscii = 0
    
    Case Else:  'character allowed in filenames
  End Select

End Sub

