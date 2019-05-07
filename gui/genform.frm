VERSION 5.00
Begin VB.Form genform 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generating WAV files..."
   ClientHeight    =   1170
   ClientLeft      =   3135
   ClientTop       =   3870
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   78
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   221
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton butCancel 
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
      Left            =   2325
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   705
      Width           =   795
   End
   Begin VB.Shape bar 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   165
      Top             =   240
      Width           =   15
   End
   Begin VB.Label label 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   795
      Width           =   2040
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   315
      Left            =   150
      Top             =   225
      Width           =   2970
   End
   Begin VB.Shape bargraph 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   165
      Top             =   240
      Width           =   2955
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00E0E0E0&
      Height          =   315
      Left            =   150
      Top             =   240
      Width           =   2985
   End
End
Attribute VB_Name = "genform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gencancel
Dim instance

Private Sub butCancel_Click()
  gencancel = True
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  
  If instance <> 1 Then
    instance = 1

      Fpath = DS_PATH & "\" & mainform.kitList.List(mainform.kitList.ListIndex)
      If Right(Fpath, 1) <> "\" Then Fpath = Fpath & "\"
      
      Counter = 0 'count how many for bargraph
      For itm = 0 To mainform.fileList.ListCount - 1
        If Right(mainform.fileList.List(itm), 1) <> " " Then
          Counter = Counter + 1
        End If
      Next itm
     
      If Counter > 0 Then
        progress = 0
        itm = 0

        While itm < mainform.fileList.ListCount And gencancel = False
          If Right(mainform.fileList.List(itm), 1) <> " " Then
            WAVname = Fpath & mainform.fileList.List(itm) & ".wav"
            DSname = Fpath & mainform.fileList.List(itm) & ".ds"
            label = mainform.fileList.List(itm)
            DoEvents
          
            If Dir(Fpath & "\" & WAVname) = "" Then 'generate WAV
              rtn = ds2wav(DSname, WAVname, 0)
              ''Debug.Print rtn, DSname, WAVname
            End If
    
            progress = progress + 1
            bar.Width = progress / Counter * bargraph.Width
          End If
        
          DoEvents
          'screen.mousepointer = 0
          itm = itm + 1
    
        Wend
 
    End If
  
    genform.Hide
    DoEvents
    instance = 0
    
    Unload genform
  End If

End Sub

