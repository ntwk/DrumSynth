VERSION 5.00
Begin VB.Form mainform 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DrumSynth 2.0"
   ClientHeight    =   7110
   ClientLeft      =   1170
   ClientTop       =   1950
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "DRUMSYN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7110
   ScaleWidth      =   10230
   Begin VB.HScrollBar sliMasterdB 
      Height          =   240
      LargeChange     =   3
      Left            =   7530
      Max             =   20
      Min             =   -20
      TabIndex        =   74
      Top             =   6200
      Width           =   2000
   End
   Begin VB.ListBox kitList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2400
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   1670
   End
   Begin VB.CheckBox chkFixedRandom 
      BackColor       =   &H00BEC8BA&
      Caption         =   "always same random sequence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4710
      TabIndex        =   25
      Top             =   3465
      Value           =   1  'Checked
      Width           =   2550
   End
   Begin VB.CheckBox chkMasterOn 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7545
      TabIndex        =   42
      Top             =   5130
      Width           =   2550
   End
   Begin VB.HScrollBar sliFres 
      Height          =   240
      LargeChange     =   10
      Left            =   7515
      Max             =   99
      TabIndex        =   39
      Top             =   4020
      Width           =   2000
   End
   Begin VB.CheckBox chkHPF 
      BackColor       =   &H00D0D0C0&
      Caption         =   "High-pass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7545
      TabIndex        =   40
      Top             =   4365
      Width           =   1035
   End
   Begin VB.CommandButton butHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   61
      Top             =   6720
      Width           =   285
   End
   Begin VB.CheckBox MasterFilter 
      BackColor       =   &H00D0D0C0&
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   38
      Top             =   3735
      Width           =   2595
   End
   Begin VB.CheckBox chkOTrack2 
      BackColor       =   &H00CDC0C0&
      Caption         =   "Tracking   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3060
      TabIndex        =   19
      Top             =   5550
      Width           =   1185
   End
   Begin VB.CheckBox chkOTrack1 
      BackColor       =   &H00CDC0C0&
      Caption         =   "Tracking   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1845
      TabIndex        =   16
      Top             =   5550
      Width           =   1170
   End
   Begin VB.HScrollBar sliOParam 
      Height          =   240
      LargeChange     =   10
      Left            =   1845
      Max             =   100
      TabIndex        =   21
      Top             =   6180
      Width           =   1995
   End
   Begin VB.TextBox MasterLength 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7530
      TabIndex        =   45
      Text            =   "100.0"
      Top             =   5820
      Width           =   720
   End
   Begin VB.TextBox MasterTune 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7530
      TabIndex        =   43
      Text            =   "0.00"
      Top             =   5440
      Width           =   720
   End
   Begin VB.CheckBox chkOn 
      BackColor       =   &H00C4B8C4&
      Caption         =   "Noise Band 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   4725
      TabIndex        =   30
      Top             =   5280
      Width           =   2550
   End
   Begin VB.HScrollBar sliLev 
      Height          =   240
      Index           =   4
      LargeChange     =   9
      Left            =   4695
      Max             =   181
      Min             =   1
      TabIndex        =   31
      Top             =   5550
      Value           =   128
      Width           =   1995
   End
   Begin VB.TextBox txtNF2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4695
      TabIndex        =   32
      Text            =   "3100"
      Top             =   5850
      Width           =   765
   End
   Begin VB.HScrollBar sliNdF2 
      Height          =   240
      LargeChange     =   10
      Left            =   4695
      Max             =   100
      TabIndex        =   33
      Top             =   6195
      Value           =   40
      Width           =   2000
   End
   Begin VB.ComboBox comOW2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3045
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4860
      Width           =   795
   End
   Begin VB.ComboBox comOW1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1845
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4860
      Width           =   780
   End
   Begin VB.ComboBox comDF 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8550
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   3180
      Width           =   990
   End
   Begin VB.ComboBox comDB 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7515
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   3180
      Width           =   990
   End
   Begin VB.ListBox fileList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3570
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   1665
   End
   Begin VB.CheckBox chkOn 
      BackColor       =   &H00C0D2D2&
      Caption         =   "Distortion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   7530
      TabIndex        =   34
      Top             =   2580
      Width           =   2595
   End
   Begin VB.HScrollBar sliDist 
      Height          =   240
      LargeChange     =   3
      Left            =   7515
      Max             =   60
      TabIndex        =   35
      Top             =   2850
      Width           =   2000
   End
   Begin VB.CommandButton butLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Load"
      Height          =   360
      Left            =   30
      TabIndex        =   3
      Top             =   6165
      Width           =   810
   End
   Begin VB.CommandButton butSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      Height          =   360
      Left            =   885
      TabIndex        =   4
      Top             =   6165
      Width           =   795
   End
   Begin VB.TextBox comment 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   405
      TabIndex        =   62
      Top             =   6780
      Width           =   8460
   End
   Begin VB.HScrollBar sliNdF 
      Height          =   240
      LargeChange     =   10
      Left            =   4695
      Max             =   100
      TabIndex        =   29
      Top             =   4800
      Value           =   40
      Width           =   2000
   End
   Begin VB.TextBox txtNF 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4695
      TabIndex        =   28
      Text            =   "1200"
      Top             =   4440
      Width           =   765
   End
   Begin VB.ComboBox comOMethod 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1845
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5790
      Width           =   2010
   End
   Begin VB.TextBox txtOF2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3060
      TabIndex        =   18
      Text            =   "630"
      Top             =   5220
      Width           =   780
   End
   Begin VB.TextBox txtOF1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1845
      TabIndex        =   15
      Text            =   "315"
      Top             =   5220
      Width           =   780
   End
   Begin VB.CommandButton butPlay 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Play"
      Height          =   345
      Left            =   9330
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6720
      Width           =   840
   End
   Begin VB.HScrollBar sliNSlope 
      Height          =   240
      LargeChange     =   10
      Left            =   4695
      Max             =   100
      Min             =   -100
      SmallChange     =   2
      TabIndex        =   24
      Top             =   3150
      Width           =   2000
   End
   Begin VB.TextBox txtTPhase 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1845
      TabIndex        =   11
      Text            =   "0"
      Top             =   3810
      Width           =   780
   End
   Begin VB.HScrollBar sliTDroop 
      Height          =   240
      LargeChange     =   10
      Left            =   1845
      Max             =   100
      TabIndex        =   10
      Top             =   3510
      Width           =   1995
   End
   Begin VB.TextBox txtTF2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3075
      TabIndex        =   9
      Text            =   "120"
      Top             =   3150
      Width           =   765
   End
   Begin VB.TextBox txtTF1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1845
      TabIndex        =   8
      Text            =   "200"
      Top             =   3150
      Width           =   780
   End
   Begin VB.HScrollBar sliLev 
      Height          =   240
      Index           =   3
      LargeChange     =   9
      Left            =   4695
      Max             =   181
      Min             =   1
      TabIndex        =   27
      Top             =   4140
      Value           =   128
      Width           =   1995
   End
   Begin VB.HScrollBar sliLev 
      Height          =   240
      Index           =   2
      LargeChange     =   9
      Left            =   1845
      Max             =   181
      Min             =   1
      TabIndex        =   13
      Top             =   4560
      Value           =   128
      Width           =   1995
   End
   Begin VB.HScrollBar sliLev 
      Height          =   240
      Index           =   1
      LargeChange     =   9
      Left            =   4695
      Max             =   181
      Min             =   1
      TabIndex        =   23
      Top             =   2850
      Value           =   128
      Width           =   1995
   End
   Begin VB.HScrollBar sliLev 
      Height          =   240
      Index           =   0
      LargeChange     =   9
      Left            =   1845
      Max             =   181
      Min             =   1
      TabIndex        =   7
      Top             =   2835
      Value           =   128
      Width           =   1995
   End
   Begin VB.PictureBox envbox 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   2145
      Left            =   1755
      Picture         =   "DRUMSYN.frx":030A
      ScaleHeight     =   -104
      ScaleMode       =   0  'User
      ScaleTop        =   102
      ScaleWidth      =   8804.166
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   60
      Width           =   8400
   End
   Begin VB.CheckBox chkOn 
      BackColor       =   &H00C4B8C4&
      Caption         =   "Noise Band 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   4725
      TabIndex        =   26
      Top             =   3870
      Width           =   2565
   End
   Begin VB.CheckBox chkOn 
      BackColor       =   &H00CDC0C0&
      Caption         =   "Overtones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   1875
      TabIndex        =   12
      Top             =   4290
      Width           =   2580
   End
   Begin VB.CheckBox chkOn 
      BackColor       =   &H00BEC8BA&
      Caption         =   "Noise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   4725
      TabIndex        =   22
      Top             =   2580
      Value           =   1  'Checked
      Width           =   2565
   End
   Begin VB.CheckBox chkOn 
      BackColor       =   &H00BEC0CA&
      Caption         =   "Tone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1875
      TabIndex        =   6
      Top             =   2565
      Value           =   1  'Checked
      Width           =   2550
   End
   Begin VB.CheckBox chkOFilter 
      BackColor       =   &H00D0D0C0&
      Caption         =   "Filter overtones only "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7545
      TabIndex        =   41
      Top             =   4650
      Width           =   2430
   End
   Begin VB.CommandButton nothing 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
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
      Left            =   9495
      TabIndex        =   5
      Top             =   6765
      Width           =   615
   End
   Begin VB.Label txtMasterdB 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0 dB"
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
      Height          =   240
      Left            =   9560
      TabIndex        =   73
      Top             =   6220
      Width           =   615
   End
   Begin VB.Label txtFres 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Q = 0"
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
      Height          =   240
      Left            =   9540
      TabIndex        =   72
      Top             =   4040
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   10140
      X2              =   10140
      Y1              =   4950
      Y2              =   3660
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   10140
      X2              =   7425
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Label txtOParam 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mix"
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
      Height          =   285
      Left            =   3880
      TabIndex        =   71
      Top             =   6200
      Width           =   630
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00E0E0E0&
      X1              =   7335
      X2              =   4605
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00E0E0E0&
      X1              =   7335
      X2              =   4605
      Y1              =   5115
      Y2              =   5115
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00E0E0E0&
      X1              =   7335
      X2              =   4605
      Y1              =   6510
      Y2              =   6510
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00E0E0E0&
      X1              =   10140
      X2              =   7425
      Y1              =   3570
      Y2              =   3570
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00E0E0E0&
      X1              =   10140
      X2              =   7425
      Y1              =   6510
      Y2              =   6510
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00E0E0E0&
      X1              =   4500
      X2              =   1755
      Y1              =   6510
      Y2              =   6510
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00E0E0E0&
      X1              =   4500
      X2              =   1755
      Y1              =   4170
      Y2              =   4170
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00E0E0E0&
      X1              =   7335
      X2              =   7335
      Y1              =   3720
      Y2              =   2535
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00E0E0E0&
      X1              =   7335
      X2              =   7335
      Y1              =   5100
      Y2              =   3810
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00E0E0E0&
      X1              =   7335
      X2              =   7335
      Y1              =   6495
      Y2              =   5220
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00E0E0E0&
      X1              =   10140
      X2              =   10140
      Y1              =   3555
      Y2              =   2520
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00E0E0E0&
      X1              =   10140
      X2              =   10140
      Y1              =   6510
      Y2              =   5050
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00E0E0E0&
      X1              =   4500
      X2              =   4500
      Y1              =   6495
      Y2              =   4260
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00E0E0E0&
      X1              =   4500
      X2              =   4500
      Y1              =   4155
      Y2              =   2505
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "% Timestretch"
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
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   69
      Top             =   5850
      Width           =   1695
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Semitones transpose"
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
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   70
      Top             =   5470
      Width           =   1665
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   -180
      X2              =   10985
      Y1              =   6615
      Y2              =   6615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   -270
      X2              =   10985
      Y1              =   6630
      Y2              =   6630
   End
   Begin VB.Label txtLev 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-6.0 dB"
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
      Height          =   240
      Index           =   4
      Left            =   6720
      TabIndex        =   68
      Top             =   5565
      Width           =   615
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz centre frequency"
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
      Height          =   285
      Left            =   5490
      TabIndex        =   67
      Top             =   5895
      Width           =   1755
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
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
      Height          =   255
      Left            =   6735
      TabIndex        =   66
      Top             =   6225
      Width           =   600
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   1
      Left            =   3885
      TabIndex        =   65
      Top             =   4890
      Width           =   285
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   0
      Left            =   2685
      TabIndex        =   64
      Top             =   4890
      Width           =   255
   End
   Begin VB.Label txtDist 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0 dB"
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
      Height          =   240
      Left            =   9540
      TabIndex        =   63
      Top             =   2880
      Width           =   615
   End
   Begin VB.Image imgEnvRange 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   9705
      Picture         =   "DRUMSYN.frx":9E8C
      Top             =   2190
      Width           =   435
   End
   Begin VB.Image imgEnvSel 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1755
      Picture         =   "DRUMSYN.frx":A00E
      Top             =   2190
      Width           =   435
   End
   Begin VB.Image warning 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   8880
      Picture         =   "DRUMSYN.frx":A190
      Top             =   6690
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
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
      Height          =   240
      Left            =   6735
      TabIndex        =   60
      Top             =   4815
      Width           =   600
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz centre frequency"
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
      Height          =   285
      Left            =   5490
      TabIndex        =   59
      Top             =   4485
      Width           =   1755
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz"
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
      Height          =   285
      Left            =   3870
      TabIndex        =   58
      Top             =   5265
      Width           =   495
   End
   Begin VB.Label label94 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz"
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
      Height          =   285
      Left            =   2655
      TabIndex        =   57
      Top             =   5265
      Width           =   360
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "White"
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
      Height          =   240
      Left            =   6735
      TabIndex        =   56
      Top             =   3165
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Degrees phase at start "
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
      Height          =   285
      Left            =   2655
      TabIndex        =   55
      Top             =   3855
      Width           =   1710
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz"
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
      Height          =   285
      Left            =   3870
      TabIndex        =   54
      Top             =   3195
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz to"
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
      Height          =   285
      Left            =   2595
      TabIndex        =   53
      Top             =   3195
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Linear"
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
      Height          =   285
      Left            =   3870
      TabIndex        =   52
      Top             =   3510
      Width           =   630
   End
   Begin VB.Label txtLev 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-6.0 dB"
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
      Height          =   240
      Index           =   3
      Left            =   6720
      TabIndex        =   51
      Top             =   4155
      Width           =   615
   End
   Begin VB.Label txtLev 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-6.0 dB"
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
      Height          =   240
      Index           =   2
      Left            =   3870
      TabIndex        =   50
      Top             =   4575
      Width           =   615
   End
   Begin VB.Label txtLev 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-6.0 dB"
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
      Height          =   240
      Index           =   1
      Left            =   6720
      TabIndex        =   49
      Top             =   2880
      Width           =   645
   End
   Begin VB.Label txtLev 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-6.0 dB"
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
      Height          =   240
      Index           =   0
      Left            =   3870
      TabIndex        =   48
      Top             =   2850
      Width           =   630
   End
   Begin VB.Label envlab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tone"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2235
      TabIndex        =   44
      Top             =   2220
      Width           =   5000
   End
   Begin VB.Label txtEnvRange 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "200 ms"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8400
      TabIndex        =   47
      Top             =   2220
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00BEC0CA&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1650
      Index           =   5
      Left            =   1755
      Top             =   2535
      Width           =   2760
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00BEC8BA&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1200
      Index           =   0
      Left            =   4605
      Top             =   2535
      Width           =   2745
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CDC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   2265
      Index           =   4
      Left            =   1755
      Top             =   4260
      Width           =   2760
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C4B8C4&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1305
      Index           =   1
      Left            =   4605
      Top             =   3825
      Width           =   2745
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C4B8C4&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1290
      Index           =   2
      Left            =   4605
      Top             =   5235
      Width           =   2745
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0D2D2&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1050
      Index           =   3
      Left            =   7425
      Top             =   2535
      Width           =   2730
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1470
      Index           =   6
      Left            =   7425
      Top             =   5055
      Width           =   2730
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1305
      Index           =   7
      Left            =   7425
      Top             =   3660
      Width           =   2730
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuMakeWAVs 
         Caption         =   "Generate All WAVs"
      End
      Begin VB.Menu mnuDeleteWAVs 
         Caption         =   "Delete All WAVs"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete File"
      End
      Begin VB.Menu mnuX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewKit 
         Caption         =   "New Kit..."
      End
      Begin VB.Menu mnuLocation 
         Caption         =   "Set Location..."
      End
      Begin VB.Menu mnuX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuX3 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuWave 
      Caption         =   "Wave"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuX4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuX5 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dragmin, dragmax, dragpnt   'envelope point dragging
Dim EnvRangeDir, EnvRangeIndex  'for changing EnvRange
Dim EnvSelDir, EnvSel           'for changing CurEnv
Const Fs = 44100                'sampling rate

Dim lockout                     'to stop dragging overflowing

Dim CrappyVBisSettledDown

Dim but1down                    'to allow right-click to open or show menu

Dim formX, FormY

Private Sub butHelp_Click()
  rtn = WinHelp(mainform.hWnd, App.Path & "\drumsyn.hlp", &H3, 0)
  'openfile "e:\drumsyn\farfisa\fisa kick.ds"
End Sub

Private Sub butLoad_Click()
  Me.Caption = "Loading...": DoEvents
  
  If fileList.ListIndex > -1 Then
    ff = DS_PATH & "\"
    ff = ff & kitList.List(kitList.ListIndex) & "\"
    ff = ff & fileList.List(fileList.ListIndex) & ".ds"
    Me.Caption = "Loading " & ff & "...": DoEvents
    
    openfile (ff)
    changed (-1)
    Call drawAllEnvs
  End If

End Sub

Private Sub butPlay_Click()
  
  On Error Resume Next
  
  mainform.MousePointer = 13
  warning.Visible = False
  butPlay.Enabled = False
  If Not playable Then
    DoEvents
    rtn = PlaySound(filename_mute, ByVal 0&, SND_ASYNC Or SND_NODEFAULT)
'Debug.Print rtn
    cap = mainform.Caption
    savefile (filename_ds)
    rtn = ds2wav(filename_ds, filename, mainform.hWnd)
    If rtn = 0 Then playable = True
    mainform.Caption = cap
  End If

  DoEvents
  If playable Then
    rtn = PlaySound(filename, ByVal 0&, SND_ASYNC Or SND_NODEFAULT)
    If rtn = 0 Then warning.Visible = True
  Else
    warning.Visible = True
  End If
  
  'butPlay.SetFocus
  butPlay.Enabled = True
  mainform.MousePointer = 0
'Debug.Print
  
End Sub

Private Sub butSave_Click()
  
  On Error GoTo bserr

  Load fileform
  fn = fileList.List(fileList.ListIndex)
  If Right(fn, 1) = " " Then fn = Left(fn, Len(fn) - 1)
  fileform.txt = fn
  
  fileform.Show 1
  
  mainform.MousePointer = 11
  
  If AskFileName <> "" Then
    ff = DS_PATH & "\" & kitList.List(kitList.ListIndex) & "\"
    
    If Dir(ff & AskFileName & ".wav") <> "" And Dir(ff & AskFileName & ".ds") = "" Then
      ret = MsgBox("Overwrite Existing .WAV File?", 33, AskFileName)
      If ret = 1 Then GoTo JUMP 'overwrite WAV without DS
    Else
JUMP:
      savefile (ff & AskFileName & ".ds")
      If playable Then
        FileCopy filename, ff & AskFileName & ".wav"
      Else
        rtn = ds2wav(ff & AskFileName & ".ds", ff & AskFileName & ".wav", 0)
      End If
    End If

  End If
  Call fillfilelist
  mainform.MousePointer = 0

Exit Sub

bserr:
  'Debug.Print bserr!
  warning.Visible = True
  Resume Next

End Sub

Private Sub changed(n)
  If n < 0 Then GoTo skipcheck
  If chkOn(n) Then
skipcheck:
    warning.Visible = False
    playable = False
  End If
End Sub

Private Sub chkDist_Click()
  changed (4)
End Sub

Private Sub chkFixedRandom_Click()
  changed (1)
End Sub

Private Sub chkHPF_Click()
  changed (-1)
End Sub

Private Sub chkMasterOn_Click()
  changed (-1)
End Sub

Private Sub chkOFilter_Click()
  changed (-1)
End Sub

Private Sub chkOn_Click(Index As Integer)
  changed (-1)
  For i = 0 To 4
    tot = tot + chkOn(i)
  Next i
  If tot = 0 Then
    butPlay.Enabled = False
  Else
    butPlay.Enabled = True
  End If
  If Index = 0 Then
    chkOTrack1.Enabled = chkOn(Index)
    chkOTrack2.Enabled = chkOn(Index)
  End If
End Sub

Private Sub chkOn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ndex = Index
  If ndex < 5 Then
    If ndex = 2 And CurEnv = 3 Then ndex = 6
    If chkOn(Index) Or Button > 1 Then EnvSelChange (ndex + 100)
  End If
End Sub

Private Sub chkOTrack1_Click()
  changed (2)
End Sub

Private Sub chkOTrack2_Click()
  changed (2)
End Sub

Private Sub comDB_Click()
  changed (5)
End Sub

Private Sub comDF_Click()
  changed (5)
End Sub

Private Sub comment_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub comOMethod_Click()
  changed (2)
  
  Select Case comOMethod.ListIndex
    Case 0:    txtOParam = "Mix"
    Case 1, 2: txtOParam = "Depth"
    Case 3:    txtOParam = "Reson."
  End Select
  
  tmp = True
  If comOMethod.ListIndex = 3 Then tmp = False
  chkOTrack1.Enabled = tmp
  chkOTrack2.Enabled = tmp
  comOW1.Enabled = tmp
  comOW2.Enabled = tmp
  
  Call ShowEnvName(EnvSel)
  Call drawAllEnvs
  Call DrawEnv(CurEnv)
End Sub

Private Sub comOW1_Click()
  changed (2)
End Sub

Private Sub comOW2_Click()
  changed (2)
End Sub

Private Sub envbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  ew = envbox.ScaleWidth / 100
  dragpnt = -1
  dragging = -1
  ep = 0
  While envpts(CurEnv, 0, ep + 1) >= 0
    If Abs(X - envpts(CurEnv, 0, ep)) < ew Then   'for each point
      If Abs(Y - envpts(CurEnv, 1, ep)) < 4 Then  'if selected
        dragging = ep              'start dragging
      End If
    End If
    If X > envpts(CurEnv, 0, ep) Then epos = ep 'position for insert
    ep = ep + 1
  Wend
  
  If Abs(X - envpts(CurEnv, 0, ep)) < ew Then 'do--while instead?
    If Abs(Y - envpts(CurEnv, 1, ep)) < 3 Then
      dragging = ep
    End If
  End If

  If dragging = -1 And Button = 1 Then 'insert point
    If ep < 30 Then
      If X > envpts(CurEnv, 0, ep) Then
        envpts(CurEnv, 0, ep) = X
      End If
      For j = ep + 1 To (epos + 1) Step -1
        envpts(CurEnv, 0, j + 1) = envpts(CurEnv, 0, j)
        envpts(CurEnv, 1, j + 1) = envpts(CurEnv, 1, j)
      Next j
      envpts(CurEnv, 0, epos + 1) = X
      envpts(CurEnv, 1, epos + 1) = Y
      Call showEnvPos(X, Y * 327.67)
      Call DrawEnv(CurEnv)
      dragging = epos + 1
      ep = ep + 1
    Else
      Beep
    End If
    Call DrawEnv(CurEnv)

  ElseIf Button = 2 Then PopupMenu mnuWave, 2
  
  End If
  
  If dragging >= 0 Then 'drag        '****need to drag on mouse-down!!!
    If dragging = 0 Then
      dragmin = 0
      dragmax = 0
    ElseIf dragging = ep Then
      dragmin = envpts(CurEnv, 0, ep - 1)
      dragmax = envbox.ScaleWidth - 5 '10 sec?
    Else
      dragmin = envpts(CurEnv, 0, dragging - 1)
      dragmax = envpts(CurEnv, 0, dragging + 1)
    End If
    dragpnt = dragging
  End If
  
End Sub

Private Sub envbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If lockout Then
    'ignore, system too busy!
  Else
    lockout = True
      If dragpnt > -1 Then
        If Y < -16 Or Y > 120 Then 'delete point
          If dragpnt > 0 And envpts(CurEnv, 0, dragpnt + 1) >= 0 Then
            j = dragpnt
            While (envpts(CurEnv, 0, j)) >= 0
              envpts(CurEnv, 0, j) = envpts(CurEnv, 0, j + 1)
              envpts(CurEnv, 1, j) = envpts(CurEnv, 1, j + 1)
              j = j + 1
            Wend
            dragpnt = -1
            Call DrawEnv(CurEnv)
          End If
        Else
          If (Y < 0) Or (envpts(CurEnv, 0, dragpnt + 1) < 0) Then Y = 0
          If Y > 100 Then Y = 100
          envpts(CurEnv, 1, dragpnt) = Y
          If X < dragmin Then X = dragmin
          If X > dragmax Then X = dragmax
          envpts(CurEnv, 0, dragpnt) = X
          Call DrawEnv(CurEnv)
          Call showEnvPos(X, Y * 327.67)
        End If
        changed (-1)
      ElseIf Button = 2 Then
        Call showEnvPos(X, Y * 327.67)
      End If
      lockout = False
  End If

End Sub

Private Sub envbox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  dragpnt = -1
  DoEvents
  Call ShowEnvName(CurEnv)
  envbox.MousePointer = 0
  changed (-1)
  
End Sub

Private Sub envbox_Paint() 'needed if no autoredraw
  If CrappyVBisSettledDown Then Call DrawEnv(CurEnv)
End Sub

Private Sub EnvRangeChange(ERDir)
  
  EnvRangeIndex = EnvRangeIndex + ERDir
  If EnvRangeIndex < 0 Then EnvRangeIndex = 0
  If EnvRangeIndex > 12 Then EnvRangeIndex = 12

  Select Case EnvRangeIndex
    Case 12: txtEnvRange = "16 ms": EnvRange = Fs * 0.016
    Case 11: txtEnvRange = "25 ms": EnvRange = Fs * 0.025
    Case 10: txtEnvRange = "40 ms": EnvRange = Fs * 0.04
    Case 9: txtEnvRange = "63 ms": EnvRange = Fs * 0.063
    Case 8: txtEnvRange = "100 ms": EnvRange = Fs * 0.1
    Case 7: txtEnvRange = "160 ms": EnvRange = Fs * 0.16
    Case 6: txtEnvRange = "250 ms": EnvRange = Fs * 0.25
    Case 5: txtEnvRange = "400 ms": EnvRange = Fs * 0.4
    Case 4: txtEnvRange = "630 ms": EnvRange = Fs * 0.63
    Case 3: txtEnvRange = "1 sec": EnvRange = Fs * 1#
    Case 2: txtEnvRange = "1.6 sec": EnvRange = Fs * 1.6
    Case 1: txtEnvRange = "3 sec": EnvRange = Fs * 3
    Case 0: txtEnvRange = "10 sec": EnvRange = Fs * 10
  End Select
  envbox.ScaleWidth = EnvRange
  Call drawAllEnvs
  Call DrawEnv(CurEnv)

End Sub

Private Sub EnvSelChange(ESDir)

  If ESDir > 99 Then 'direct selection
    Select Case ESDir
      Case 100: EnvSel = 1
      Case 101: EnvSel = 2
      Case 102: EnvSel = 3
      Case 103: EnvSel = 5
      Case 104: EnvSel = 6
      Case 105: EnvSel = 7
      Case 106: EnvSel = 4
    End Select
  Else
    EnvSel = EnvSel + ESDir
    If EnvSel < 1 Then EnvSel = 7
    If EnvSel > 7 Then EnvSel = 1
  End If
  
  CurEnv = EnvSel
  Select Case CurEnv
    Case 1: chkOn(0).SetFocus
    Case 2: chkOn(1).SetFocus
    Case 3: chkOn(2).SetFocus
    Case 4: chkOn(2).SetFocus
    Case 5: chkOn(3).SetFocus
    Case 6: chkOn(4).SetFocus
    Case 7: MasterFilter.SetFocus
  End Select
 
  Call ShowEnvName(EnvSel)
  Call drawAllEnvs
  Call DrawEnv(CurEnv)
  DoEvents

End Sub

Private Sub fileList_DblClick()
  Call fileList_MouseDown(1, 0, 0, 0)
End Sub

Private Sub fileList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim ff As String

  If Button = 1 Or Button = 99 Then
    'trigger
    ff = fileList.List(fileList.ListIndex)
  
    If Right(ff, 1) = " " Then 'not .ds file
      butLoad.Enabled = False
      ff = Left(ff, Len(ff) - 1)
    Else
      butLoad.Enabled = True
    End If
    
    ff = "\" & ff & ".wav"
    ff = kitList.List(kitList.ListIndex) & ff
    ff = DS_PATH & "\" & ff
    rtn = PlaySound(ff, ByVal 0&, SND_ASYNC Or SND_NODEFAULT)
      
    'remember mouse button is down
    but1down = True
  Else
    If but1down = True Then
      butLoad_Click
      but1down = False
    Else
      mnuDeleteWAVs.Caption = "Delete WAVs (" & TotalSize() & ")"
      PopupMenu mnuMain, 2
    End If
  End If
    
End Sub

Private Sub fileList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  but1down = False
End Sub

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)

  but1down = False
  
  On Error GoTo KDerr
  
  If keycode = 13 Or (keycode = 17 And Shift = 2) Then
    If butPlay.Enabled Then butPlay_Click
  End If
  
  If keycode > 111 And keycode < 124 Then
     fileList.ListIndex = -1
     fileList.ListIndex = keycode - 112 + (Shift * 12)
     Call fileList_MouseDown(1, 0, 0, 0)
     but1down = False
     keycode = 0
  End If

Exit Sub

KDerr:
  Resume Next

End Sub

Private Sub Form_Load()
   
  On Error GoTo FLerr

  Dim tmps As String

  Select Case Int(Timer * 97) Mod 6
    Case 0: comment.Text = "Right-click the file list and envelope display for menu options"
    Case 1: comment.Text = "You can use this space to add a comment to each sound!"
    Case 2: comment.Text = "DrumSynth 2.0 by maxim digital audio, Copyright ©2000 Paul Kellett  http://www.mda-vst.com"
    Case 3: comment.Text = "Use FruityLoops to sequence your DrumSynth sounds!  http://www.fruityloops.com"
    Case Else: comment.Text = "You can use [Ctrl] and [Enter] to trigger the Play button!"
  End Select

  INI_FILENAME = App.Path
  If Right(INI_FILENAME, 1) <> "\" Then INI_FILENAME = INI_FILENAME & "\"
  INI_FILENAME = INI_FILENAME & "DrumSyn.ini"
  Xtmp = getini("Window", "X", "-1")
  Ytmp = getini("Window", "Y", "-1")
  DS_PATH = getini("Files", "Path", App.Path)
  SaveWAV = Val(getini("Files", "SaveWAVs", "1"))
  If Right(DS_PATH, 1) = "\" Then DS_PATH = Left(DS_PATH, Len(DS_PATH) - 1)
  If Xtmp > 0 Then mainform.Left = Xtmp Mod Screen.Width
  If Ytmp > 0 Then mainform.Top = Ytmp Mod Screen.Height
  
  
  formX = mainform.Width: FormY = mainform.Height
  mainform.Show
'  MsgBox "Window should be visible!"

  If ScreenColors() = 256 Or ScreenColours = 16 Then 'not sure this works any more...
    For i = 0 To 7
      Shape1(i).BackColor = &HC0C0C0
    Next i
    For i = 0 To 5
      chkOn(i).BackColor = &HC0C0C0
    Next i
    chkMasterOn.BackColor = &HC0C0C0
    chkHPF.BackColor = &HC0C0C0
    chkFixedRandom.BackColor = &HC0C0C0
    chkOTrack1.BackColor = &HC0C0C0
    chkOTrack2.BackColor = &HC0C0C0
    chkOFilter.BackColor = &HC0C0C0
    MasterFilter.BackColor = &HC0C0C0
    For i = 0 To 7
      envcol(i) = &HC0C0C0
    Next i
  Else
    envcol(1) = &HB0C0FF
    envcol(2) = &HC0FFC0
    envcol(3) = &HFFC0C0
    envcol(4) = &HFFC0C0
    envcol(5) = &HFFC0FF
    envcol(6) = &HFFC0FF
    envcol(7) = &HF0F0C0
  End If
  
  Call Setup
  Call ReadDirs
  Call GetWorkingPath
  Call Synth
  
  If Command <> "" Then
    If Left(Command, 1) = Chr(34) Then 'quotes
      openfile Mid(Command, 2, Len(Command) - 2)
    Else
      openfile tmps
    End If
  End If

Exit Sub

FLerr:
  'Debug.Print FLerr!
  'If errCount < 4 Then
  '  MsgBox errorText, 48
  '  errorCount = errorCount + 1
  'End If
  Resume Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error Resume Next

  INI_FILENAME = App.Path
  If Right(INI_FILENAME, 1) <> "\" Then INI_FILENAME = INI_FILENAME & "\"
  INI_FILENAME = INI_FILENAME & "DrumSyn.ini"
  
  setini "Window", "X", mainform.Left
  setini "Window", "Y", mainform.Top
  setini "Files", "Path", (DS_PATH)
  setini "Files", "SaveWAVs", (SaveWAV)
  
  Kill filename
  Kill filename_ds
  Kill filename_mute
  tmp = App.Path: If Right(tmp, 1) <> "\" Then tmp = tmp & "\"
  SetAttr tmp & "DrumSyn.GID", 0
  Kill tmp & "DrumSyn.GID"
 
End Sub


Private Sub Form_Resize()

 ' If mainform.WindowState = 0 Then 'normal
 '   mainform.Width = formX: mainform.Height = FormY
 ' End If

End Sub



Private Sub imgEnvRange_DblClick()
  Call EnvRangeChange(EnvRangeDir)
End Sub

Private Sub imgEnvRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  EnvRangeDir = 1 - 2 * Int((X / 200))
  Call EnvRangeChange(EnvRangeDir)
 
End Sub

Private Sub imgEnvSel_DblClick()
  Call EnvSelChange(EnvSelDir)
End Sub

Private Sub imgEnvSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  EnvSelDir = 2 * Int((X / 200)) - 1
  Call EnvSelChange(EnvSelDir)

End Sub

Private Sub kitList_Click()

  On Error GoTo klcerr

  If CrappyVBisSettledDown Then
    'Else
      Call fillfilelist
    'End If
  End If

Exit Sub

klcerr:
  'Debug.Print bserr!
  Resume Next

End Sub

Private Function Log10(X)
  If X <= 0 Then
    Log10 = -999
  Else
    Log10 = Log(X) / 2.3025851
  End If
End Function

Private Sub MasterFilter_Click()
  changed (-1)
  If MasterFilter Then EnvSelChange (105)
End Sub

Private Sub MasterFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button > 1 Then EnvSelChange 100 + 5
End Sub

Private Sub MasterLength_Change()
  changed (-1)
End Sub

Private Sub MasterTune_Change()
  changed (-1)
End Sub

Private Sub ReadDirs()

  On Error Resume Next

  'errText ("ReadDirs()")
  CrappyVBisSettledDown = False

  kitList.Clear
  rd = Dir(DS_PATH & "\*.*", 16)
  While rd <> ""
    If rd <> "." And rd <> ".." Then
      If GetAttr(DS_PATH & "\" & rd) And 16 Then
        Mid(rd, 1, 1) = UCase(Mid(rd, 1, 1))
        ucs = 0
        For uci = 1 To Len(rd)
          If Asc(Mid(rd, uci, 1)) < 65 And Asc(Mid(rd, uci, 1)) <> 32 Then ucs = 1
        Next uci
        If ucs = 1 Then rd = UCase(rd)
        kitList.AddItem rd
        kitList.ListIndex = 0
        'fPath = kitList.List(0)
      End If
    End If
    rd = Dir
  Wend
  'KitList.AddItem "[new]"

  'errText ("ReadDirs() Finished")

  CrappyVBisSettledDown = True
  'errText ("Calling FillFileList()")
  Call fillfilelist

End Sub

Private Sub Setup()
  
  'errText ("Setup()")
  
  'initialise envelopes
  For ce = 1 To 6
    If ce = 2 Then
      cee = 6
    ElseIf ce = 6 Then
      cee = 2
    Else
      cee = ce
    End If
    envpts(cee, 0, 0) = 0: envpts(cee, 1, 0) = 100
    envpts(cee, 0, 1) = (1 - ce / 8) * 2000: envpts(cee, 1, 1) = 20
    envpts(cee, 0, 2) = (1 - ce / 8) * 6000: envpts(cee, 1, 2) = 0
    envpts(cee, 0, 3) = -1
  Next ce

  envpts(7, 0, 0) = 0: envpts(7, 1, 0) = 100 'filter env
  envpts(7, 0, 1) = 444000: envpts(7, 1, 1) = 100
  envpts(7, 0, 2) = 444000: envpts(7, 1, 2) = 0
  envpts(7, 0, 3) = -1

  CurEnv = 1
  dragpnt = -1
  EnvRangeIndex = 6
  EnvSel = 1
  'Call drawAllEnvs

  'Fill combo boxes
  comOW1.AddItem "Sine": comOW2.AddItem "Sine"
  comOW1.AddItem "Sine²": comOW2.AddItem "Sine²"
  comOW1.AddItem "Tri": comOW2.AddItem "Tri"
  comOW1.AddItem "Saw": comOW2.AddItem "Saw"
  comOW1.AddItem "Square": comOW2.AddItem "Square"
  comOW1.ListIndex = 0: comOW2.ListIndex = 0

  comOMethod.AddItem "A + B"
  comOMethod.AddItem "A . B (freq modulation)"
  comOMethod.AddItem "A x B (ring modulation)"
  comOMethod.AddItem "A....B (cymbal tones)"
  comOMethod.ListIndex = 0
  
  comDB.AddItem "[none]"
  comDB.AddItem "14 bit"
  comDB.AddItem "12 bit"
  comDB.AddItem "10 bit"
  comDB.AddItem " 8 bit"
  comDB.AddItem " 6 bit"
  comDB.AddItem " 4 bit"
  comDB.AddItem " 2 bit"
  comDB.ListIndex = 0

  comDF.AddItem "[none]"
  comDF.AddItem "11 kHz"
  comDF.AddItem " 7 kHz"
  comDF.AddItem " 5 kHz"
  comDF.AddItem " 4 kHz"
  comDF.AddItem " 3 kHz"
  comDF.AddItem " 2 kHz"
  comDF.AddItem " 1 kHz"
  comDF.ListIndex = 0

End Sub

Private Sub ShowEnvName(CEnv)
  
  Select Case CEnv
     Case 1: envlab = "Tone"
     Case 2: envlab = "Noise"
     Case 3:
       If comOMethod.ListIndex = 0 Then
         envlab = "Overtone A"
       Else
         envlab = "Overtone level"
       End If
     Case 4:
       If comOMethod.ListIndex = 0 Then
         envlab = "Overtone B"
       Else
         envlab = "Overtone modulation depth"
         If comOMethod.ListIndex = 3 Then envlab = "Overtone frequency range"
       End If
     Case 5: envlab = "Noise Band 1"
     Case 6: envlab = "Noise Band 2"
     Case 7: envlab = "Filter frequency"
  End Select

End Sub

Private Sub showEnvPos(X, Y) 'show envelope position in ms/dB
  
  If X >= 0 And X < envbox.ScaleWidth Then
    If Y >= -5 And Y <= 32768 Then
      If CurEnv = 4 Or CurEnv = 7 Then
        envlab = toTime(X) & ": " & Format(0.00305 * Y, "0") & " %"
      Else
        envlab = toTime(X) & ": " & todB(Y)
      End If
    Else
      envlab = ""
    End If
  Else
    envlab = ""
  End If

End Sub

Private Sub mnuLocation_Click()
  
  On Error GoTo newlocerr

  Load locform
  locform.Show 1
    
  Call ReadDirs
  
Exit Sub
                     
newlocerr:
  'Debug.Print newlocerr!
  mainform.warning.Visible = True
  Resume Next

End Sub

Private Sub sliDist_Change()
  txtDist = Format(-(sliDist), "0") & " dB"
  changed (5)
End Sub

Private Sub sliFres_Change()
  txtFres = "Q = " & sliFres
  changed (-1)
End Sub

Private Sub sliLev_Change(Index As Integer)
  txtLev(Index).Caption = todB(sliLev(Index) ^ 2)
  changed (Index)
End Sub

Private Sub sliNBD_Change()
'  label12 = sliNBD
'  changed (3)
End Sub

Private Sub sliMasterdB_Change()
  If sliMasterdB > 0 Then
    txtMasterdB = "+" & sliMasterdB & " dB"
  Else
    txtMasterdB = sliMasterdB & " dB"
  End If
  changed (-1)
End Sub

Private Sub sliNdF_Change()
  changed (3)
End Sub

Private Sub sliNdF2_Change()
  changed (4)
End Sub

Private Sub sliNSlope_Change()
  changed (1)
  Select Case sliNSlope
    Case Is < -80: Label4 = "Red"
    Case Is < -20: Label4 = "Pink"
     Case Is < 20: Label4 = "White"
     Case Is < 80: Label4 = "Azure"
        Case Else: Label4 = "Blue"
  End Select
 
End Sub

Private Sub sliOBright_Change()
  changed (2)
End Sub

Private Sub sliOParam_Change()
  changed (2)
End Sub

Private Sub sliOVari_Change()
  changed (2)
End Sub

Private Sub sliTDroop_Change()
  
  Select Case sliTDroop
    Case Is < 2: Label1 = "Linear"
    Case Is < 25: Label1 = "Slow"
    Case Is < 66: Label1 = "Exp"
    Case Else: Label1 = "Fast"
  End Select
  
  changed (0)

End Sub

Private Function todB(a) '32767ths -> dB
  
  If a < 1 Then
    todB = "Mute"
  Else
    todB = Format(20 * Log10(a / 32767), "0.0") & " dB"
  End If

End Function

Private Function TotalSize() 'size of .WAV files in current dir

  On Error GoTo totsizerr

  totsiz = 0
  pp = DS_PATH & "\" & mainform.kitList.List(mainform.kitList.ListIndex)
  fl = Dir(pp & "\*.wav")
  While fl <> ""
    totsiz = totsiz + FileLen(pp & "\" & fl) / 1024
    fl = Dir
  Wend
  TotalSize = Format(totsiz, "0") & "k"

Exit Function

totsizerr:
  'Debug.Print totsizerr!
  Resume Next

End Function

Private Function toTime(t) 'samples -> ms
  
  Select Case t
    Case 0: toTime = "0 ms"
    Case Is < (Fs / 100): toTime = Format(1000 * t / Fs, "0.000") & " ms"
    Case Is < (Fs / 10): toTime = Format(1000 * t / Fs, "0.0") & " ms"
    Case Else: toTime = Format(1000 * t / Fs, "0") & " ms"
  End Select

End Function

Private Sub txtNF_Change()
  changed (3)
End Sub

Private Sub txtNF2_Change()
  changed (4)

End Sub

Private Sub txtOF1_Change()
  changed (2)
End Sub

Private Sub txtOF2_Change()
  changed (2)
End Sub

Private Sub txtTF1_Change()
  changed (0)
End Sub

Private Sub txtTF2_Change()
  changed (0)
End Sub

Private Sub txtTPhase_Change()
  changed (0)
End Sub


Private Sub mnuCopy_Click() 'copy current envelope
  
  mnuPaste.Enabled = True
  For i = 0 To 18
    envbuf(0, i) = envpts(CurEnv, 0, i)
    envbuf(1, i) = envpts(CurEnv, 1, i)
  Next i
  
End Sub

Private Sub mnuDelete_Click() 'delete one .ds & .wav sound
  
  On Error GoTo mnudelerr

  LI = mainform.fileList.ListIndex
  
  If LI >= 0 Then
    
    fn = mainform.kitList.List(mainform.kitList.ListIndex)
    fn = fn & "\" & mainform.fileList.List(LI)
  
    If MsgBox("Delete this sound?", 49, fn) = 1 Then
      
      fn = DS_PATH & "\"
      fn = fn & mainform.kitList.List(mainform.kitList.ListIndex)
      fn = fn & "\" & mainform.fileList.List(LI)
      
      If Right(fn, 1) = " " Then
        Kill Left(fn, Len(fn) - 1) & ".wav"
      Else
        Kill fn & ".ds" 'kill
        Kill fn & ".wav" 'kill
      End If
      
      fillfilelist 'refresh list

    End If
  End If

Exit Sub

mnudelerr:
  'Debug.Print mnudelerr!
  mainform.warning.Visible = True
  Resume Next

End Sub

Private Sub mnuDeleteWAVs_Click()
  
  On Error GoTo mnudelwaverr

 ' If MsgBox("Delete all the WAV files for this kit?", 33) = 1 Then
    
    fn = DS_PATH & "\"
    fn = fn & mainform.kitList.List(mainform.kitList.ListIndex) & "\"
    
    For i = 0 To mainform.fileList.ListCount - 1
      ent = mainform.fileList.List(i)
      If Right(ent, 1) <> " " Then
        Kill fn & ent & ".wav"
      End If
    Next i

 ' End If

Exit Sub

mnudelwaverr:
  'Debug.Print mnudelwaverr!
  'mainform.warning.Visible = True
  Resume Next

End Sub

Private Sub mnuMakeWAVs_Click() 'generate all .wavs

  kit = mainform.kitList.List(mainform.kitList.ListIndex)
  
  'If MsgBox("Generate all the WAV files for this kit?", 33, kit) = 1 Then
    genform.Show 1
  'End If

End Sub

Private Sub mnuNewKit_Click()
    
  On Error GoTo newkiterr

  Load fileform
  fileform.txt = ""
  fileform.Caption = "New Kit"
  fileform.chkSaveWAV.Visible = False
  fileform.Show 1
    
  If AskFileName = "" Then
    'nothing
  Else
    MkDir DS_PATH & "\" & AskFileName
    mainform.kitList.AddItem UCase(AskFileName), mainform.kitList.ListCount
    mainform.kitList.ListIndex = mainform.kitList.ListCount - 1
  End If
  
Exit Sub
                     
newkiterr:
  mainform.warning.Visible = True
  Resume Next

End Sub

Private Sub mnuPaste_Click()
 
  For i = 0 To 18
    envpts(CurEnv, 0, i) = envbuf(0, i)
    envpts(CurEnv, 1, i) = envbuf(1, i)
  Next i

  changed (-1)
  Call drawAllEnvs

End Sub

Private Sub mnuReset_Click()

  If CurEnv = 7 Then
    envpts(7, 0, 0) = 0: envpts(7, 1, 0) = 100
    envpts(7, 0, 1) = 442000: envpts(7, 1, 1) = 100
    envpts(7, 0, 2) = 442000: envpts(7, 1, 2) = 0
    envpts(7, 0, 3) = -1
  Else
    envpts(CurEnv, 0, 0) = 0: envpts(CurEnv, 1, 0) = 100
    envpts(CurEnv, 0, 1) = (1 - CurEnv / 8) * 6000: envpts(CurEnv, 1, 1) = 30
    envpts(CurEnv, 0, 2) = (1 - CurEnv / 8) * 12000: envpts(CurEnv, 1, 2) = 0
    envpts(CurEnv, 0, 3) = -1
  End If
   
  Call drawAllEnvs

End Sub

