VERSION 5.00
Begin VB.Form FrmWaterEffect 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Water Effect"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "FrmWaterEffect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdDIBTest 
      Caption         =   "&DIB Test"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox picWaterEffect 
      Height          =   4575
      Left            =   120
      Picture         =   "FrmWaterEffect.frx":0ECA
      ScaleHeight     =   300
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   500
      TabIndex        =   6
      Top             =   120
      Width           =   7575
   End
   Begin VB.Timer tmrWaterEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   4800
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOriginalpicture 
      Caption         =   "&Original Picture"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdBoats 
      Caption         =   "&Boats"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRaindrop 
      Caption         =   "&Raindrop"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox picOriginal 
      Height          =   4575
      Left            =   120
      Picture         =   "FrmWaterEffect.frx":58FC
      ScaleHeight     =   300
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label lblCopyright 
      Caption         =   $"FrmWaterEffect.frx":A32E
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   7575
   End
End
Attribute VB_Name = "FrmWaterEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' __        __    _              _____  __  __          __
' \ \      / /_ _| |_ ___ _ __  | ____|/ _|/ _| ___  ___| |_
'  \ \ /\ / / _` | __/ _ \ '__| |  _| | |_| |_ / _ \/ __| __|
'   \ V  V / (_| | ||  __/ |    | |___|  _|  _|  __/ (__| |_
'    \_/\_/ \__,_|\__\___|_|    |_____|_| |_|  \___|\___|\__|
'
'
' Form:                 FrmWaterEffect
' Copyright:            Code/gfx by Reiner Rottmann (mail@Reiner-Rottmann.de)
'                       Original Pascal Code by Roy Willemse (r.willemse@dynamind.nl)
' Creation Date:        01/08/2002
' Changes:
' 01/08/2002            Reiner Rottmann     Initial Version
' 10/08/2002            Asmodi              Added DIB rendering engine
' 14/08/2002            Reiner Rottmann     New sample form created
'                                           New sample picture included
' 15/08/2002            Reiner Rottmann     English translation of the commentary
' 16/08/2002            Reiner Rottmann     Raindrop now starts in the middle
'---------------------------------------------------------------------------------------
' Todo List:
' [ ] Clean up sourcecode
' [ ] Enable Water Effect rendering on mousemove event
'
Option Explicit

Private Sub cmdQuit_Click()

    subUnload
    End

End Sub

' A short test for the DIB engine
Private Sub cmdDIBTest_Click()

    Call subDIBTest(picWaterEffect)

End Sub

' Show original picture and disable timer
Private Sub cmdOriginalpicture_Click()

    FrmWaterEffect.Caption = "Water Effect"
    tmrWaterEffect.Enabled = False
    FrmWaterEffect.picWaterEffect.Picture = FrmWaterEffect.picOriginal.Picture

End Sub


' Simulate a random rain drop
Private Sub cmdRaindrop_Click()

    tmrWaterEffect.Enabled = True
    WaveMapDrop MaxX / 2, MaxY / 2, 10, 25
    ' RainDrop

End Sub

Private Sub Form_Activate()

  ' Some doevents assure that the program has enough time to load the picture

    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    ' Load picture into DIB
    Call subPicToDIB

End Sub

Private Sub Form_Load()

  ' A betatest message
  ' MsgBox "Die Geschwindigkeit muss noch optimiert werden. Ansonsten ist der Effekt soweit fertig.", vbInformation + vbOKOnly, "Betatest Nachricht"

    InitializeWaterEffekt

End Sub

' Makes the program ready to unload
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    subUnload
    End

End Sub

' Disturb the water on mouseclick
Private Sub picWaterEffect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'WaveMap(CT, X, Y) = -100
    'tmrWaterEffect.Enabled = True

End Sub

' Disturb the water on mousemove (deactivated due to the slow rendering)
Private Sub picWaterEffect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'WaveMap(CT, X, Y) = -100
    'tmrWaterEffect.Enabled = True

End Sub

' Renders the effect on the picturebox
Private Sub tmrWaterEffect_Timer()

    FrmWaterEffect.Caption = "Water Effect - [rendering in progress]"
    RenderWaveMapWithDIB

End Sub


