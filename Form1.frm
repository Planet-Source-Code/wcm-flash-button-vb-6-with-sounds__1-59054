VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Flash Buttons with Sounds"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3000
      Picture         =   "Form1.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   360
      Width           =   480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Information"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   6735
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"Form1.frx":1594
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash s1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash s2 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash s3 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash s4 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Sub Form_Load()
 
 s1.Movie = App.Path & "\Payroll.swf"
 s1.Menu = False

 s2.Movie = App.Path & "\Deductions.swf"
 s2.Menu = False
 
 s3.Movie = App.Path & "\201.swf"
 s3.Menu = False
 
 s4.Movie = App.Path & "\Exit.swf"
 s4.Menu = False



End Sub

Private Sub s1_FSCommand(ByVal command As String, ByVal args As String)
If command = "RollOver" Then
PlaySound App.Path & "\ImpulseSwish.WAV", ByVal 0&, SND_FILENAME Or SND_ASYNC
Text1 = "This is for the Payslip Generation, Printing and Distribution."
End If

If command = "ButtonClick" Then
MsgBox "You have Clicked the Payslip Button", vbInformation, "Flash Buttons"
End If


End Sub

Private Sub s2_FSCommand(ByVal command As String, ByVal args As String)
If command = "RollOver" Then
PlaySound App.Path & "\ImpulseSwish.WAV", ByVal 0&, SND_FILENAME Or SND_ASYNC
Text1 = "Corresponding Deductions are Inputted here.(SSS, Philhealth, Pag-Ibig, TAX, etc."
End If

If command = "ButtonClick" Then
MsgBox "You have Clicked the Deductions Button", vbInformation, "Flash Buttons"
End If


End Sub

Private Sub s3_FSCommand(ByVal command As String, ByVal args As String)
If command = "RollOver" Then
PlaySound App.Path & "\ImpulseSwish.WAV", ByVal 0&, SND_FILENAME Or SND_ASYNC
Text1 = "This Option Provides Access to all of Employee Records. You may use this option to Add new Records, Update and Delete."

End If

If command = "ButtonClick" Then
MsgBox "You have Clicked the 201 File Button", vbInformation, "Flash Button"
End If


End Sub

Private Sub s4_FSCommand(ByVal command As String, ByVal args As String)
If command = "RollOver" Then
PlaySound App.Path & "\ImpulseSwish.WAV", ByVal 0&, SND_FILENAME Or SND_ASYNC
Text1 = "Log-Out from the System Application"
End If

If command = "ButtonClick" Then
MsgBox "You Click the Exit Button", vbExclamation, "Flash Buttons"
End
End If


End Sub

