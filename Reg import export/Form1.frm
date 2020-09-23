VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registery import export "
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Import reg file"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export reg file"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Small but usefull for beginners :)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "/is = Silent import"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "/i  = Import"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "/e = Export"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dot REG Project"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Command2.Enabled = True

'This will export the registery info from (HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run)to your app.path

Shell "regedit /e exported.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
End Sub

Private Sub Command2_Click()

'This will import the registery info to (HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run)from your app.path

Shell "regedit /i exported.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"

End Sub


'Cut here

'A small note*
'--------------
'Note that you can play with the loactions of the reg file let say instead of making in save or load from hkey current user
'You can use any other loaction,:) choose it for your own need ! Thx a lot !

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'ITS A SMALL PROJECT BUT USEFULL TO PLAY AROUND WITH OR INTO REGEDIT !
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


Private Sub Form_Load()
Command2.Enabled = False

End Sub
