VERSION 5.00
Begin VB.Form frmMainMenu 
   AutoRedraw      =   -1  'True
   Caption         =   "Main Menu"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   9240
      TabIndex        =   4
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9240
      TabIndex        =   2
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dillon Chan 2015"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Danaus Plexippus Simulator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   6975
      Left            =   0
      Picture         =   "frmMainMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11130
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Beep
    End
End Sub

Private Sub cmdStart_Click()
    Unload Me
    frmSimMain.Show
End Sub

Private Sub Picture1_Click()

End Sub
