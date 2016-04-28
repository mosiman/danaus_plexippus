VERSION 5.00
Begin VB.Form frmCornVsPopVsTime 
   Caption         =   "Corn vs Monarch vs Time"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   6780
      Left            =   120
      Picture         =   "frmCornVsPopVsTime.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   10230
   End
End
Attribute VB_Name = "frmCornVsPopVsTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExit_Click()
    Unload Me
End Sub
