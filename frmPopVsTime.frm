VERSION 5.00
Begin VB.Form frmPopVsTime 
   Caption         =   "Monarch Population vs. Time"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10470
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   6780
      Left            =   120
      Picture         =   "frmPopVsTime.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   10230
   End
End
Attribute VB_Name = "frmPopVsTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExit_Click()
    Unload Me
End Sub

