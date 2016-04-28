VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Left            =   9600
      Top             =   5880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   735
      Left            =   8760
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   9360
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   600
      Y2              =   4920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YCount As Single
Dim XCount As Single
    Const X1 = 480
    Const Y1 = 4920
    
Option Explicit

Private Sub Command1_Click()
    Const X1 = 480
    Const Y1 = 4920
    
    Dim K As Integer, X As Integer
    
    tmrTime.Enabled = True
End Sub

Private Sub Form_Load()
    XCount = 1991
End Sub

Private Sub tmrTime_Timer()
    Const e = 2.71828
    If XCount > 2116 Then
        tmrTime.Enabled = False
    Else
        YCount = ((1 * 10 ^ 94) * (e ^ -0.098 * XCount)) / 1000
    End If
    
    Form1.PSet (X1 + XCount, Y2 - YCount)
    
    XCount = XCount + 1
End Sub
