VERSION 5.00
Begin VB.Form frmSummer 
   Caption         =   "Summer"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<----"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "---->"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Summer"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   2
      Left            =   1320
      Picture         =   "frmSpring.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   4935
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   1
      Left            =   1320
      Picture         =   "frmSpring.frx":2F1B
      Stretch         =   -1  'True
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   7215
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   7215
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   0
      Left            =   1320
      Picture         =   "frmSpring.frx":1A263
      Stretch         =   -1  'True
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   7215
   End
End
Attribute VB_Name = "frmSummer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Const MAX = 10
Const MSG_1 = "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
Const MSG_2 = "A Monarch Catepillar in third instar (skin has been shed three times). Here it is feeding on milkweed. Milkweed is an important source of food and a place to lay eggs for D. Plexippus. Changes in milkweed population cause huge problems with D. Plexippus. Photo by World Unit 9 on Flickr."
Const MSG_3 = "There are many generations of D. Plexippus in the late-spring/summer season. During this time they breed normally. Photo by wwarby on Flickr."

'Const NumImg = 2 'first index is 0
Const TITLE_1 = "Summer"
Dim LabelMsg(1 To MAX) As String
Dim CurrentImage As Integer
Dim NumImg As Integer
Option Explicit



Private Sub picPicture_Click(Index As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    NextImg lblText, imgPictures, NumImg, CurrentImage, True
    cmdPrevious.Enabled = True
    If CurrentImage = NumImg Then
        cmdNext.Enabled = False
    End If
    
End Sub

Private Sub cmdPrevious_Click()
    NextImg lblText, imgPictures, NumImg, CurrentImage, False
    cmdNext.Enabled = True
    If CurrentImage = 0 Then
        cmdPrevious.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    
    lblText(0) = MSG_1
    lblText(1) = MSG_2
    lblText(2) = MSG_3
    
    CurrentImage = -1
    NumImg = 2
    
    NextImg lblText, imgPictures, NumImg, CurrentImage, True
    cmdPrevious.Enabled = False
    
End Sub

