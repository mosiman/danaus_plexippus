VERSION 5.00
Begin VB.Form frmOverWintering 
   Caption         =   "Overwintering"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGraphs 
      Caption         =   "Show Data"
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<----"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "---->"
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Winter"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   2
      Left            =   1320
      Picture         =   "frmOverWintering.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4935
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   1
      Left            =   1320
      Picture         =   "frmOverWintering.frx":22D9E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   7215
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   7215
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   0
      Left            =   1320
      Picture         =   "frmOverWintering.frx":3EDEF
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   7215
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   3
      Left            =   1320
      Picture         =   "frmOverWintering.frx":C75EB
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4935
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   4
      Left            =   1320
      Picture         =   "frmOverWintering.frx":DE2D2
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   7215
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   7215
   End
End
Attribute VB_Name = "frmOverWintering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAX = 10
Const MSG_1 = "D. Plexippus migrate to Oyamel trees (abies religiosa) located in the Trans-Mexican Volcanic Belt. Photo by Mihai Costea"
Const MSG_2 = "By clustering tightly and using the location of Oyamel, D. Plexippus create a microenvironment that allow it to minimize energy usage (cold blooded organisms use less energy in cold environments). Photo by Pablo Leautaud."
Const MSG_3 = "Oyamel population is declining due to illegal logging. Despite efforts by the Mexican government, logging still continues. Oyamel population is expected to be 3.5% of its current population by 2090 (C. Saenz-Romero et .al). Photo by World Wildlife Foundation."



'Const NumImg = 2 'first index is 0
Const TITLE_1 = "Summer"
Dim LabelMsg(1 To MAX) As String
Dim CurrentImage As Integer
Dim NumImg As Integer

Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGraphs_Click()
    If CurrentImage = 2 Then
        Load frmPopVsOyamel
        frmPopVsOyamel.Show vbModal
    Else
        MsgBox "No data to display", vbInformation, "No Data"
    End If
End Sub

Private Sub Form_Load()
    lblText(0) = MSG_1
    lblText(1) = MSG_2
    lblText(2) = MSG_3
    
    CurrentImage = -1
    NumImg = 2 '0 is first index
    
    NextImg lblText, imgPictures, NumImg, CurrentImage, True
    cmdPrevious.Enabled = False
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
