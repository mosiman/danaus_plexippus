VERSION 5.00
Begin VB.Form frmMigrate 
   Caption         =   "Migration"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGraphs 
      Caption         =   "Show Data"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "---->"
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<----"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   7215
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   7215
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   4
      Left            =   1440
      Picture         =   "frmMigrate.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4935
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   3
      Left            =   1440
      Picture         =   "frmMigrate.frx":2C233
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   7215
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   0
      Left            =   1440
      Picture         =   "frmMigrate.frx":42F1A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   7215
   End
   Begin VB.Label lblText 
      Caption         =   "A Monarch Butterfly (D. Plexippus) laying eggs on milkweed (late-spring generation). Photo by Frank Matheson."
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   7215
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   1
      Left            =   1440
      Picture         =   "frmMigrate.frx":492C8
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4935
   End
   Begin VB.Image imgPictures 
      Height          =   3735
      Index           =   2
      Left            =   1440
      Picture         =   "frmMigrate.frx":6FC96
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Late-Summer ---- Fall"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmMigrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAX = 10
Const MSG_1 = "Environmental cues such as weather trigger physiological changes in D. Plexippus. They go into reproductive diapause and head south to warmer climates for the winter. Lack of mositure and extremely cold temperatures is what drives them south. Photo by Paul Opier. "
Const MSG_2 = "Milkweed is the primary source of food for D. Plexippus. Instead of using energy to mate and reproduce, energy is used to travel (diapause). D. Plexippus while migrating can travel upwards of 25 miles a day. Photo by Marty Nevils Davis."
Const MSG_3 = "Considered a common weed, glyphosates are diminishing milkweed populations. Glyphosate-resistant crops such as corn are becoming increasingly large in numbers. Photo by niklask on DeviantArt."
Const MSG_4 = "Increased revenue from corn is favourable economically, thus it is hard to convince people of the extent of its effects on D. Plexippus and other organisms. Image by Hasbro."
Const MSG_5 = "As temperature decreases, precipitation and unexpected weather pose greater hazards to Monarchs. D. Plexippus can survive low temperature in the winter, but may freeze to death if wet. Image by Howie Garber."


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
        'show form for graph corn vs butterfly
        Load frmCornVsPopVsTime
        frmCornVsPopVsTime.Show vbModal
    ElseIf CurrentImage = 3 Then
        'show form for graph money vs butterfly
        Load frmMoneyVsPopVsTime
        frmMoneyVsPopVsTime.Show vbModal
    Else
        MsgBox "No data to display.", vbInformation, "No Data"
    End If
    
End Sub

Private Sub Form_Load()
    lblText(0) = MSG_1
    lblText(1) = MSG_2
    lblText(2) = MSG_3
    lblText(3) = MSG_4
    lblText(4) = MSG_5
    
    CurrentImage = -1
    NumImg = 4 '0 is first index
    
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
