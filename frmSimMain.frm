VERSION 5.00
Begin VB.Form frmSimMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Migration Routes"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChooseGraph 
      Caption         =   "Graphs"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox chkButterfly 
      Caption         =   "D. Plexippus images off"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdSimStop 
      Caption         =   "Stop Simulation"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13200
      Top             =   5040
   End
   Begin VB.CommandButton cmdSimRun 
      Caption         =   "Run Time Simulation"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton YearMinusOne 
      Caption         =   "-1"
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton YearPlusOne 
      Caption         =   "+1"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1 butterfly: 10000000 D. Plexippus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   12
      Top             =   8160
      Width           =   3975
   End
   Begin VB.Label lblOverwinteringText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3480
      TabIndex        =   9
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label lblMigrateText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4320
      TabIndex        =   7
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label lblMigrate 
      BackStyle       =   0  'Transparent
      Height          =   3255
      Left            =   3960
      TabIndex        =   6
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Label lblStartText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      TabIndex        =   5
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   8775
   End
   Begin VB.Label lblPopulation 
      BackStyle       =   0  'Transparent
      Caption         =   "Population:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   7560
      Width           =   5175
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Image imgBackground 
      Height          =   8925
      Left            =   0
      Picture         =   "frmSimMain.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12405
   End
   Begin VB.Image imgDPlexippus 
      Height          =   480
      Index           =   0
      Left            =   13080
      Picture         =   "frmSimMain.frx":1BE420
      Stretch         =   -1  'True
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmSimMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SimStop As Integer
Option Explicit

Private Sub chkButterfly_Click()
    Dim K As Integer
    
    If chkButterfly.Value = 1 Then
        For K = 1 To imgDPlexippus.UBound
            imgDPlexippus(K).Visible = False
        Next K
    Else
        For K = 1 To imgDPlexippus.UBound
            imgDPlexippus(K).Visible = True
        Next K
    End If
End Sub

Private Sub cmdChooseGraphs_Click()

End Sub

Private Sub cmdChooseGraph_Click()
    Load frmChooseGraph
    frmChooseGraph.Show vbModal
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSimRun_Click()
    'input box inputs starting year, simulation goes till population = 0
    'increase yearcount in relation to timer
    
    Dim UserIn As Integer
    UserIn = Val(InputBox$("enter a year to start simulation"))
    If UserIn > 1990 And UserIn < 2216 Then
        YearCount = UserIn
        tmrTime.Enabled = True
    Else
        MsgBox "Invalid input. Valid input: 1990 < t < 2216", vbInformation, "Invalid Input"
    End If
End Sub

Private Sub cmdSimStop_Click()
    tmrTime.Enabled = False
End Sub

Private Sub Form_Load()
    'ImgCount = 0
    YearCount = 2010
    Temperature = 35
    'PopulationCount = 1000000
    
    UpdateForm YearCount, PopulationCount, Temperature
    'CreateNewImage frmSimMain
    
    Randomize
    
End Sub


Private Sub imgBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStartText.Caption = ""
    lblMigrateText.Caption = ""
    lblOverwinteringText.Caption = ""
End Sub

Private Sub lblMigrate_Click()
    Load frmMigrate
    frmMigrate.Show vbModal
End Sub

Private Sub lblMigrate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMigrateText.Caption = "Click here to learn more about migration"
    lblMigrateText.ZOrder vbBringToFront
End Sub

Private Sub lblMigrateText_Click()
    Load frmMigrate
    frmMigrate.Show vbModal
End Sub

Private Sub lblOverwinteringText_Click()
    Load frmOverWintering
    frmOverWintering.Show vbModal
End Sub

Private Sub lblOverwinteringText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblOverwinteringText.Caption = "Click here to learn more about overwintering"
    lblOverwinteringText.ZOrder vbBringToFront
End Sub

Private Sub lblPopulation_Click()
    Load frmPopVsTime
    frmPopVsTime.Show vbModal
End Sub

Private Sub lblStart_Click()
    Load frmSummer
    frmSummer.Show vbModal
End Sub

Private Sub lblStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStartText.Caption = "Click here to learn more about non-migrational generations"
    lblStartText.ZOrder vbBringToFront
    
End Sub

Private Sub lblStartText_Click()
    Load frmSummer
    frmSummer.Show vbModal
End Sub

Private Sub lblYear_Click()
     Dim UserInput As Integer
     
     UserInput = Val(InputBox$("Enter a year:", "Year Input"))
     
     If UserInput > 0 Then
        YearCount = UserInput
        UpdateForm YearCount, Temperature, PopulationCount
    Else
        MsgBox "Invalid input", vbExclamation, "Invalid Input"
    End If
    
End Sub

Private Sub tmrTime_Timer()

    If YearCount <> 2216 Then
        YearPlusOne_Click
    Else
        tmrTime.Enabled = False
    End If
    
End Sub

Private Sub YearMinusOne_Click()
    YearCount = YearCount - 1
    
    UpdateForm YearCount, Temperature, PopulationCount

End Sub

Private Sub YearPlusOne_Click()
    YearCount = YearCount + 1
    
    UpdateForm YearCount, Temperature, PopulationCount
End Sub
