VERSION 5.00
Begin VB.Form frmChooseGraph 
   Caption         =   "Choose a Graph"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Click to view graph."
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monarch Poplation vs Oyamel Population vs Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monarch Poplation vs Money vs Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   6615
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monarch Poplation vs Corn Production vs Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monarch Poplation vs Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   6615
   End
End
Attribute VB_Name = "frmChooseGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Label1_Click()
    Load frmPopVsTime
    frmPopVsTime.Show vbModal
End Sub

Private Sub Label2_Click()
    Load frmCornVsPopVsTime
    frmCornVsPopVsTime.Show vbModal
End Sub

Private Sub Label3_Click()
    Load frmMoneyVsPopVsTime
    frmMoneyVsPopVsTime.Show vbModal
End Sub

Private Sub Label4_Click()
    Load frmPopVsOyamel
    frmPopVsOyamel.Show vbModal
End Sub
