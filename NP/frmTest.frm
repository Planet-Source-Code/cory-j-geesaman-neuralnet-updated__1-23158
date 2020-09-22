VERSION 5.00
Object = "*\AprjNP.vbp"
Begin VB.Form frmTest 
   Caption         =   "NeuralNet - By: Cory J. Geesaman"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6885
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "&Diminsions"
      Height          =   2295
      Left            =   5640
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "10"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "10"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "30"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Depth"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Height:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Width:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   465
      End
   End
   Begin prjNP.NeuralProcessor NeuralProcessor 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7435
      BackColor       =   0
      BorderStyle     =   1
      Picture         =   "frmTest.frx":0442
   End
   Begin VB.CommandButton Command3 
      Caption         =   "S&top Net"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Start Net"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Init Net"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If NeuralProcessor.InitNet(CLng(Text1.Text), CLng(Text2.Text), CLng(Text3.Text)) = False Then MsgBox "Error initializing NeuralNet", vbCritical, "An Error Occoured"
End Sub

Private Sub Command2_Click()
NeuralProcessor.StartNet
End Sub

Private Sub Command3_Click()
NeuralProcessor.StopNet
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
