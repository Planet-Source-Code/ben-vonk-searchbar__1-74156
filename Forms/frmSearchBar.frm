VERSION 5.00
Begin VB.Form frmSearchBar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SearchBar Demo"
   ClientHeight    =   3132
   ClientLeft      =   2016
   ClientTop       =   1860
   ClientWidth     =   3972
   FillColor       =   &H00404040&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBarBorder 
      Caption         =   "&BarBorder"
      Height          =   252
      Left            =   1440
      TabIndex        =   13
      Top             =   1200
      Width           =   2415
   End
   Begin prjSearchBar.SearchBar sbrSearching 
      Height          =   1212
      Left            =   120
      TabIndex        =   12
      Top             =   180
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   2138
   End
   Begin VB.ComboBox cmbBars 
      Height          =   348
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   612
   End
   Begin VB.CheckBox chkBarStyle 
      Caption         =   "&Large BarStyle"
      Height          =   252
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkKeepSearchPointer 
      Caption         =   "&Keep search position"
      Height          =   252
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.CheckBox chkRotation 
      Caption         =   "&Rotate ClockWise"
      Height          =   252
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cbdQuit 
      Caption         =   "&Quit"
      Height          =   372
      Left            =   2760
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "St&op"
      Height          =   372
      Left            =   1320
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   1
      Left            =   3000
      TabIndex        =   11
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label lblInfo 
      Caption         =   "Bars:"
      Height          =   252
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   612
   End
   Begin VB.Label lblInfo 
      Caption         =   "Cycles:"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   852
   End
   Begin VB.Line linSearchBar 
      Index           =   1
      X1              =   10
      X2              =   320
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Line linSearchBar 
      Index           =   0
      X1              =   10
      X2              =   320
      Y1              =   170
      Y2              =   170
   End
   Begin VB.Label lblBars 
      Caption         =   "Number of Bars:"
      Height          =   252
      Left            =   1440
      TabIndex        =   4
      Top             =   1608
      Width           =   1692
   End
End
Attribute VB_Name = "frmSearchBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbdQuit_Click()

   Unload Me

End Sub

Private Sub chkBarBorder_Click()

   sbrSearching.BarBorder = (chkBarBorder.Value = vbChecked)

End Sub

Private Sub chkBarStyle_Click()

   sbrSearching.BarStyle = Abs((chkBarStyle.Value = vbChecked))

End Sub

Private Sub chkKeepSearchPointer_Click()

   sbrSearching.KeepSearchPointer = (chkKeepSearchPointer.Value = vbChecked)

End Sub

Private Sub chkRotation_Click()

   sbrSearching.Rotation = Abs((chkRotation.Value = vbUnchecked))

End Sub

Private Sub cmbBars_Click()

   sbrSearching.BarsDisplayed = cmbBars.ListIndex

End Sub

Private Sub cmdStart_Click()

   If Not sbrSearching.Start Then
      lblValue.Item(0).Caption = 0
      lblValue.Item(1).Caption = 0
      sbrSearching.Start = True
   End If

End Sub

Private Sub cmdStop_Click()

   If sbrSearching.Start Then sbrSearching.Start = False

End Sub

Private Sub Form_Load()

Dim intCount As Integer

   With sbrSearching
      chkRotation.Value = Abs(Not .Rotation)
      chkKeepSearchPointer.Value = Abs(.KeepSearchPointer)
      chkBarBorder.Value = Abs(.BarBorder)
   End With
   
   With cmbBars
      For intCount = 4 To 12 Step 2
         .AddItem intCount
      Next 'intCount
      
      .ListIndex = 2
      lblBars.Top = .Top + (.Height - lblBars.Height) / 2
   End With

End Sub

Private Sub sbrSearching_Cycle(Cycles As Double, Bars As Integer)

   lblValue.Item(0).Caption = Cycles
   lblValue.Item(1).Caption = Bars

End Sub
