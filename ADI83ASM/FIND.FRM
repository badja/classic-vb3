VERSION 2.00
Begin Form frmFind 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Find"
   ClientHeight    =   1395
   ClientLeft      =   2655
   ClientTop       =   3585
   ClientWidth     =   5070
   Height          =   1800
   Left            =   2595
   LinkTopic       =   "Form2"
   ScaleHeight     =   1395
   ScaleWidth      =   5070
   Top             =   3240
   Width           =   5190
   Begin TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin CheckBox chkCase 
      Caption         =   "Match &Case"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin Frame Frame1 
      Caption         =   "Direction"
      Height          =   612
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   2052
      Begin OptionButton optDirection 
         Caption         =   "&Up"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   612
      End
      Begin OptionButton optDirection 
         Caption         =   "&Down"
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin Label Label1 
      Caption         =   "Fi&nd What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End

Sub chkCase_Click ()
    gFindCase = chkCase.Value
End Sub

Sub cmdCancel_Click ()
    Unload frmFind
End Sub

Sub cmdFind_Click ()
    gFindString = Text1.Text
    FindIt 0
End Sub

Sub Form_Activate ()
    Screen.MousePointer = 0
End Sub

Sub Form_Load ()
    cmdFind.Enabled = False
    gFindDirection = 1
End Sub

Sub Form_Unload (Cancel As Integer)
    gFindString = Text1.Text
    gFindCase = chkCase.Value
End Sub

Sub optDirection_Click (index As Integer)
    gFindDirection = index
End Sub

Sub Text1_Change ()
    FirstTime = True

    If Text1.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
End Sub

Sub Text1_GotFocus ()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

