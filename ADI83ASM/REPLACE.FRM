VERSION 2.00
Begin Form frmReplace 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Replace"
   ClientHeight    =   1755
   ClientLeft      =   2655
   ClientTop       =   3585
   ClientWidth     =   5310
   Height          =   2160
   Left            =   2595
   LinkTopic       =   "Form2"
   ScaleHeight     =   1755
   ScaleWidth      =   5310
   Top             =   3240
   Width           =   5430
   Begin TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin CheckBox chkCase 
      Caption         =   "Match &Case"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin Frame Frame1 
      Caption         =   "Direction"
      Height          =   612
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   2052
      Begin OptionButton optDirection 
         Caption         =   "&Up"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   612
      End
      Begin OptionButton optDirection 
         Caption         =   "&Down"
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin CommandButton cmdReplace 
      Caption         =   "&Replace"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   1080
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
   Begin Label Label2 
      Caption         =   "Re&place With:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
End

Sub chkCase_Click ()
    gFindCase = chkCase.Value
End Sub

Sub cmdCancel_Click ()
    Unload frmReplace
End Sub

Sub cmdReplace_Click ()
    gFindString = Text1.Text
    FindIt 1
End Sub

Sub cmdReplaceAll_Click ()
    gFindString = Text1.Text
    frmMDI.ActiveForm.Text1.SelLength = 0
    If gFindDirection = 0 Then
        frmMDI.ActiveForm.Text1.SelStart = Len(frmMDI.ActiveForm.Text1)
    Else
        frmMDI.ActiveForm.Text1.SelStart = 0
    End If
    gFirstTime = True
    Replacing = False
    FindIt 1
    Screen.MousePointer = 11
    Do Until Replacing = False
        FindIt 2
    Loop
    Screen.MousePointer = 0
End Sub

Sub Form_Activate ()
    Screen.MousePointer = 0
End Sub

Sub Form_Load ()
    cmdReplace.Enabled = False
    cmdReplaceAll.Enabled = False
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
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    Else
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
    End If
    Replacing = False
End Sub

Sub Text1_GotFocus ()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Sub Text2_GotFocus ()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)
End Sub

