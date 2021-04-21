VERSION 2.00
Begin MDIForm frmMDI 
   Caption         =   "ADI TI-83 Assembler"
   ClientHeight    =   4785
   ClientLeft      =   1620
   ClientTop       =   1815
   ClientWidth     =   5685
   Height          =   5475
   Icon            =   MDI.FRX:0000
   Left            =   1560
   LinkTopic       =   "MDIForm1"
   Top             =   1185
   Width           =   5805
   Begin PictureBox picToolbar 
      Align           =   1  'Align Top
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      Begin CommonDialog CMDialog1 
         CancelError     =   -1  'True
         DefaultExt      =   "z80"
         Filter          =   "Z80 Files (*.z80)|*.z80|ASM Files (*.asm)|*.asm|All Files (*.*)|*.*"
         Left            =   3240
         Top             =   120
      End
      Begin CommonDialog CMDialog2 
         CancelError     =   -1  'True
         DefaultExt      =   "83p"
         DialogTitle     =   "Compile"
         Filter          =   "TI-83 Program (*.83p)|*.83p|All Files (*.*)|*.*"
         Left            =   3720
         Top             =   120
      End
      Begin CommonDialog CMDialog3 
         CancelError     =   -1  'True
         Left            =   4200
         Top             =   120
      End
      Begin Line linToolbar 
         X1              =   0
         X2              =   9600
         Y1              =   420
         Y2              =   420
      End
      Begin Image imgEditJumpButtonUp 
         Height          =   330
         Left            =   6360
         Picture         =   MDI.FRX:0302
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgFileCompileButtonUp 
         Height          =   330
         Left            =   5640
         Picture         =   MDI.FRX:0484
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgFileSaveButtonUp 
         Height          =   330
         Left            =   4920
         Picture         =   MDI.FRX:0606
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgEditJumpButtonDn 
         Height          =   330
         Left            =   6000
         Picture         =   MDI.FRX:0788
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgFileCompileButtonDn 
         Height          =   330
         Left            =   5280
         Picture         =   MDI.FRX:090A
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgFileSaveButtonDn 
         Height          =   330
         Left            =   4560
         Picture         =   MDI.FRX:0A8C
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgEditJumpButton 
         Height          =   330
         Left            =   1590
         Picture         =   MDI.FRX:0C0E
         Top             =   45
         Width           =   360
      End
      Begin Image imgFileCompileButton 
         Height          =   330
         Left            =   1245
         Picture         =   MDI.FRX:0D90
         Top             =   45
         Width           =   360
      End
      Begin Image imgFileSaveButton 
         Height          =   330
         Left            =   900
         Picture         =   MDI.FRX:0F12
         Top             =   45
         Width           =   360
      End
      Begin Image imgFileNewButton 
         Height          =   330
         Left            =   90
         Picture         =   MDI.FRX:1094
         Top             =   45
         Width           =   360
      End
      Begin Image imgFileOpenButton 
         Height          =   330
         Left            =   435
         Picture         =   MDI.FRX:1216
         Top             =   45
         Width           =   360
      End
      Begin Image imgCutButton 
         Height          =   330
         Left            =   2055
         Picture         =   MDI.FRX:1398
         Top             =   45
         Width           =   360
      End
      Begin Image imgCopyButton 
         Height          =   330
         Left            =   2400
         Picture         =   MDI.FRX:151A
         Top             =   45
         Width           =   360
      End
      Begin Image imgPasteButton 
         Height          =   330
         Left            =   2745
         Picture         =   MDI.FRX:169C
         Top             =   45
         Width           =   360
      End
      Begin Image imgFileNewButtonDn 
         Height          =   330
         Left            =   3120
         Picture         =   MDI.FRX:181E
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgFileNewButtonUp 
         Height          =   330
         Left            =   3480
         Picture         =   MDI.FRX:19A0
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgFileOpenButtonUp 
         Height          =   330
         Left            =   4200
         Picture         =   MDI.FRX:1B22
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgFileOpenButtonDn 
         Height          =   330
         Left            =   3840
         Picture         =   MDI.FRX:1CA4
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgCutButtonUp 
         Height          =   330
         Left            =   7080
         Picture         =   MDI.FRX:1E26
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgCutButtonDn 
         Height          =   330
         Left            =   6720
         Picture         =   MDI.FRX:1FA8
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgCopyButtonUp 
         Height          =   330
         Left            =   7800
         Picture         =   MDI.FRX:212A
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgCopyButtonDn 
         Height          =   330
         Left            =   7440
         Picture         =   MDI.FRX:22AC
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgPasteButtonDn 
         Height          =   330
         Left            =   8160
         Picture         =   MDI.FRX:242E
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin Image imgPasteButtonUp 
         Height          =   330
         Left            =   8520
         Picture         =   MDI.FRX:25B0
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin Menu mnuFOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin Menu mnuFAbout 
         Caption         =   "A&bout ADI TI-83 Assembler..."
      End
      Begin Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin Menu mnuFExit 
         Caption         =   "E&xit"
      End
      Begin Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
   End
   Begin Menu mnuOptions 
      Caption         =   "&Options"
      Begin Menu mnuOToolbar 
         Caption         =   "&Toolbar"
      End
      Begin Menu mnuOSep1 
         Caption         =   "-"
      End
      Begin Menu mnuODefaultFont 
         Caption         =   "Default &Font..."
      End
      Begin Menu mnuOFixedPitch 
         Caption         =   "&Show Fixed-Pitch Fonts Only"
      End
      Begin Menu mnuOSep2 
         Caption         =   "-"
      End
      Begin Menu mnuOSetup 
         Caption         =   "&Compile Setup..."
      End
   End
End

Sub imgCopyButton_Click ()
    imgCopyButton.Refresh
    EditCopyProc
End Sub

Sub imgCopyButton_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCopyButton.Picture = imgCopyButtonDn.Picture
End Sub

Sub imgCopyButton_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > imgCopyButton.Width Or Y < 0 Or Y > imgCopyButton.Height Then
            imgCopyButton.Picture = imgCopyButtonUp.Picture
        Else
            imgCopyButton.Picture = imgCopyButtonDn.Picture
        End If
    End Select
End Sub

Sub imgCopyButton_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCopyButton.Picture = imgCopyButtonUp.Picture
End Sub

Sub imgCutButton_Click ()
    imgCutButton.Refresh
    EditCutProc
End Sub

Sub imgCutButton_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCutButton.Picture = imgCutButtonDn.Picture
End Sub

Sub imgCutButton_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > imgCutButton.Width Or Y < 0 Or Y > imgCutButton.Height Then
            imgCutButton.Picture = imgCutButtonUp.Picture
        Else
            imgCutButton.Picture = imgCutButtonDn.Picture
        End If
    End Select
End Sub

Sub imgCutButton_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCutButton.Picture = imgCutButtonUp.Picture
End Sub

Sub imgEditJumpButton_Click ()
    imgEditJumpButton.Refresh
    JumpLabel
End Sub

Sub imgEditJumpButton_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEditJumpButton.Picture = imgEditJumpButtonDn.Picture
End Sub

Sub imgEditJumpButton_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > imgEditJumpButton.Width Or Y < 0 Or Y > imgEditJumpButton.Height Then
            imgEditJumpButton.Picture = imgEditJumpButtonUp.Picture
        Else
            imgEditJumpButton.Picture = imgEditJumpButtonDn.Picture
        End If
    End Select
End Sub

Sub imgEditJumpButton_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEditJumpButton.Picture = imgEditJumpButtonUp.Picture
End Sub

Sub imgFileCompileButton_Click ()
    imgFileCompileButton.Refresh
    Compile
End Sub

Sub imgFileCompileButton_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFileCompileButton.Picture = imgFileCompileButtonDn.Picture
End Sub

Sub imgFileCompileButton_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > imgFileCompileButton.Width Or Y < 0 Or Y > imgFileCompileButton.Height Then
            imgFileCompileButton.Picture = imgFileCompileButtonUp.Picture
        Else
            imgFileCompileButton.Picture = imgFileCompileButtonDn.Picture
        End If
    End Select
End Sub

Sub imgFileCompileButton_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFileCompileButton.Picture = imgFileCompileButtonUp.Picture
End Sub

Sub imgFileNewButton_Click ()
    imgFileNewButton.Refresh
    FileNew
End Sub

Sub imgFileNewButton_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFileNewButton.Picture = imgFileNewButtonDn.Picture
End Sub

Sub imgFileNewButton_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > imgFileNewButton.Width Or Y < 0 Or Y > imgFileNewButton.Height Then
            imgFileNewButton.Picture = imgFileNewButtonUp.Picture
        Else
            imgFileNewButton.Picture = imgFileNewButtonDn.Picture
        End If
    End Select
End Sub

Sub imgFileNewButton_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFileNewButton.Picture = imgFileNewButtonUp.Picture
End Sub

Sub imgFileOpenButton_Click ()
    imgFileOpenButton.Refresh
    FOpenProc
End Sub

Sub imgFileOpenButton_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFileOpenButton.Picture = imgFileOpenButtonDn.Picture
End Sub

Sub imgFileOpenButton_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > imgFileOpenButton.Width Or Y < 0 Or Y > imgFileOpenButton.Height Then
            imgFileOpenButton.Picture = imgFileOpenButtonUp.Picture
        Else
            imgFileOpenButton.Picture = imgFileOpenButtonDn.Picture
        End If
    End Select
End Sub

Sub imgFileOpenButton_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFileOpenButton.Picture = imgFileOpenButtonUp.Picture
End Sub

Sub imgFileSaveButton_Click ()
    imgFileSaveButton.Refresh
    Dim Filename As String

    If Left(frmMDI.ActiveForm.Caption, 8) = "Untitled" Then
        ' The file hasn't been saved yet,
        ' get the filename, then call the
        ' save procedure
        Filename = GetFileName()
    Else
        ' The caption contains the name of the open file
        Filename = frmMDI.ActiveForm.Caption
    End If
    ' call the save procedure, if Filename = Empty then
    ' the user selected Cancel in the Save As dialog, otherwise
    ' save the file
    If Filename <> "" Then
        SaveFileAs Filename
        ' Update the recent files menu
        UpdateFileMenu (Filename)
    End If
End Sub

Sub imgFileSaveButton_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFileSaveButton.Picture = imgFileSaveButtonDn.Picture
End Sub

Sub imgFileSaveButton_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > imgFileSaveButton.Width Or Y < 0 Or Y > imgFileSaveButton.Height Then
            imgFileSaveButton.Picture = imgFileSaveButtonUp.Picture
        Else
            imgFileSaveButton.Picture = imgFileSaveButtonDn.Picture
        End If
    End Select
End Sub

Sub imgFileSaveButton_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFileSaveButton.Picture = imgFileSaveButtonUp.Picture
End Sub

Sub imgPasteButton_Click ()
    imgPasteButton.Refresh
    EditPasteProc
End Sub

Sub imgPasteButton_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPasteButton.Picture = imgPasteButtonDn.Picture
End Sub

Sub imgPasteButton_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > imgPasteButton.Width Or Y < 0 Or Y > imgPasteButton.Height Then
            imgPasteButton.Picture = imgPasteButtonUp.Picture
        Else
            imgPasteButton.Picture = imgPasteButtonDn.Picture
        End If
    End Select
End Sub

Sub imgPasteButton_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPasteButton.Picture = imgPasteButtonUp.Picture
End Sub

Sub MDIForm_Load ()
    Screen.MousePointer = 11
    LoadSettings
    CMDialog1.Flags = &H1000& Or &H4& Or &H2&
    CMDialog2.Flags = &H4& Or &H2&
    ' Application starts here (Load event of Startup form).
    Show
    ' Always set working directory to directory containing the application.
    ChDir App.Path
    
    'Initialize document form arrays, and show first document.
    ReDim Document(1)
    ReDim FState(1)
    Document(1).Tag = 1
    FState(1).Dirty = False
    Document(1).Show

    ' Read ADI83ASM.INI and set recent file menu items appropriately
    GetRecentFiles
    Screen.MousePointer = 0
End Sub

Sub MDIForm_Unload (Cancel As Integer)
    ' If the Unload was not canceled (in the QueryUnload events for the Notepad forms)
    ' there will be no document windows left, so go ahead and end the application.

    If Not AnyPadsLeft() Then
        SaveSettings
        End
    End If

End Sub

Sub mnuFAbout_Click ()
    Screen.MousePointer = 11
    frmAboutBox.Show modal
End Sub

Sub mnuFExit_Click ()
    SaveSettings
    End
End Sub

Sub mnuFNew_Click ()
    FileNew
End Sub

Sub mnuFOpen_Click ()
    FOpenProc
End Sub

Sub mnuODefaultFont_Click ()
    DefaultFont
End Sub

Sub mnuOFixedPitch_Click ()
    OptionsFixedPitchProc Me
End Sub

Sub mnuOptions_Click ()
    mnuOToolbar.Checked = frmMDI!picToolbar.Visible
    mnuOFixedPitch.Checked = Fixed
End Sub

Sub mnuOSetup_Click ()
    CompileSetup
End Sub

Sub mnuOToolbar_Click ()
    OptionsToolbarProc Me
End Sub

Sub mnuRecentFile_Click (index As Integer)
  Filename = Right(mnuRecentFile(index).Caption, Len(mnuRecentFile(index).Caption) - 3)
  OpenFile (Filename)
  UpdateFileMenu (Filename)
  ' Update recent files list.
  GetRecentFiles
End Sub

Sub picToolbar_GotFocus ()
    If AnyPadsLeft() Then ActiveForm.SetFocus
End Sub

Sub picToolbar_Resize ()
    linToolbar.X2 = picToolbar.Width
End Sub

