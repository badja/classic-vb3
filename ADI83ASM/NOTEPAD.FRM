VERSION 2.00
Begin Form frmNotePad 
   Caption         =   "Untitled"
   ClientHeight    =   3975
   ClientLeft      =   885
   ClientTop       =   2910
   ClientWidth     =   5610
   Height          =   4665
   Icon            =   NOTEPAD.FRX:0000
   Left            =   825
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   5610
   Top             =   2280
   Visible         =   0   'False
   Width           =   5730
   Begin TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   3855
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5655
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
      Begin Menu mnuFClose 
         Caption         =   "&Close"
      End
      Begin Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin Menu mnuFSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin Menu mnuFSaveAs 
         Caption         =   "Save &As..."
      End
      Begin Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin Menu mnuFCompile 
         Caption         =   "Co&mpile..."
         Shortcut        =   ^M
      End
      Begin Menu mnuFAbout 
         Caption         =   "A&bout ADI TI-83 Assembler..."
      End
      Begin Menu mnuFSep3 
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
   Begin Menu mnuEdit 
      Caption         =   "&Edit"
      Begin Menu mnuECut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin Menu mnuECopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin Menu mnuEPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin Menu mnuEDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin Menu mnuESep1 
         Caption         =   "-"
      End
      Begin Menu mnuESelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin Menu mnuETime 
         Caption         =   "Time/&Date"
      End
      Begin Menu mnuEJump 
         Caption         =   "&Jump to Label"
         Shortcut        =   ^J
      End
   End
   Begin Menu mnuSearch 
      Caption         =   "&Search"
      Begin Menu mnuSFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin Menu mnuSFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin Menu mnuSReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^R
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
      Begin Menu mnuOFont 
         Caption         =   "&Font..."
      End
      Begin Menu mnuODefaultFont 
         Caption         =   "&Default Font..."
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
   Begin Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin Menu mnuWCascade 
         Caption         =   "&Cascade"
      End
      Begin Menu mnuWTileHoriz 
         Caption         =   "Tile &Horizontally"
      End
      Begin Menu mnuWTileVert 
         Caption         =   "&Tile Vertically"
      End
      Begin Menu mnuWArrange 
         Caption         =   "&Arrange Icons"
      End
   End
End

Sub Form_Activate ()
    Replacing = False
    frmMDI.ActiveForm.Text1.SetFocus
End Sub

Sub Form_Load ()
    WindowState = DocState
    Text1.FontName = frmMDI!picToolbar.FontName
    Text1.FontSize = frmMDI!picToolbar.FontSize
    Text1.FontBold = frmMDI!picToolbar.FontBold
    Text1.FontItalic = frmMDI!picToolbar.FontItalic
    If Not Opening Then
	On Error Resume Next
	' open the header file
	If Right(App.Path, 1) <> "\" Then Slash = "\"
	Filename = App.Path & Slash & "HEADER.ASM"
	Open Filename For Input As #1
	If Err Then
	    MsgBox "Can't open header file: " + Filename, 48
	Else
	    Small = False
	    Text1 = Input$(LOF(1), 1)
	    If Not Small Then
		MsgBox "Header file is too large to open: " + Filename, 48
	    End If
	End If
	Close #1
    Else
	Opening = False
    End If
End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)
    Dim Msg, Filename, NL
    Dim Response As Integer

    If FState(Me.Tag).Dirty Then
	Filename = Me.Caption
	NL = Chr$(10) & Chr$(13)
	Msg = "The text in [" & Filename & "] has changed."
	Msg = Msg & NL
	Msg = Msg & "Do you want to save the changes?"
	Response = MsgBox(Msg, 51, frmMDI.Caption)
	Select Case Response
	' User selects Yes
	Case 6
	    'Get the filename to save the file
	    If Left(Filename, 8) = "Untitled" Then
		' The file hasn't been saved yet,
		' get the filename, then call the
		' save procedure
		Filename = GetFileName()
	    End If
	    'If the user did notspecify a file name,
	    'cancel the unload; otherwise, save it.
	    If Filename = "" Then
		Cancel = True
	    Else
		SaveFileAs (Filename)
		UpdateFileMenu (Filename)
	    End If

	' User selects No
	' Ok to unload
	Case 7
	    Cancel = False
	' User selects Cancel
	' Cancel the unload
	Case 2
	    Cancel = True
	End Select
    End If
End Sub

Sub Form_Resize ()
    If WindowState <> 1 And ScaleHeight <> 0 Then
	Text1.Visible = False
	Text1.Height = ScaleHeight + 15
	Text1.Width = ScaleWidth + 15
	Text1.Visible = True
    End If
    
    DocState = WindowState
    If DocState = 1 Then DocState = 0
End Sub

Sub Form_Unload (Cancel As Integer)
    FState(Me.Tag).Deleted = True
    
'Hide toolbar edit buttons if no notepad windows
    If Not AnyPadsLeft() Then
	frmMDI!imgFileSaveButton.Visible = False
	frmMDI!imgFileCompileButton.Visible = False
	frmMDI!imgEditJumpButton.Visible = False
	frmMDI!imgCutButton.Visible = False
	frmMDI!imgCopyButton.Visible = False
	frmMDI!imgPasteButton.Visible = False
    End If
End Sub

Sub mnuECopy_Click ()
    EditCopyProc
End Sub

Sub mnuECut_Click ()
    EditCutProc
End Sub

Sub mnuEDelete_Click ()
  ' If cursor is not at the end of the notepad.
  If Screen.ActiveControl.SelStart <> Len(Screen.ActiveControl.Text) Then
    ' If nothing is selected, extend selection by one.
    If Screen.ActiveControl.SelLength = 0 Then
      Screen.ActiveControl.SelLength = 1
      ' If cursor is on a blank line, extend selection by two.
      If Asc(Screen.ActiveControl.SelText) = 13 Then
	Screen.ActiveControl.SelLength = 2
      End If
    End If
    ' Delete selected text.
    Screen.ActiveControl.SelText = ""
  End If
End Sub

Sub mnuEJump_Click ()
    JumpLabel
End Sub

Sub mnuEPaste_Click ()
    EditPasteProc
End Sub

Sub mnuESelectAll_Click ()
    frmMDI.ActiveForm.Text1.SelStart = 0
    frmMDI.ActiveForm.Text1.SelLength = Len(frmMDI.ActiveForm.Text1.Text)
End Sub

Sub mnuETime_Click ()
    Dim TimeStr As String, DateStr As String
    
    Text1.SelText = Now
End Sub

Sub mnuFAbout_Click ()
    Screen.MousePointer = 11
    frmAboutBox.Show modal
End Sub

Sub mnuFClose_Click ()
    Unload Me
End Sub

Sub mnuFCompile_Click ()
    Compile
End Sub

Sub mnuFExit_Click ()
    ' Unloading the MDI form invokes the QueryUnload event
    ' for each child form, then the MDI form - before unloading
    ' the MDI form. Setting the Cancel argument to True in any of the
    ' QueryUnload events aborts the unload.

    Unload frmMDI


End Sub

Sub mnuFNew_Click ()
    FileNew
End Sub

Sub mnuFOpen_Click ()
    FOpenProc
End Sub

Sub mnuFSave_Click ()
    Dim Filename As String

    If Left(Me.Caption, 8) = "Untitled" Then
	' The file hasn't been saved yet,
	' get the filename, then call the
	' save procedure
	Filename = GetFileName()
    Else
	' The caption contains the name of the open file
	Filename = Me.Caption
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

Sub mnuFSaveAs_Click ()
    Dim SaveFileName As String

    SaveFileName = GetFileName()
    If SaveFileName <> "" Then SaveFileAs (SaveFileName)
    ' Update the recent files menu
    UpdateFileMenu (SaveFileName)
End Sub

Sub mnuODefaultFont_Click ()
    DefaultFont
End Sub

Sub mnuOFixedPitch_Click ()
    OptionsFixedPitchProc Me
End Sub

Sub mnuOFont_Click ()
    On Error Resume Next
    If Fixed Then
	frmMDI.CMDialog3.Flags = &H3& Or &H4000&
    Else
	frmMDI.CMDialog3.Flags = &H3&
    End If
    frmMDI.CMDialog3.FontName = Text1.FontName
    frmMDI.CMDialog3.FontSize = Text1.FontSize
    frmMDI.CMDialog3.FontBold = Text1.FontBold
    frmMDI.CMDialog3.FontItalic = Text1.FontItalic
    frmMDI.CMDialog3.Action = 4
    Screen.MousePointer = 0
    If Err <> 32755 Then 'user has not pressed cancel
	Text1.FontName = frmMDI.CMDialog3.FontName
	Text1.FontSize = frmMDI.CMDialog3.FontSize
	Text1.FontBold = frmMDI.CMDialog3.FontBold
	Text1.FontItalic = frmMDI.CMDialog3.FontItalic
    End If
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
  ' Update recent files list for new notepad.
  GetRecentFiles
End Sub

Sub mnuSFind_Click ()
    If frmReplace.Visible Then Unload frmReplace
    If Me!Text1.SelText <> "" Then
	frmFind!Text1.Text = Me!Text1.SelText
    Else
	frmFind!Text1.Text = FindString
    End If
    gFirstTime = True
    Screen.MousePointer = 11
    frmFind.Show
End Sub

Sub mnuSFindNext_Click ()
    If Len(gFindString) > 0 Then
	FindIt 0
    Else
	mnuSFind_Click
    End If
End Sub

Sub mnuSReplace_Click ()
    If frmFind.Visible Then Unload frmFind
    If Me!Text1.SelText <> "" Then
	frmReplace!Text1.Text = Me!Text1.SelText
    Else
	frmReplace!Text1.Text = FindString
    End If
    gFirstTime = True
    Screen.MousePointer = 11
    frmReplace.Show
End Sub

Sub mnuWArrange_Click ()
    frmMDI.Arrange ARRANGE_ICONS
End Sub

Sub mnuWCascade_Click ()
    frmMDI.Arrange CASCADE
End Sub

Sub mnuWTileHoriz_Click ()
    frmMDI.Arrange TILE_HORIZONTAL
End Sub

Sub mnuWTileVert_Click ()
    frmMDI.Arrange TILE_VERTICAL
End Sub

Sub Text1_Change ()
    If Me.Tag <> "" Then FState(Me.Tag).Dirty = True
    Small = True
End Sub

Sub Text1_GotFocus ()
    If frmFind.Visible Then
	frmFind.ZOrder 0
    End If
    If frmReplace.Visible Then
	frmReplace.ZOrder 0
    End If
End Sub

