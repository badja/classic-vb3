Declare Function GetModuleUsage% Lib "Kernel" (ByVal hModule%)

Sub Compile ()
    SourcePathFilename = frmMDI.ActiveForm.Caption
    If Left(SourcePathFilename, 8) = "Untitled" Or FState(frmMDI.ActiveForm.Tag).Dirty Then
	MsgBox "You must save the file before compiling.", 64, "Compile"
	Exit Sub
    End If
    
    'Displays a Save As dialog and returns a file name
    'or an empty string if the user cancels
    On Error Resume Next
    frmMDI.CMDialog2.Filename = GetStem(SourcePathFilename) & ".83p"
    frmMDI.CMDialog2.Action = 2
    If Err <> 32755 Then      'User has not cancelled dialog
	TASMPath = TrimSlash(frmSetup.txtTASM)
	If TASMPath = "" Then TASMPath = "."
	ProgramPathFilename = frmMDI.CMDialog2.Filename
	ProgramPathFileStem = GetStem(ProgramPathFilename)
	ProgramPath = TrimSlash(GetPath(ProgramPathFilename))
	ProgramFilename = frmMDI.CMDialog2.Filetitle
	ProgramFileStem = GetStem(ProgramFilename)
	ProgramName = ProgramFileStem
	If frmSetup.chkAutoName = 0 Then
	    ProgramName = UCase(Left(InputBox("Enter the name for the program:", "Program Name", ProgramName), 8))
	    If ProgramName = "" Then Exit Sub
	End If
	If frmSetup.chkKeep Then Keep = "k" Else Keep = "d"
	LinkPath = frmSetup.txt83lnk
	If Right(LinkPath, 1) <> "\" And LinkPath <> "" Then LinkPath = LinkPath & "\"
	OriDir = CurDir
	ChDir App.Path
	Open "~aditemp.bat" For Output As #1
	Print #1, "@echo off"
	Print #1, "call adi1 " & TASMPath & " " & SourcePathFilename & " " & ProgramPathFileStem & " " & Keep
	Print #1, "cd " & ProgramPath
	Print #1, "call adi2 " & ProgramFileStem & " " & ProgramFilename & " " & ProgramName & " " & Keep & " " & LinkPath
	Print #1, ":end"
	Print #1, "pause"
	Close #1
	Screen.MousePointer = 11
	x% = Shell("~aditemp.bat", 1)
	While GetModuleUsage(x%) > 0    ' Has Shelled program finished?
	    z% = DoEvents()             ' If not, yield to Windows.
	Wend
	ChDir App.Path
	Kill "~aditemp.bat"
	ChDir OriDir
	Screen.MousePointer = 0
    End If
End Sub

Sub FOpenProc ()
    Dim retval
    On Error Resume Next
    Dim OpenFileName As String
    frmMDI.CMDialog1.Filename = ""
    frmMDI.CMDialog1.Action = 1
    If Err <> 32755 Then 'user pressed cancel
	OpenFileName = frmMDI.CMDialog1.Filename
	OpenFile (OpenFileName)
	UpdateFileMenu (OpenFileName)
    End If
End Sub

Function GetFileName ()
    'Displays a Save As dialog and returns a file name
    'or an empty string if the user cancels
    On Error Resume Next
    frmMDI.CMDialog1.Filename = ""
    frmMDI.CMDialog1.Action = 2
    If Err <> 32755 Then      'User cancelled dialog
	GetFileName = frmMDI.CMDialog1.Filename
    Else
	GetFileName = ""
    End If
End Function

Function GetPath (ByVal Filename As String) As String
    For Find = Len(Filename) To 1 Step -1
	If Mid(Filename, Find, 1) = "\" Then
	    GetPath = Left(Filename, Find)
	    Exit For
	End If
    Next Find
End Function

Function GetStem (ByVal Filename As String) As String
    Dot = InStr(Filename, ".")
    If Dot Then
	GetStem = Left(Filename, Dot - 1)
    Else
	GetStem = Filename
    End If
End Function

Function OnRecentFilesList (Filename) As Integer
  Dim i

  For i = 1 To 4
    If Right(frmMDI.mnuRecentFile(i).Caption, Len(frmMDI.mnuRecentFile(i).Caption) - 3) = Filename Then
      OnRecentFilesList = i
      Exit Function
    End If
  Next i
    OnRecentFilesList = 0
End Function

Sub OpenFile (Filename)
    Dim NL, TextIn, GetLine
    Dim fIndex As Integer

    NL = Chr$(13) + Chr$(10)
    
    On Error Resume Next
    ' open the selected file
    Open Filename For Input As #1
    If Err Then
	MsgBox "Can't open file: " + Filename, 48
	Exit Sub
    End If
    ' change mousepointer to an hourglass
    Screen.MousePointer = 11
    
    ' change form's caption and display new text
    fIndex = FindFreeIndex()
    Opening = True
    Small = False
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = UCase$(Filename)
    Document(fIndex).Text1.Text = Input$(LOF(1), 1)
    FState(fIndex).Dirty = False
    Document(fIndex).Show
    Close #1
    ' reset mouse pointer
    Screen.MousePointer = 0
    If Small Then
	If Not Started And Not FState(1).Dirty Then
	    Started = True
	    Unload Document(1)
	End If
	' Make sure toolbar edit buttons are visible
	frmMDI!imgFileSaveButton.Visible = True
	frmMDI!imgFileCompileButton.Visible = True
	frmMDI!imgEditJumpButton.Visible = True
	frmMDI!imgCutButton.Visible = True
	frmMDI!imgCopyButton.Visible = True
	frmMDI!imgPasteButton.Visible = True
    Else
	MsgBox "File is too large to open: " + Filename, 48
	Unload Document(fIndex)
    End If
End Sub

Sub SaveFileAs (Filename)
On Error Resume Next
    Dim Contents As String

    ' open the file
    Open Filename For Output As #1
    ' put contents of the notepad into a variable
    Contents = frmMDI.ActiveForm.Text1.Text
    ' display hourglass
    Screen.MousePointer = 11
    ' write variable contents to saved file
    Print #1, Contents;
    Close #1
    ' reset the mousepointer
    Screen.MousePointer = 0
    ' set the Notepad's caption

    If Err Then
	MsgBox Error, 48, App.Title
    Else
	frmMDI.ActiveForm.Caption = UCase$(Filename)
	' reset the dirty flag
	FState(frmMDI.ActiveForm.Tag).Dirty = False
    End If
End Sub

Function TrimSlash (ByVal Path As String) As String
    If Right(Path, 1) = "\" And Right(Path, 2) <> ":\" Then
	TrimSlash = Left(Path, Len(Path) - 1)
    Else
	TrimSlash = Path
    End If
End Function

Sub UpdateFileMenu (Filename)
	Dim retval
	' Check if OpenFileName is already on MRU list.
	retval = OnRecentFilesList(Filename)
	If retval = 0 Then
	    ' Write OpenFileName to ADI83ASM.INI
	    WriteRecentFiles (Filename)
	Else
	    ShiftRecentFiles (retval)
	End If
	' Update menus for most recent file list.
	GetRecentFiles
End Sub

