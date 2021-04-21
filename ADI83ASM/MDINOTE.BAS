Option Explicit

Global Const modal = 1
Global Const CASCADE = 0
Global Const TILE_HORIZONTAL = 1
Global Const TILE_VERTICAL = 2
Global Const ARRANGE_ICONS = 3

Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type
Global FState()  As FormState
Global Document() As New frmNotePad
Global gFindString, gFindCase As Integer, gFindDirection As Integer
Global gCurPos As Integer, gFirstTime As Integer
Global ArrayNum As Integer

' API functions used to read and write to ADI83ASM.INI.
' Used for handling the recent files list.
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Integer

Global DocState As Integer
Global Fixed As Integer
Global Opening As Integer
Global Replacing As Integer
Global Small As Integer
Global Started As Integer

Function AnyPadsLeft () As Integer
    Dim i As Integer

    ' Cycle throught the document array.
    ' Return True if there is at least one
    ' open document remaining.
    For i = 1 To UBound(Document)
        If Not FState(i).Deleted Then
            AnyPadsLeft = True
            Exit Function
        End If
    Next

End Function

Sub CenterForm (frmParent As Form, frmChild As Form)
' This procedure centers a child form over a parent form.
' Calling this routine loads the dialog. Use the Show method
' to display the dialog after calling this routine ( ie MyFrm.Show 1)

Dim l, t
  ' get left offset
  l = frmParent.Left + ((frmParent.Width - frmChild.Width) / 2)
  If (l + frmChild.Width > Screen.Width) Then
    l = Screen.Width = frmChild.Width
  End If

  ' get top offset
  t = frmParent.Top + ((frmParent.Height - frmChild.Height) / 2)
  If (t + frmChild.Height > Screen.Height) Then
    t = Screen.Height - frmChild.Height
  End If

  ' center the child formfv
  frmChild.Move l, t

End Sub

Sub CompileSetup ()
    Screen.MousePointer = 11
    frmSetup.Show modal
End Sub

Sub DefaultFont ()
    On Error Resume Next
    If Fixed Then
        frmMDI.CMDialog3.Flags = &H3& Or &H4000&
    Else
        frmMDI.CMDialog3.Flags = &H3&
    End If
    frmMDI.CMDialog3.FontName = frmMDI!picToolbar.FontName
    frmMDI.CMDialog3.FontSize = frmMDI!picToolbar.FontSize
    frmMDI.CMDialog3.FontBold = frmMDI!picToolbar.FontBold
    frmMDI.CMDialog3.FontItalic = frmMDI!picToolbar.FontItalic
    frmMDI.CMDialog3.Action = 4
    If Err <> 32755 Then 'user has not pressed cancel
        frmMDI!picToolbar.FontName = frmMDI.CMDialog3.FontName
        frmMDI!picToolbar.FontSize = frmMDI.CMDialog3.FontSize
        frmMDI!picToolbar.FontBold = frmMDI.CMDialog3.FontBold
        frmMDI!picToolbar.FontItalic = frmMDI.CMDialog3.FontItalic
    End If
End Sub

Sub EditCopyProc ()
    ' Copy selected text to Clipboard.
    ClipBoard.SetText frmMDI.ActiveForm.ActiveControl.SelText
End Sub

Sub EditCutProc ()
    ' Copy selected text to Clipboard.
    ClipBoard.SetText frmMDI.ActiveForm.ActiveControl.SelText
    ' Delete selected text.
    frmMDI.ActiveForm.ActiveControl.SelText = ""
End Sub

Sub EditPasteProc ()
    ' Place text from Clipboard into active control.
    frmMDI.ActiveForm.ActiveControl.SelText = ClipBoard.GetText()
End Sub

Sub FileNew ()
    Dim fIndex As Integer

    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = "Untitled:" & fIndex
    FState(fIndex).Dirty = False
    Document(fIndex).Show
    
    GetRecentFiles
    ' Make sure toolbar edit buttons are visible
    frmMDI!imgFileSaveButton.Visible = True
    frmMDI!imgFileCompileButton.Visible = True
    frmMDI!imgEditJumpButton.Visible = True
    frmMDI!imgCutButton.Visible = True
    frmMDI!imgCopyButton.Visible = True
    frmMDI!imgPasteButton.Visible = True
    Started = True
End Sub

Function FindFreeIndex () As Integer
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)

    ' Cycle throught the document array. If one of the
    ' documents has been deleted, then return that
    ' index.
    For i = 1 To ArrayCount
        If FState(i).Deleted Then
            FindFreeIndex = i
            FState(i).Deleted = False
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.

    ReDim Preserve Document(ArrayCount + 1)
    ReDim Preserve FState(ArrayCount + 1)
    FindFreeIndex = UBound(Document)
End Function

Sub FindIt (Replace As Integer)
    Dim Start, pos, findstring, sourcestring, msg, response, Offset
    Dim dif As Integer
    
    If (gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart) Then
        Offset = 1
    Else
        Offset = 0
    End If

    If gFirstTime Then Offset = 0

    Start = frmMDI.ActiveForm.ActiveControl.SelStart + Offset
        
    If gFindCase Then
        findstring = gFindString
        sourcestring = frmMDI.ActiveForm.ActiveControl.Text
    Else
        findstring = UCase(gFindString)
        sourcestring = UCase(frmMDI.ActiveForm.ActiveControl.Text)
    End If
            
    If gFindDirection = 1 Then
        pos = InStr(Start + 1, sourcestring, findstring)
    Else
        For pos = Start - 1 To 0 Step -1
            If pos = 0 Then Exit For
            If Mid(sourcestring, pos, Len(findstring)) = findstring Then Exit For
        Next
    End If

    If Replace And Replacing Then
        If gFindDirection = 1 Then dif = Len(frmReplace.Text2) - Len(frmMDI.ActiveForm.ActiveControl.SelText)
        frmMDI.ActiveForm.ActiveControl.SelText = frmReplace.Text2
    End If
    
    ' If string is found
    If pos Then
        frmMDI.ActiveForm.ActiveControl.SelStart = pos - 1 + dif
        frmMDI.ActiveForm.ActiveControl.SelLength = Len(findstring)
        If Replace Then Replacing = True
    Else
        If Replace < 2 Then
            msg = "Cannot find " & Chr(34) & gFindString & Chr(34)
            response = MsgBox(msg, 64, App.Title)
        End If
        If Replace Then Replacing = False
    End If
    
    gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart
    gFirstTime = False

End Sub

Sub GetRecentFiles ()
  Dim retval, key, i, j
  Dim IniString As String

  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)

  ' Get recent file strings from ADI83ASM.INI
  For i = 1 To 4
    key = "RecentFile" & i
    retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "ADI83ASM.INI")
    If retval And Left(IniString, 8) <> "Not Used" Then
      ' Update the MDI form's menu.
      frmMDI.mnuRecentFile(0).Visible = True
      frmMDI.mnuRecentFile(i).Caption = "&" & i & " " & IniString
      frmMDI.mnuRecentFile(i).Visible = True
  
      ' Iterate through all the notepads and update each menu.
      For j = 1 To UBound(Document)
        If Not FState(j).Deleted Then
          Document(j).mnuRecentFile(0).Visible = True
          Document(j).mnuRecentFile(i).Caption = "&" & i & " " & IniString
          Document(j).mnuRecentFile(i).Visible = True
        End If
      Next j
    End If
  Next i

End Sub

Sub JumpLabel ()
    Dim Bounds As String, Char As String, Label As String
    Dim Cursor As Integer, LabelPos As Integer
    Dim Start As Integer, Finish As Integer
    If frmMDI.ActiveForm.Text1.SelLength = 0 Then
        Bounds = " ,:" + Chr(13) + Chr(10)
        Cursor = frmMDI.ActiveForm.Text1.SelStart
        For Start = Cursor To 1 Step -1
            Char = Mid(frmMDI.ActiveForm.Text1, Start, 1)
            If InStr(Bounds, Char) Then Exit For
        Next Start
        For Finish = Cursor + 1 To Len(frmMDI.ActiveForm.Text1)
            Char = Mid(frmMDI.ActiveForm.Text1, Finish, 1)
            If InStr(Bounds, Char) Then Exit For
        Next Finish
        frmMDI.ActiveForm.Text1.SelStart = Start
        frmMDI.ActiveForm.Text1.SelLength = Finish - Start - 1
    End If
    Label = frmMDI.ActiveForm.Text1.SelText + ":"
    If Left(frmMDI.ActiveForm.Text1, Len(Label)) = Label Then
        LabelPos = 0
    Else
        LabelPos = InStr(1, frmMDI.ActiveForm.Text1, Chr(13) + Chr(10) + Label, 1) + 1
        If LabelPos = 1 Then LabelPos = -1
    End If
    If LabelPos > -1 Then
        frmMDI.ActiveForm.Text1.SelStart = LabelPos
    Else
        MsgBox "Cannot find the label " & Chr(34) & Label & Chr(34), 64, "Jump to Label"
    End If
End Sub

Sub LoadSettings ()
    Dim retval As Integer
    Dim ReturnBool As String
    Dim ReturnString As String
    Dim State As Integer
    ReturnBool = " "
    ReturnString = String(255, 0)
    retval = GetPrivateProfileString("ADI TI-83 Assembler", "Toolbar", "1", ReturnBool, 2, "ADI83ASM.INI")
    frmMDI!picToolbar.Visible = -Val(ReturnBool)
    retval = GetPrivateProfileString("ADI TI-83 Assembler", "FixedPitch", "1", ReturnBool, 2, "ADI83ASM.INI")
    Fixed = -Val(ReturnBool)
    retval = GetPrivateProfileString("Compile", "AutoName", "1", ReturnBool, 2, "ADI83ASM.INI")
    frmSetup.chkAutoName = ReturnBool
    retval = GetPrivateProfileString("Compile", "KeepFiles", "0", ReturnBool, 2, "ADI83ASM.INI")
    frmSetup.chkKeep = ReturnBool
    retval = GetPrivateProfileString("Compile", "TASMDir", "", ReturnString, 256, "ADI83ASM.INI")
    frmSetup.txtTASM = Left(ReturnString, retval)
    retval = GetPrivateProfileString("Compile", "83lnkDir", "", ReturnString, 256, "ADI83ASM.INI")
    frmSetup.txt83lnk = Left(ReturnString, retval)
    retval = GetPrivateProfileString("Font", "Name", "Fixedsys", ReturnString, 256, "ADI83ASM.INI")
    frmMDI!picToolbar.FontName = Left(ReturnString, retval)
    retval = GetPrivateProfileString("Font", "Bold", "0", ReturnBool, 2, "ADI83ASM.INI")
    frmMDI!picToolbar.FontBold = -Val(ReturnBool)
    retval = GetPrivateProfileString("Font", "Italic", "0", ReturnBool, 2, "ADI83ASM.INI")
    frmMDI!picToolbar.FontItalic = -Val(ReturnBool)
    retval = GetPrivateProfileString("Font", "Size", "9", ReturnString, 256, "ADI83ASM.INI")
    frmMDI!picToolbar.FontSize = Left(ReturnString, retval)
    retval = GetPrivateProfileString("Window", "Left", "1560", ReturnString, 256, "ADI83ASM.INI")
    frmMDI.Left = Left(ReturnString, retval)
    retval = GetPrivateProfileString("Window", "Top", "1185", ReturnString, 256, "ADI83ASM.INI")
    frmMDI.Top = Left(ReturnString, retval)
    retval = GetPrivateProfileString("Window", "Width", "5805", ReturnString, 256, "ADI83ASM.INI")
    frmMDI.Width = Left(ReturnString, retval)
    retval = GetPrivateProfileString("Window", "Height", "5475", ReturnString, 256, "ADI83ASM.INI")
    frmMDI.Height = Left(ReturnString, retval)
    retval = GetPrivateProfileString("Window", "State", "0", ReturnBool, 2, "ADI83ASM.INI")
    frmMDI.WindowState = 2 * Val(ReturnBool)
    retval = GetPrivateProfileString("Window", "DocumentState", "0", ReturnBool, 2, "ADI83ASM.INI")
    DocState = 2 * Val(ReturnBool)
End Sub

Sub OptionsFixedPitchProc (CurrentForm As Form)
    CurrentForm.mnuOFixedPitch.Checked = Not CurrentForm.mnuOFixedPitch.Checked
    If TypeOf CurrentForm Is MDIForm Then
    Else
        frmMDI.mnuOFixedPitch.Checked = CurrentForm.mnuOFixedPitch.Checked
    End If
    If CurrentForm.mnuOFixedPitch.Checked Then
        Fixed = True
    Else
        Fixed = False
    End If
End Sub

Sub OptionsToolbarProc (CurrentForm As Form)
    CurrentForm.mnuOToolbar.Checked = Not CurrentForm.mnuOToolbar.Checked
    If TypeOf CurrentForm Is MDIForm Then
    Else
        frmMDI.mnuOToolbar.Checked = CurrentForm.mnuOToolbar.Checked
    End If
    If CurrentForm.mnuOToolbar.Checked Then
        frmMDI!picToolbar.Visible = True
    Else
        frmMDI!picToolbar.Visible = False
    End If
End Sub

Sub SaveSettings ()
    Dim retval As Integer
    Dim State As Integer
    Dim OriDir As String
    On Error Resume Next
    retval = WritePrivateProfileString("ADI TI-83 Assembler", "Toolbar", -frmMDI!picToolbar.Visible, "ADI83ASM.INI")
    retval = WritePrivateProfileString("ADI TI-83 Assembler", "FixedPitch", Format(-Fixed), "ADI83ASM.INI")
    retval = WritePrivateProfileString("Compile", "AutoName", frmSetup.chkAutoName, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Compile", "KeepFiles", frmSetup.chkKeep, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Compile", "TASMDir", frmSetup.txtTASM, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Compile", "83lnkDir", frmSetup.txt83lnk, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Font", "Name", frmMDI!picToolbar.FontName, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Font", "Bold", -frmMDI!picToolbar.FontBold, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Font", "Italic", -frmMDI!picToolbar.FontItalic, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Font", "Size", frmMDI!picToolbar.FontSize, "ADI83ASM.INI")
    State = frmMDI.WindowState
    If State = 1 Then State = 0
    frmMDI.WindowState = 0
    retval = WritePrivateProfileString("Window", "Left", frmMDI.Left, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Window", "Top", frmMDI.Top, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Window", "Width", frmMDI.Width, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Window", "Height", frmMDI.Height, "ADI83ASM.INI")
    retval = WritePrivateProfileString("Window", "State", Format(State / 2), "ADI83ASM.INI")
    retval = WritePrivateProfileString("Window", "DocumentState", Format(DocState / 2), "ADI83ASM.INI")
    OriDir = CurDir
    ChDir App.Path
    Kill "~aditemp.bat"
    ChDir OriDir
End Sub

Sub ShiftRecentFiles (Position As Integer)
    Dim i, key, retval
    Dim IniString As String, IniSwap As String
    If Position > 1 Then
        IniString = String(255, 0)
        key = "RecentFile" & Position
        retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "ADI83ASM.INI")
        IniSwap = IniString
        For i = Position To 2 Step -1
            key = "RecentFile" & (i - 1)
            retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "ADI83ASM.INI")
            key = "RecentFile" & i
            retval = WritePrivateProfileString("Recent Files", key, IniString, "ADI83ASM.INI")
        Next i
        retval = WritePrivateProfileString("Recent Files", "RecentFile1", IniSwap, "ADI83ASM.INI")
    End If
End Sub

Sub WriteRecentFiles (OpenFileName)
  Dim i, j, key, retval
  Dim IniString As String
  IniString = String(255, 0)

  ' Copy RecentFile1 to RecentFile2, etc.
  For i = 3 To 1 Step -1
    key = "RecentFile" & i
    retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "ADI83ASM.INI")
    If retval And Left(IniString, 8) <> "Not Used" Then
      key = "RecentFile" & (i + 1)
      retval = WritePrivateProfileString("Recent Files", key, IniString, "ADI83ASM.INI")
    End If
  Next i
  
  ' Write openfile to first Recent File.
    retval = WritePrivateProfileString("Recent Files", "RecentFile1", OpenFileName, "ADI83ASM.INI")

End Sub

