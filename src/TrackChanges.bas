Attribute VB_Name = "TrackChanges"
'==============================================================================
' Excel Track Changes  v2.0
' https://github.com/misht-world/excel-track-changes-macro
'
' How it works
'   * "Start tracking" copies the values of every sheet into a very hidden
'     snapshot sheet (_tc_1, _tc_2, ...).
'   * Every tracked sheet gets a hidden sheet-level name _TC_Diff that compares
'     a cell with the same cell of its snapshot, and one conditional formatting
'     rule "=_TC_Diff" over the whole sheet.
'   * From then on no VBA code runs while you edit, so Undo (Ctrl+Z) keeps
'     working, nothing slows down, and your own formatting is never touched:
'     the highlight is drawn on top of it and disappears when the value
'     returns to the original.
'
' Only the buttons below change the workbook (and, as any macro, they clear
' the Undo history once when you press them).
'==============================================================================
Option Explicit

Private Const REG_APP As String = "ExcelTrackChanges"
Private Const REG_SEC As String = "Settings"
Private Const HELP_URL As String = "https://github.com/misht-world/excel-track-changes-macro"

Private Const NM_DIFF As String = "_TC_Diff"            ' sheet-level: comparison formula
Private Const NM_SNAP As String = "_TC_Snapshot"        ' sheet-level: name of the snapshot sheet
Private Const NM_IS_SNAP As String = "_TC_IsSnapshot"   ' sheet-level marker on snapshot sheets
Private Const SNAP_PREFIX As String = "_tc_"
Private Const RULE_FORMULA As String = "=_TC_Diff"

Private Const STYLE_FONT As Long = 0
Private Const STYLE_FILL As Long = 1
Private Const STYLE_BOTH As Long = 2
Private Const COLOR_COUNT As Long = 5

Private mRibbon As Object
Private mScreen As Boolean
Private mEvents As Boolean
Private mActive As Object
Private mSilent As Boolean
Private mLastMessage As String

'==============================================================================
' Commands (also available through Alt+F8 by typing the name)
'==============================================================================

Public Sub TC_Start()
    Dim wb As Workbook
    If Not GetWorkbook(wb) Then Exit Sub

    If IsTracking(wb) Then
        If Box(("Tracking is already on in this workbook." & vbCrLf & vbCrLf & _
                    "Accept the current state as the new original?"), _
                  vbQuestion + vbYesNo, AppTitle()) = vbYes Then
            AcceptAll wb
        End If
        Exit Sub
    End If

    If wb.ProtectStructure Then
        Box "The workbook structure is protected. Unprotect it (Review > Protect Workbook) and try again.", _
               vbExclamation, AppTitle()
        Exit Sub
    End If

    Dim skipped As String
    BeginWork wb
    On Error GoTo Fail
    skipped = TrackAllSheets(wb)
    EndWork

    Dim msg As String
    msg = ("Tracking is on. Changed cells are highlighted, Undo (Ctrl+Z) works as usual." & vbCrLf & _
            "The original values are stored inside the workbook, so tracking continues after you save and reopen it.")
    If Len(skipped) > 0 Then msg = msg & vbCrLf & vbCrLf & SkippedText(skipped)
    Box msg, vbInformation, AppTitle()
    InvalidateRibbon
    Exit Sub
Fail:
    Dim e As String: e = Err.Description
    EndWork
    ShowError e
End Sub

Public Sub TC_AcceptAll()
    Dim wb As Workbook
    If Not GetWorkbook(wb) Then Exit Sub
    If Not RequireTracking(wb) Then Exit Sub
    If Box(("Accept all changes?" & vbCrLf & vbCrLf & _
                "The current values become the new original and all highlighting disappears. This cannot be undone."), _
              vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Sub
    AcceptAll wb
End Sub

Public Sub TC_AcceptSelection()
    Dim wb As Workbook
    If Not GetWorkbook(wb) Then Exit Sub
    If Not RequireTracking(wb) Then Exit Sub

    If TypeName(Selection) <> "Range" Or TypeName(ActiveSheet) <> "Worksheet" Then
        Box "Select the cells whose changes you want to accept.", vbInformation, AppTitle()
        Exit Sub
    End If

    Dim ws As Worksheet, snap As Worksheet
    Set ws = ActiveSheet
    Set snap = GetSnapshotSheet(ws)
    If snap Is Nothing Then
        Box "This sheet is not tracked yet. Press ""Refresh"" to start tracking it.", _
               vbInformation, AppTitle()
        Exit Sub
    End If

    Dim sel As Range
    Set sel = Selection
    BeginWork wb
    On Error GoTo Fail
    Dim area As Range, part As Range
    For Each area In sel.Areas
        ' clear what the snapshot had there, then copy the current values
        Set part = Application.Intersect(snap.Range(area.Address), snap.UsedRange)
        If Not part Is Nothing Then part.Clear
        Set part = Application.Intersect(area, ws.UsedRange)
        If Not part Is Nothing Then CopyValues part, snap.Range(part.Address)
    Next area
    EndWork
    Flash "Changes in the selected cells are accepted."
    Exit Sub
Fail:
    Dim e As String: e = Err.Description
    EndWork
    ShowError e
End Sub

Public Sub TC_Refresh()
    ' Rebuilds the highlighting rules (e.g. after pasting cells with their
    ' formatting) and starts tracking sheets added after tracking was turned on.
    Dim wb As Workbook
    If Not GetWorkbook(wb) Then Exit Sub
    If Not RequireTracking(wb) Then Exit Sub

    Dim ws As Worksheet, added As String, skipped As String
    BeginWork wb
    On Error GoTo Fail
    For Each ws In wb.Worksheets
        If Not IsSnapshotSheet(ws) Then
            If ws.ProtectContents Then
                skipped = skipped & vbCrLf & "  - " & ws.Name
            ElseIf GetSnapshotSheet(ws) Is Nothing Then
                TrackSheet ws
                added = added & vbCrLf & "  - " & ws.Name
            Else
                SetDiffName ws, GetSnapshotSheet(ws)
                ApplyRule ws
            End If
        End If
    Next ws
    EndWork

    If Len(added) > 0 Or Len(skipped) > 0 Then
        Dim msg As String
        msg = "Highlighting is refreshed."
        If Len(added) > 0 Then
            msg = msg & vbCrLf & vbCrLf & _
                  "Now also tracking (current values are taken as the original):" & added
        End If
        If Len(skipped) > 0 Then msg = msg & vbCrLf & vbCrLf & SkippedText(skipped)
        Box msg, vbInformation, AppTitle()
    Else
        Flash "Highlighting is refreshed."
    End If
    Exit Sub
Fail:
    Dim e As String: e = Err.Description
    EndWork
    ShowError e
End Sub

Public Sub TC_Stop()
    Dim wb As Workbook
    If Not GetWorkbook(wb) Then Exit Sub
    If Not IsTracking(wb) Then
        If HasLeftovers(wb) Then
            RemoveLeftovers wb
        Else
            Box "Tracking is not on in this workbook.", _
                   vbInformation, AppTitle()
        End If
        Exit Sub
    End If
    If Box(("Turn off tracking?" & vbCrLf & vbCrLf & _
                "The highlighting and the stored original values will be removed. Your data and formatting stay as they are."), _
              vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Sub
    If wb.ProtectStructure Then
        Box "The workbook structure is protected. Unprotect it and try again.", vbExclamation, AppTitle()
        Exit Sub
    End If

    Dim ws As Worksheet, snaps As New Collection, skipped As String, i As Long
    BeginWork wb
    On Error GoTo Fail
    For Each ws In wb.Worksheets
        If IsSnapshotSheet(ws) Then
            snaps.Add ws
        ElseIf ws.ProtectContents Then
            skipped = skipped & vbCrLf & "  - " & ws.Name
        Else
            UntrackSheet ws
        End If
    Next ws
    If Len(skipped) = 0 Then
        Application.DisplayAlerts = False
        For i = 1 To snaps.Count
            snaps(i).Visible = xlSheetHidden
            snaps(i).Delete
        Next i
        Application.DisplayAlerts = True
    End If
    EndWork
    InvalidateRibbon

    If Len(skipped) > 0 Then
        Box "Tracking is turned off on all sheets except the protected ones below. " & _
                 "Unprotect them and press ""Turn off"" again to remove the stored original values." & skipped, _
               vbExclamation, AppTitle()
    Else
        Flash "Tracking is turned off."
    End If
    Exit Sub
Fail:
    Dim e As String: e = Err.Description
    EndWork
    ShowError e
End Sub

Private Function HasLeftovers(ByVal wb As Workbook) As Boolean
    ' Rules and names that came along when cells were copied from a tracked
    ' workbook (they link back to that workbook).
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not FindSheetLevelName(ws, NM_DIFF) Is Nothing Or HasRule(ws) Then
            HasLeftovers = True
            Exit Function
        End If
    Next ws
End Function

Private Sub RemoveLeftovers(ByVal wb As Workbook)
    If Box(("This workbook is not tracked, but it contains highlighting copied from a tracked workbook (it links back to that file)." & _
             vbCrLf & vbCrLf & "Remove it?"), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Sub

    Dim ws As Worksheet, skipped As String
    BeginWork wb
    On Error GoTo Fail
    For Each ws In wb.Worksheets
        If ws.ProtectContents Then
            skipped = skipped & vbCrLf & "  - " & ws.Name
        Else
            UntrackSheet ws
        End If
    Next ws
    EndWork
    If Len(skipped) > 0 Then
        Box "Removed everywhere except the protected sheets:" & skipped, _
            vbExclamation, AppTitle()
    Else
        Flash "Copied highlighting is removed."
    End If
    Exit Sub
Fail:
    Dim e As String: e = Err.Description
    EndWork
    ShowError e
End Sub

'==============================================================================
' Highlight style. Stored per user; applied to the active workbook at once.
'==============================================================================

Public Sub TC_SetStyle(ByVal styleIndex As Long, ByVal colorIndex As Long)
    If styleIndex >= 0 And styleIndex <= STYLE_BOTH Then SaveSetting REG_APP, REG_SEC, "Style", CStr(styleIndex)
    If colorIndex >= 0 And colorIndex < COLOR_COUNT Then SaveSetting REG_APP, REG_SEC, "Color", CStr(colorIndex)
    ApplyStyleToActiveWorkbook
    InvalidateRibbon
End Sub

Private Sub ApplyStyleToActiveWorkbook()
    Dim wb As Workbook
    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then Exit Sub
    If Not IsTracking(wb) Then Exit Sub

    Dim ws As Worksheet
    BeginWork wb
    On Error GoTo Fail
    For Each ws In wb.Worksheets
        If Not IsSnapshotSheet(ws) And Not ws.ProtectContents Then
            If Not GetSnapshotSheet(ws) Is Nothing Then ApplyRule ws
        End If
    Next ws
    EndWork
    Exit Sub
Fail:
    Dim e As String: e = Err.Description
    EndWork
    ShowError e
End Sub

Private Function GetStyle() As Long
    GetStyle = CLng(Val(GetSetting(REG_APP, REG_SEC, "Style", "0")))
    If GetStyle < 0 Or GetStyle > STYLE_BOTH Then GetStyle = STYLE_FONT
End Function

Private Function GetColorIndex() As Long
    GetColorIndex = CLng(Val(GetSetting(REG_APP, REG_SEC, "Color", "0")))
    If GetColorIndex < 0 Or GetColorIndex >= COLOR_COUNT Then GetColorIndex = 0
End Function

Private Function FontColor(ByVal idx As Long) As Long
    Select Case idx
        Case 1: FontColor = RGB(192, 0, 0)      ' red
        Case 2: FontColor = RGB(0, 128, 0)      ' green
        Case 3: FontColor = RGB(197, 90, 17)    ' orange
        Case 4: FontColor = RGB(112, 48, 160)   ' purple
        Case Else: FontColor = RGB(0, 0, 255)   ' blue
    End Select
End Function

Private Function FillColor(ByVal idx As Long) As Long
    Select Case idx
        Case 1: FillColor = RGB(255, 199, 206)
        Case 2: FillColor = RGB(198, 239, 206)
        Case 3: FillColor = RGB(255, 230, 153)
        Case 4: FillColor = RGB(225, 211, 242)
        Case Else: FillColor = RGB(189, 215, 238)
    End Select
End Function

Private Function StyleLabel(ByVal idx As Long) As String
    Select Case idx
        Case STYLE_FILL: StyleLabel = "Cell fill"
        Case STYLE_BOTH: StyleLabel = "Text and fill"
        Case Else: StyleLabel = "Text color"
    End Select
End Function

Private Function ColorLabel(ByVal idx As Long) As String
    Select Case idx
        Case 1: ColorLabel = "Red"
        Case 2: ColorLabel = "Green"
        Case 3: ColorLabel = "Orange"
        Case 4: ColorLabel = "Purple"
        Case Else: ColorLabel = "Blue"
    End Select
End Function

'==============================================================================
' Tracking internals
'==============================================================================

Private Function TrackAllSheets(ByVal wb As Workbook) As String
    ' Returns the list of skipped (protected) sheets.
    Dim ws As Worksheet, skipped As String
    For Each ws In wb.Worksheets
        If Not IsSnapshotSheet(ws) Then
            If ws.ProtectContents Then
                skipped = skipped & vbCrLf & "  - " & ws.Name
            Else
                TrackSheet ws
            End If
        End If
    Next ws
    TrackAllSheets = skipped
End Function

Private Sub AcceptAll(ByVal wb As Workbook)
    Dim skipped As String
    BeginWork wb
    On Error GoTo Fail
    skipped = TrackAllSheets(wb)
    EndWork
    If Len(skipped) > 0 Then
        Box "Changes are accepted." & vbCrLf & vbCrLf & SkippedText(skipped), _
               vbInformation, AppTitle()
    Else
        Flash "All changes are accepted."
    End If
    Exit Sub
Fail:
    Dim e As String: e = Err.Description
    EndWork
    ShowError e
End Sub

Private Sub TrackSheet(ByVal ws As Worksheet)
    ' Takes a fresh snapshot of the sheet and (re)applies the highlight rule.
    Dim snap As Worksheet
    Set snap = GetSnapshotSheet(ws)
    If snap Is Nothing Then
        Set snap = CreateSnapshotSheet(ws.Parent)
        SetSheetLevelName ws, NM_SNAP, "=""" & snap.Name & """"
    End If

    snap.Cells.Clear
    Dim used As Range
    Set used = ws.UsedRange
    CopyValues used, snap.Range(used.Address)

    SetDiffName ws, snap
    ApplyRule ws
End Sub

Private Sub UntrackSheet(ByVal ws As Worksheet)
    RemoveRules ws
    DeleteSheetLevelName ws, NM_DIFF
    DeleteSheetLevelName ws, NM_SNAP
End Sub

Private Function CreateSnapshotSheet(ByVal wb As Workbook) As Worksheet
    Dim sh As Worksheet
    Set sh = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = FreeSnapshotName(wb)
    SetSheetLevelName sh, NM_IS_SNAP, "=TRUE"
    sh.Visible = xlSheetVeryHidden
    Set CreateSnapshotSheet = sh
End Function

Private Function FreeSnapshotName(ByVal wb As Workbook) As String
    Dim i As Long
    Do
        i = i + 1
    Loop While SheetExists(wb, SNAP_PREFIX & i)
    FreeSnapshotName = SNAP_PREFIX & i
End Function

Private Sub CopyValues(ByVal src As Range, ByVal dst As Range)
    ' Copies values exactly: text stays text ("001", "=abc"), empty strings
    ' returned by formulas stay non-blank, errors stay errors.
    ' Does not use the clipboard, so filtered rows are copied too.
    Dim v As Variant
    dst.NumberFormat = "@"
    v = src.Value2

    If Not IsArray(v) Then
        If VarType(v) = vbString Then
            If Len(v) = 0 Then
                SetEmptyText dst
                Exit Sub
            End If
        End If
        dst.Value2 = v
        Exit Sub
    End If

    dst.Value2 = v

    Dim r As Long, c As Long, pending As Range, n As Long
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            If VarType(v(r, c)) = vbString Then
                If Len(v(r, c)) = 0 Then
                    If pending Is Nothing Then
                        Set pending = dst.Cells(r, c)
                    Else
                        Set pending = Application.Union(pending, dst.Cells(r, c))
                    End If
                    n = n + 1
                    If n >= 500 Then
                        SetEmptyText pending
                        Set pending = Nothing
                        n = 0
                    End If
                End If
            End If
        Next c
    Next r
    If Not pending Is Nothing Then SetEmptyText pending
End Sub

Private Sub SetEmptyText(ByVal rng As Range)
    ' A cell that holds an empty string (not a blank cell): formula ="".
    ' Must not be text-formatted, otherwise the formula is stored as text.
    Dim a As Range
    For Each a In rng.Areas
        a.NumberFormat = "General"
        a.Formula = "="""""
    Next a
End Sub

Private Sub SetDiffName(ByVal ws As Worksheet, ByVal snap As Worksheet)
    ' Relative R1C1 formula: compares the cell being drawn (RC) with the same
    ' cell of the snapshot. Type-aware (blank <> 0 <> ""), case-sensitive,
    ' and compares error values by their kind.
    Dim q As String, f As String
    q = "'" & Replace(snap.Name, "'", "''") & "'!RC"
    f = "=IF(TYPE(RC)<>TYPE(" & q & "),TRUE," & _
        "IF(TYPE(RC)=16,ERROR.TYPE(RC)<>ERROR.TYPE(" & q & ")," & _
        "NOT(EXACT(RC," & q & "))))"

    Dim n As Name
    Set n = FindSheetLevelName(ws, NM_DIFF)
    If n Is Nothing Then Set n = ws.Names.Add(Name:=NM_DIFF, RefersTo:="=FALSE", Visible:=False)
    n.RefersToR1C1 = f
End Sub

Private Sub ApplyRule(ByVal ws As Worksheet)
    RemoveRules ws

    Dim fc As FormatCondition, st As Long, ci As Long
    Set fc = ws.Cells.FormatConditions.Add(Type:=xlExpression, Formula1:=RULE_FORMULA)
    st = GetStyle()
    ci = GetColorIndex()
    If st = STYLE_FONT Or st = STYLE_BOTH Then fc.Font.Color = FontColor(ci)
    If st = STYLE_FILL Or st = STYLE_BOTH Then fc.Interior.Color = FillColor(ci)
    fc.StopIfTrue = False
    fc.SetFirstPriority
End Sub

Private Function HasRule(ByVal ws As Worksheet) As Boolean
    Dim i As Long, f As String
    For i = 1 To ws.Cells.FormatConditions.Count
        f = vbNullString
        On Error Resume Next
        f = ws.Cells.FormatConditions(i).Formula1
        On Error GoTo 0
        If StrComp(f, RULE_FORMULA, vbTextCompare) = 0 Then
            HasRule = True
            Exit Function
        End If
    Next i
End Function

Private Sub RemoveRules(ByVal ws As Worksheet)
    ' Removes only our rules (and their fragments left after pasting cells).
    Dim i As Long, fc As Object, f As String
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        Set fc = ws.Cells.FormatConditions(i)
        f = vbNullString
        On Error Resume Next
        f = fc.Formula1
        On Error GoTo 0
        If StrComp(f, RULE_FORMULA, vbTextCompare) = 0 Then fc.Delete
    Next i
End Sub

'==============================================================================
' Lookup helpers
'==============================================================================

Private Function IsTracking(ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not FindSheetLevelName(ws, NM_SNAP) Is Nothing Then
            IsTracking = True
            Exit Function
        End If
    Next ws
End Function

Private Function IsSnapshotSheet(ByVal ws As Worksheet) As Boolean
    IsSnapshotSheet = Not FindSheetLevelName(ws, NM_IS_SNAP) Is Nothing
End Function

Private Function GetSnapshotSheet(ByVal ws As Worksheet) As Worksheet
    Dim n As Name, s As String
    Set n = FindSheetLevelName(ws, NM_SNAP)
    If n Is Nothing Then Exit Function
    s = n.RefersTo                                  ' ="_tc_1"
    If Left$(s, 2) = "=""" And Right$(s, 1) = """" Then s = Mid$(s, 3, Len(s) - 3)
    If Not SheetExists(ws.Parent, s) Then Exit Function
    If TypeName(ws.Parent.Sheets(s)) <> "Worksheet" Then Exit Function
    If Not IsSnapshotSheet(ws.Parent.Sheets(s)) Then Exit Function
    Set GetSnapshotSheet = ws.Parent.Sheets(s)
End Function

Private Function FindSheetLevelName(ByVal ws As Worksheet, ByVal nm As String) As Name
    Dim n As Name, full As String
    For Each n In ws.Names
        full = n.Name                               ' Sheet1!_TC_Diff
        If StrComp(Mid$(full, InStrRev(full, "!") + 1), nm, vbTextCompare) = 0 Then
            If TypeName(n.Parent) = "Worksheet" Then
                If n.Parent Is ws Then
                    Set FindSheetLevelName = n
                    Exit Function
                End If
            End If
        End If
    Next n
End Function

Private Sub SetSheetLevelName(ByVal ws As Worksheet, ByVal nm As String, ByVal refersTo As String)
    Dim n As Name
    Set n = FindSheetLevelName(ws, nm)
    If n Is Nothing Then
        ws.Names.Add Name:=nm, RefersTo:=refersTo, Visible:=False
    Else
        n.RefersTo = refersTo
    End If
End Sub

Private Sub DeleteSheetLevelName(ByVal ws As Worksheet, ByVal nm As String)
    Dim n As Name
    Set n = FindSheetLevelName(ws, nm)
    If Not n Is Nothing Then n.Delete
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = wb.Sheets(sheetName)
    SheetExists = Not o Is Nothing
End Function

'==============================================================================
' UI helpers
'==============================================================================

Private Function GetWorkbook(ByRef wb As Workbook) As Boolean
    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then
        Box "Open a workbook first. Files in Protected View must be enabled for editing.", _
               vbInformation, AppTitle()
        Exit Function
    End If
    GetWorkbook = True
End Function

Private Function RequireTracking(ByVal wb As Workbook) As Boolean
    If IsTracking(wb) Then
        RequireTracking = True
    Else
        If Box("Tracking is not on in this workbook. Turn it on now?", _
                  vbQuestion + vbYesNo, AppTitle()) = vbYes Then TC_Start
    End If
End Function

Private Sub BeginWork(ByVal wb As Workbook)
    mScreen = Application.ScreenUpdating
    mEvents = Application.EnableEvents
    Set mActive = Nothing
    On Error Resume Next
    Set mActive = wb.ActiveSheet
    On Error GoTo 0
    Application.ScreenUpdating = False
    Application.EnableEvents = False     ' do not trigger the workbook's own macros
    Application.Cursor = xlWait
End Sub

Private Sub EndWork()
    On Error Resume Next
    If Not mActive Is Nothing Then mActive.Activate
    Set mActive = Nothing
    Application.DisplayAlerts = True
    Application.Cursor = xlDefault
    Application.EnableEvents = mEvents
    Application.ScreenUpdating = mScreen
End Sub

Private Sub ShowError(ByVal description As String)
    Box "Something went wrong:" & vbCrLf & description, vbExclamation, AppTitle()
End Sub

Private Function SkippedText(ByVal list As String) As String
    SkippedText = "Protected sheets were skipped (unprotect them and press ""Refresh""):" & list
End Function

Private Sub Flash(ByVal message As String)
    ' Short status bar message that clears itself after a few seconds.
    Application.StatusBar = AppTitle() & ": " & message
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 4), "'" & ThisWorkbook.Name & "'!TC_ClearStatus"
End Sub

Public Sub TC_ClearStatus()
    Application.StatusBar = False
End Sub

Private Function AppTitle() As String
    AppTitle = "Track Changes"
End Function

'==============================================================================
' Automation / tests: in silent mode questions are answered "Yes" and no
' message boxes are shown. TC_LastMessage returns the text of the last one.
'==============================================================================

Public Sub TC_SetSilent(ByVal silent As Boolean)
    mSilent = silent
End Sub

Public Function TC_LastMessage() As String
    TC_LastMessage = mLastMessage
End Function

Private Function Box(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle, ByVal title As String) As VbMsgBoxResult
    mLastMessage = prompt
    If mSilent Then
        If (buttons And vbYesNo) = vbYesNo Then Box = vbYes Else Box = vbOK
    Else
        Box = MsgBox(prompt, buttons, title)
    End If
End Function

'==============================================================================
' Ribbon callbacks (customUI14.xml)
'==============================================================================

Public Sub TC_RibbonLoad(ByVal ribbon As Object)
    Set mRibbon = ribbon
End Sub

Private Sub InvalidateRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Public Sub TC_OnAction(ByVal control As Object)
    Select Case control.ID
        Case "tcStart": TC_Start
        Case "tcAcceptAll": TC_AcceptAll
        Case "tcAcceptSel": TC_AcceptSelection
        Case "tcRefresh": TC_Refresh
        Case "tcStop": TC_Stop
        Case "tcHelp": ThisWorkbook.FollowHyperlink HELP_URL
    End Select
End Sub

Public Sub TC_GetLabel(ByVal control As Object, ByRef returnedVal)
    Select Case control.ID
        Case "tcTab": returnedVal = "Track Changes"
        Case "tcGroupTrack": returnedVal = "Tracking"
        Case "tcGroupStyle": returnedVal = "Highlight"
        Case "tcGroupHelp": returnedVal = "Help"
        Case "tcStart": returnedVal = "Start tracking"
        Case "tcAcceptAll": returnedVal = "Accept all"
        Case "tcAcceptSel": returnedVal = "Accept selection"
        Case "tcRefresh": returnedVal = "Refresh"
        Case "tcStop": returnedVal = "Turn off"
        Case "tcStyle": returnedVal = "Style"
        Case "tcColor": returnedVal = "Color"
        Case "tcHelp": returnedVal = "About"
    End Select
End Sub

Public Sub TC_GetScreentip(ByVal control As Object, ByRef returnedVal)
    Select Case control.ID
        Case "tcStart"
            returnedVal = "Remember the current values of all sheets and highlight every cell that differs from them."
        Case "tcAcceptAll"
            returnedVal = "Take the current values as the new original. All highlighting disappears."
        Case "tcAcceptSel"
            returnedVal = "Accept changes only in the selected cells."
        Case "tcRefresh"
            returnedVal = "Restore the highlighting after pasting cells with formatting, and start tracking newly added sheets."
        Case "tcStop"
            returnedVal = "Remove the highlighting and the stored original values. Your data and formatting are not changed."
        Case "tcStyle"
            returnedVal = "How to highlight changed cells."
        Case "tcColor"
            returnedVal = "Highlight color."
        Case "tcHelp"
            returnedVal = "Open the project page."
    End Select
End Sub

Public Sub TC_GetItemCount(ByVal control As Object, ByRef returnedVal)
    If control.ID = "tcStyle" Then returnedVal = STYLE_BOTH + 1 Else returnedVal = COLOR_COUNT
End Sub

Public Sub TC_GetItemLabel(ByVal control As Object, ByVal index As Integer, ByRef returnedVal)
    If control.ID = "tcStyle" Then returnedVal = StyleLabel(index) Else returnedVal = ColorLabel(index)
End Sub

Public Sub TC_GetSelectedItemIndex(ByVal control As Object, ByRef returnedVal)
    If control.ID = "tcStyle" Then returnedVal = GetStyle() Else returnedVal = GetColorIndex()
End Sub

Public Sub TC_OnSelect(ByVal control As Object, ByVal selectedId As String, ByVal selectedIndex As Integer)
    If control.ID = "tcStyle" Then
        TC_SetStyle selectedIndex, -1
    Else
        TC_SetStyle -1, selectedIndex
    End If
End Sub
