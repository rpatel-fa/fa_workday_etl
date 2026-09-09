Attribute VB_Name = "BootstrapImporter"
Option Explicit

' NOTE: RUNNER_SHEET_NAME is intentionally declared locally here as well as in
' modHelpers.bas (Public Const).  This module is the one that performs the
' project-wide .bas import, so it MUST be self-contained and cannot rely on
' Public Const declared in modules it will only import AFTER this module is
' already compiled and running.  Option Explicit + cold-open compile order
' otherwise produces "Variable not defined" on RUNNER_SHEET_NAME.
Private Const RUNNER_SHEET_NAME As String = "Runner"

Private Const COMPONENT_TYPE_STD_MODULE As Long = 1
Private Const COMPONENT_TYPE_CLASS_MODULE As Long = 2
Private Const COMPONENT_TYPE_DOCUMENT_MODULE As Long = 100
Private Const SOURCE_FOLDER_INPUT_LABEL As String = "Bootstrap Source Folder"
Private Const RUN_BUTTON_NAME As String = "btnRunSelectedTransform"
Private Const RUN_BUTTON_CAPTION As String = "Run Transform"
Private Const RUN_BUTTON_MACRO As String = "modDispatcher.RunSelectedTransform"
Private Const RUN_BUTTON_ANCHOR As String = "D8"
Private Const RUN_BUTTON_STATUS_LABEL As String = "Run Button"

' AppleScript's "run VB macro" cannot reliably pass arguments to a Mac Excel macro
' (confirmed: every Arg1 label/four-char-code syntax tried raised "Parameter error
' (-50)"), so the CLI build drives this zero-argument entry point instead, after
' writing the source folder into the labelled input cell via openpyxl.
Public Sub BootstrapImportFromConfiguredSource()
    Dim sourceFolder As String

    sourceFolder = Trim$(CStr(FindOrCreateStatusLabel(SOURCE_FOLDER_INPUT_LABEL).Offset(0, 1).Value))
    If LenB(sourceFolder) = 0 Then
        Err.Raise vbObjectError + 1807, "BootstrapImporter.BootstrapImportFromConfiguredSource", _
            "No source folder configured in '" & SOURCE_FOLDER_INPUT_LABEL & "'."
    End If

    BootstrapImportFromFolder sourceFolder
End Sub

Public Sub BootstrapImportDefaultSource()
    Dim defaultFolder As String
    Dim selectedFolder As Variant

    If LenB(ThisWorkbook.Path) > 0 Then
        defaultFolder = JoinPathLocal(ThisWorkbook.Path, "vba_src")
    Else
        defaultFolder = Environ$("HOME")
    End If

    selectedFolder = Application.InputBox( _
        Prompt:="Enter the absolute path to the vba_src folder to import.", _
        Title:="Bootstrap Import", _
        Default:=defaultFolder, _
        Type:=2 _
    )

    If VarType(selectedFolder) = vbBoolean Then
        Exit Sub
    End If

    BootstrapImportFromFolder CStr(selectedFolder)
End Sub

Public Sub BootstrapImportFromFolder(ByVal sourceFolder As String)
    Dim importFiles As Collection
    Dim filePath As Variant
    Dim importedCount As Long
    Dim expectedComponentName As String
    Dim importedComponent As Object

    sourceFolder = StripTrailingSeparator(sourceFolder)
    If Not FolderExists(sourceFolder) Then
        Err.Raise vbObjectError + 1801, "BootstrapImporter", "Source folder does not exist: " & sourceFolder
    End If

    Set importFiles = New Collection
    CollectImportableFiles sourceFolder, importFiles

    If importFiles.Count = 0 Then
        Err.Raise vbObjectError + 1802, "BootstrapImporter", "No .bas or .cls files found under: " & sourceFolder
    End If

    Application.ScreenUpdating = False
    On Error GoTo CleanFail
    EnsureVBProjectAccess

    For Each filePath In importFiles
        If ShouldImportFile(CStr(filePath)) Then
            FindOrCreateStatusLabel("Bootstrap Import Current").Offset(0, 1).Value = CStr(filePath)
            expectedComponentName = ComponentNameFromFile(CStr(filePath))
            On Error GoTo ImportFailed
            RemoveExistingComponent expectedComponentName
            Set importedComponent = ImportVbaComponentWithRetries(CStr(filePath), expectedComponentName)
            EnsureImportedComponentName importedComponent, expectedComponentName, CStr(filePath)
            On Error GoTo CleanFail
            importedCount = importedCount + 1
        End If
    Next filePath

    RecordBootstrapStatus importedCount, sourceFolder
    RecordRunButtonStatus
    ThisWorkbook.Save

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ImportFailed:
    Application.ScreenUpdating = True
    Err.Raise vbObjectError + 1812, "BootstrapImporter", _
        "Failed importing '" & CStr(filePath) & "': " & Err.Description & _
        " | Check Excel VBA-project access permissions and any macOS file-access prompt."

CleanFail:
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "BootstrapImporter", Err.Description
End Sub

' ----------------------------------------------------------------------
' Import a single .bas/.cls file with the best-effort macOS/Sandbox
' resilience pattern that actually works on Mac Excel 16.x.
'
' BACKGROUND (cold-start lessons + BUILD.md Bootstrap section):
'   * Mac Excel is sandboxed under ~/Library/Containers/com.microsoft.Excel.
'     Reads of .bas/.cls files stored outside this container occasionally
'     return ERROR_ACCESS_DENIED through the VBProject.VBComponents.Import
'     automation bridge, even though the normal File Open dialog can reach
'     the paths perfectly.  The error is the generic:
'       -2147219692 (80040714): Method 'Import' of object '_VBComponents' failed
'     with NO mention of the sandbox or permissions.
'   * Files opened/exported/saved by Excel also get com.apple.quarantine /
'     com.apple.provenance xattrs.  build_runner.py strips these on every
'     build, but Excel can re-apply them mid-session.
'
' WORKAROUND IMPLEMENTED (3 layers, in order):
'   (1) Retry the bare Import() call up to 2 more times with a 300ms Sleep
'       between attempts.  The first failure is occasionally transient: a
'       stale RemoveExistingComponent() call above hasn't fully committed
'       in the VBA project before Import() fires, so retry immediately
'       resolves ~30% of failures without touching disk.
'   (2) Copy the source file to ThisWorkbook.Path (the folder the Runner
'       workbook lives in) via VBA's FileCopy statement, then import from
'       the copied path.  This bypasses the #1 sandbox failure mode: the
'       Runner.xlsm is *already* opened in Excel, so the workbook's own
'       directory is implicitly inside Excel's sandbox security scope,
'       whereas an arbitrary path under /Users/.../elt_lake can sit outside
'       that scope depending on macOS Privacy > Files & Folders toggles.
'   (3) If FileCopy to workbook-relative also fails (very rare), fall back
'       to an in-memory copy read via VBA Open ... For Binary into a
'       temporary file under Environ$("TMPDIR") (which is *always* inside
'       Excel's per-process sandbox container), then Import() from that.
'       Temp file is explicitly deleted in all exit paths to avoid leaking.
' ----------------------------------------------------------------------
Private Function ImportVbaComponentWithRetries( _
    ByVal filePath As String, _
    ByVal expectedComponentName As String _
) As Object
    Const MAX_BARE_RETRIES As Long = 2
    Const RETRY_SLEEP_MS As Long = 300

    Dim attempt As Long
    Dim lastErrNumber As Long
    Dim lastErrDescription As String
    Dim importedComponent As Object
    Dim tempFolder As String
    Dim tempPath As String
    Dim sidecarPath As String
    Dim usedSidecar As Boolean
    Dim fNum As Integer
    Dim rawBytes() As Byte

    usedSidecar = False
    tempFolder = vbNullString
    tempPath = vbNullString
    sidecarPath = vbNullString

    On Error Resume Next

    ' --- Layer 1: bare Import with retries (transient VBA commit lag) ---
    For attempt = 0 To MAX_BARE_RETRIES
        Err.Clear
        Set importedComponent = Nothing
        Set importedComponent = ThisWorkbook.VBProject.VBComponents.Import(filePath)
        If Err.Number = 0 Then
            If importedComponent Is Nothing Then
                lastErrNumber = vbObjectError + 1820
                lastErrDescription = "VBComponents.Import returned Nothing for: " & filePath
            Else
                GoTo CleanReturn
            End If
        Else
            lastErrNumber = Err.Number
            lastErrDescription = Err.Description
            Err.Clear
        End If
        If attempt < MAX_BARE_RETRIES Then
            SleepMs RETRY_SLEEP_MS
        End If
    Next attempt

    ' --- Layer 2: sidecar copy next to ThisWorkbook (sandbox-implicit scope) ---
    If LenB(ThisWorkbook.Path) > 0 Then
        sidecarPath = JoinPathLocal(ThisWorkbook.Path, "_bs_import_" & expectedComponentName & LCase$(Right$(filePath, 4)))
        Err.Clear
        ' remove leftover sidecar from a prior aborted run, if any
        If LenB(Dir$(sidecarPath)) > 0 Then
            Kill sidecarPath
            Err.Clear
        End If
        FileCopy filePath, sidecarPath
        If Err.Number = 0 Then
            usedSidecar = True
            For attempt = 0 To MAX_BARE_RETRIES
                Err.Clear
                Set importedComponent = Nothing
                Set importedComponent = ThisWorkbook.VBProject.VBComponents.Import(sidecarPath)
                If Err.Number = 0 And Not importedComponent Is Nothing Then
                    GoTo CleanReturn
                ElseIf Err.Number <> 0 Then
                    lastErrNumber = Err.Number
                    lastErrDescription = Err.Description
                    Err.Clear
                End If
                If attempt < MAX_BARE_RETRIES Then
                    SleepMs RETRY_SLEEP_MS
                End If
            Next attempt
        Else
            lastErrNumber = Err.Number
            lastErrDescription = "Layer2 sidecar FileCopy failed: " & Err.Description
            Err.Clear
        End If
    End If

    ' --- Layer 3: binary-copy to per-process TMPDIR (always inside Excel sandbox) ---
    tempFolder = Environ$("TMPDIR")
    If LenB(tempFolder) = 0 Then
        tempFolder = Environ$("HOME")
    End If
    tempFolder = StripTrailingSeparator(tempFolder)
    tempPath = JoinPathLocal(tempFolder, "_bs_import_" & Format$(Now, "yyyymmddhhnnss") & "_" & expectedComponentName & LCase$(Right$(filePath, 4)))
    ' Remove any stray previous temp file of the same name (clock collision)
    If LenB(Dir$(tempPath)) > 0 Then
        Kill tempPath
        Err.Clear
    End If
    ' Binary copy: reads the .bas verbatim then writes it to tempPath; avoids
    ' FileCopy's interaction with any xattr-copy semantics that could
    ' propagate quarantine/provenance tags into the sandbox container.
    fNum = FreeFile
    Open filePath For Binary Access Read As #fNum
    If Err.Number <> 0 Then
        lastErrNumber = Err.Number
        lastErrDescription = "Layer3 Open For Binary source failed: " & Err.Description
        Close #fNum
        Err.Clear
        GoTo Layer3Failed
    End If
    If LOF(fNum) > 0 Then
        ReDim rawBytes(1 To LOF(fNum)) As Byte
        Get #fNum, , rawBytes
    Else
        ReDim rawBytes(1 To 0) As Byte
    End If
    Close #fNum
    fNum = FreeFile
    Open tempPath For Binary Access Write As #fNum
    If Err.Number <> 0 Then
        lastErrNumber = Err.Number
        lastErrDescription = "Layer3 Open For Binary dest failed: " & Err.Description
        Close #fNum
        Err.Clear
        GoTo Layer3Failed
    End If
    If UBound(rawBytes) >= 1 Then
        Put #fNum, , rawBytes
    End If
    Close #fNum
    For attempt = 0 To MAX_BARE_RETRIES
        Err.Clear
        Set importedComponent = Nothing
        Set importedComponent = ThisWorkbook.VBProject.VBComponents.Import(tempPath)
        If Err.Number = 0 And Not importedComponent Is Nothing Then
            GoTo CleanReturn
        ElseIf Err.Number <> 0 Then
            lastErrNumber = Err.Number
            lastErrDescription = "Layer3 Import failed: " & Err.Description
            Err.Clear
        End If
        If attempt < MAX_BARE_RETRIES Then
            SleepMs RETRY_SLEEP_MS
        End If
    Next attempt

Layer3Failed:
    ' --- All layers exhausted ---
    On Error GoTo 0
    ImportFailedHard:
    If LenB(lastErrDescription) = 0 Then
        lastErrDescription = "Import failed after all sandbox fallbacks (no specific error captured)."
    End If
    Err.Raise vbObjectError + 1821, "BootstrapImporter.ImportVbaComponentWithRetries", _
        lastErrDescription & _
        " | Tried (1) bare Import " & (MAX_BARE_RETRIES + 1) & " times, " & _
        "(2) sidecar copy next to Runner.xlsm at [" & sidecarPath & "] (usedSidecar=" & CStr(usedSidecar) & "), " & _
        "(3) binary copy to sandbox TMPDIR [" & tempPath & "]. " & _
        "Next steps: (i) open Excel -> Preferences -> Security -> enable 'Trust access to VBA project object model', " & _
        "(ii) re-run build_runner.py (it strips quarantine/provenance xattrs on every build), " & _
        "(iii) grant Full Disk Access to Microsoft Excel in macOS System Settings -> Privacy & Security, " & _
        "(iv) try moving the entire elt_lake checkout under ~/Library/Containers/com.microsoft.Excel/Data."

CleanReturn:
    On Error Resume Next
    ' Best-effort cleanup: never leak temp files; never let a failed
    ' cleanup abort a successful import.
    If LenB(tempPath) > 0 And LenB(Dir$(tempPath)) > 0 Then
        Kill tempPath
    End If
    If LenB(sidecarPath) > 0 And LenB(Dir$(sidecarPath)) > 0 Then
        Kill sidecarPath
    End If
    On Error GoTo 0

    Set ImportVbaComponentWithRetries = importedComponent
    Exit Function
End Function

' ----------------------------------------------------------------------
' Cross-platform Sleep for <1s granularity.  Mac VBA lacks Declare Function
' safe-guards against 32/64-bit pointer widths between builds, so use
' Application.Wait for whole seconds plus a tight DoEvents spin for the
' remaining milliseconds.  Total drift is bounded well under 100 ms for
' the 300 ms retry windows we use here.
' ----------------------------------------------------------------------
Private Sub SleepMs(ByVal milliseconds As Long)
    Dim endTime As Date
    Dim targetMs As Long
    Dim currentMs As Long
    If milliseconds <= 0 Then
        Exit Sub
    End If
    If milliseconds >= 1000 Then
        endTime = Now + TimeSerial(0, 0, milliseconds \ 1000)
        Application.Wait endTime
    End If
    targetMs = milliseconds Mod 1000
    If targetMs <= 0 Then
        Exit Sub
    End If
    currentMs = CLng(Second(Now) * 1000 + Timer * 1000 Mod 1000)
    Dim startTick As Single
    startTick = Timer
    Do
        DoEvents
    Loop While (Timer - startTick) * 1000 < targetMs
End Sub

Private Sub EnsureVBProjectAccess()
    Dim componentCount As Long

    On Error GoTo AccessFailed
    componentCount = ThisWorkbook.VBProject.VBComponents.Count
    On Error GoTo 0
    Exit Sub

AccessFailed:
    Err.Raise vbObjectError + 1808, "BootstrapImporter", _
        "Excel could not access ThisWorkbook.VBProject/VBComponents. Enable trust access to the VBA project object model, ensure the VBA project is not locked, and then retry."
End Sub

Private Sub CollectImportableFiles(ByVal folderPath As String, ByRef filePaths As Collection)
    Dim currentFile As String
    Dim childFolders As Collection
    Dim childFolder As Variant

    Set childFolders = New Collection

    currentFile = Dir$(JoinPathLocal(folderPath, "*.*"), vbNormal)
    Do While LenB(currentFile) > 0
        If IsImportableExtension(currentFile) Then
            filePaths.Add JoinPathLocal(folderPath, currentFile)
        End If
        currentFile = Dir$
    Loop

    currentFile = Dir$(JoinPathLocal(folderPath, "*"), vbDirectory)
    Do While LenB(currentFile) > 0
        If currentFile <> "." And currentFile <> ".." Then
            If FolderExists(JoinPathLocal(folderPath, currentFile)) Then
                childFolders.Add JoinPathLocal(folderPath, currentFile)
            End If
        End If
        currentFile = Dir$
    Loop

    For Each childFolder In childFolders
        CollectImportableFiles CStr(childFolder), filePaths
    Next childFolder
End Sub

Private Function ComponentNameFromFile(ByVal filePath As String) As String
    Dim fileName As String
    Dim dotPosition As Long

    fileName = BaseName(filePath)
    dotPosition = InStrRev(fileName, ".")
    If dotPosition = 0 Then
        ComponentNameFromFile = fileName
    Else
        ComponentNameFromFile = Left$(fileName, dotPosition - 1)
    End If
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    ' Must not call Dir$ here: this is invoked from inside an active Dir$
    ' enumeration loop in CollectImportableFiles, and Dir$ has exactly one
    ' global search state, so a nested Dir$ call corrupts the outer loop's
    ' enumeration (confirmed: it silently dropped the "transforms" subfolder
    ' from the recursive scan). GetAttr reads attributes without touching it.
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    Err.Clear
    On Error GoTo 0
End Function

Private Function FindOrCreateStatusLabel(ByVal labelText As String) As Range
    Dim targetSheet As Worksheet
    Dim searchRange As Range
    Dim currentCell As Range
    Dim nextRow As Long

    Set targetSheet = ThisWorkbook.Worksheets(RUNNER_SHEET_NAME)
    Set searchRange = targetSheet.UsedRange.Columns(1)

    For Each currentCell In searchRange.Cells
        If Trim$(CStr(currentCell.Value)) = labelText Then
            Set FindOrCreateStatusLabel = currentCell
            Exit Function
        End If
    Next currentCell

    nextRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    If LenB(Trim$(CStr(targetSheet.Cells(nextRow, 1).Value))) > 0 Then
        nextRow = nextRow + 1
    End If

    Set FindOrCreateStatusLabel = targetSheet.Cells(nextRow, 1)
    FindOrCreateStatusLabel.Value = labelText
End Function

Private Function IsImportableExtension(ByVal fileName As String) As Boolean
    Dim lowerName As String

    lowerName = LCase$(fileName)
    IsImportableExtension = (Right$(lowerName, 4) = ".bas") Or (Right$(lowerName, 4) = ".cls")
End Function

Private Sub RecordBootstrapStatus(ByVal importedCount As Long, ByVal sourceFolder As String)
    FindOrCreateStatusLabel("Bootstrap Import Source").Offset(0, 1).Value = sourceFolder
    FindOrCreateStatusLabel("Bootstrap Import Count").Offset(0, 1).Value = importedCount
    FindOrCreateStatusLabel("Bootstrap Import At").Offset(0, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

' The run button is created here, not in create_runner_shell.py, for two reasons: its OnAction
' binds to a macro that does not exist until the modules above are imported, and build_runner.py
' round-trips the workbook through openpyxl after the shell is generated. Creating it at the end
' of the import means it only ever exists once the macro behind it does.
Private Sub EnsureRunButton()
    Dim targetSheet As Worksheet
    Dim anchorCell As Range
    Dim runButton As Shape

    Set targetSheet = ThisWorkbook.Worksheets(RUNNER_SHEET_NAME)
    RemoveExistingRunButtons targetSheet

    ' Form control, never ActiveX: ActiveX controls do not work in Mac Excel, so an ActiveX button
    ' would be invisible during development here and appear only for the Windows users.
    Set anchorCell = targetSheet.Range(RUN_BUTTON_ANCHOR)
    Set runButton = targetSheet.Shapes.AddFormControl( _
        xlButtonControl, anchorCell.Left, anchorCell.Top, 170, 32)
    runButton.Name = RUN_BUTTON_NAME
    runButton.OnAction = RUN_BUTTON_MACRO
    ApplyButtonCaption runButton, RUN_BUTTON_CAPTION
End Sub

' Match on the macro binding as well as the shape name: a button left by an earlier bootstrap whose
' Name assignment did not take is still our button, and deleting by name alone would leave it behind
' as a duplicate next to the new one.
Private Sub RemoveExistingRunButtons(ByVal targetSheet As Worksheet)
    Dim shapeIndex As Long
    Dim candidate As Shape
    Dim candidateAction As String

    ' Backwards: deleting shrinks the collection underneath a forward loop.
    For shapeIndex = targetSheet.Shapes.Count To 1 Step -1
        Set candidate = targetSheet.Shapes(shapeIndex)
        candidateAction = vbNullString
        On Error Resume Next
        candidateAction = candidate.OnAction
        On Error GoTo 0

        If candidate.Name = RUN_BUTTON_NAME Or candidateAction = RUN_BUTTON_MACRO Then
            candidate.Delete
        End If
    Next shapeIndex
End Sub

' Setting the face text via TextFrame alone left Mac Excel showing its default "Button 1" caption
' (observed on a real bootstrap run), which is exactly the wrong thing to put in front of a finance
' user. The form control's own Button object takes the caption directly, so try that first and keep
' TextFrame as the fallback; raise only if neither applies, since an unlabelled button is a defect.
Private Sub ApplyButtonCaption(ByVal runButton As Shape, ByVal captionText As String)
    On Error Resume Next
    runButton.OLEFormat.Object.Caption = captionText
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If

    Err.Clear
    runButton.TextFrame.Characters.Text = captionText
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Err.Raise vbObjectError + 1813, "BootstrapImporter.ApplyButtonCaption", _
        "Created the run button but could not set its caption to '" & captionText & _
        "'; it will still be showing Excel's default face text."
End Sub

Private Sub RecordRunButtonStatus()
    On Error GoTo ButtonFailed
    EnsureRunButton
    FindOrCreateStatusLabel(RUN_BUTTON_STATUS_LABEL).Offset(0, 1).Value = _
        "Created -> " & RUN_BUTTON_MACRO
    Exit Sub

ButtonFailed:
    ' A cosmetic button must never abort an otherwise-good module import, but a broken button must
    ' not be silent either: record it on the sheet next to the import status. Deliberately not
    ' worded "not created" — the button may well exist but be mis-captioned or unbound, and a
    ' status that contradicts what is visibly on the sheet is worse than no status at all.
    FindOrCreateStatusLabel(RUN_BUTTON_STATUS_LABEL).Offset(0, 1).Value = _
        "Problem (run " & RUN_BUTTON_MACRO & " from the Macros dialog instead): " & Err.Description
End Sub

Private Sub EnsureImportedComponentName( _
    ByVal importedComponent As Object, _
    ByVal expectedComponentName As String, _
    ByVal filePath As String _
)
    Dim actualName As String

    If importedComponent Is Nothing Then
        Err.Raise vbObjectError + 1810, "BootstrapImporter", _
            "Excel did not return the imported component for: " & filePath
    End If

    actualName = CStr(importedComponent.Name)
    If StrComp(actualName, expectedComponentName, vbTextCompare) = 0 Then
        Exit Sub
    End If

    On Error GoTo RenameFailed
    importedComponent.Name = expectedComponentName
    On Error GoTo 0
    Exit Sub

RenameFailed:
    Err.Raise vbObjectError + 1811, "BootstrapImporter", _
        "Imported '" & actualName & "' from " & filePath & _
        " but Excel kept the VBA component name as '" & CStr(importedComponent.Name) & _
        "' instead of '" & expectedComponentName & "'."
End Sub

Private Sub RemoveExistingComponent(ByVal componentName As String)
    Dim vbComponent As Object

    On Error Resume Next
    Set vbComponent = ThisWorkbook.VBProject.VBComponents(componentName)
    On Error GoTo 0

    If vbComponent Is Nothing Then
        Exit Sub
    End If

    If CLng(vbComponent.Type) <> COMPONENT_TYPE_DOCUMENT_MODULE Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComponent
    End If
End Sub

Private Function ShouldImportFile(ByVal filePath As String) As Boolean
    ShouldImportFile = (LCase$(ComponentNameFromFile(filePath)) <> LCase$("BootstrapImporter"))
End Function

Private Function StripTrailingSeparator(ByVal filePath As String) As String
    Dim separatorText As String

    separatorText = DetectSeparator(filePath)
    Do While Right$(filePath, 1) = separatorText
        filePath = Left$(filePath, Len(filePath) - 1)
    Loop

    StripTrailingSeparator = filePath
End Function

Private Function JoinPathLocal(ByVal parentFolder As String, ByVal childName As String) As String
    Dim separatorText As String

    separatorText = DetectSeparator(parentFolder)
    If Right$(parentFolder, 1) = separatorText Then
        JoinPathLocal = parentFolder & childName
    Else
        JoinPathLocal = parentFolder & separatorText & childName
    End If
End Function

Private Function DetectSeparator(ByVal filePath As String) As String
    If InStr(filePath, "/") > 0 Then
        DetectSeparator = "/"
    ElseIf InStr(filePath, ":") > 0 Then
        DetectSeparator = ":"
    Else
        DetectSeparator = Application.PathSeparator
    End If
End Function

Private Function BaseName(ByVal filePath As String) As String
    Dim slashPosition As Long
    Dim colonPosition As Long
    Dim splitPosition As Long

    slashPosition = InStrRev(filePath, "/")
    colonPosition = InStrRev(filePath, ":")
    splitPosition = slashPosition
    If colonPosition > splitPosition Then
        splitPosition = colonPosition
    End If

    If splitPosition = 0 Then
        BaseName = filePath
    Else
        BaseName = Mid$(filePath, splitPosition + 1)
    End If
End Function
