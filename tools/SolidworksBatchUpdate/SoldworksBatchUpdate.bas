Option Explicit

Dim swApp As SldWorks.SldWorks

Private Type BatchOptions
    sourceFolder As String
    exportFolder As String
    modelInitials As String
    drawingInitials As String
    skip003 As Boolean
    doPropertyUpdate As Boolean
    doDXFExport As Boolean
End Type

Private Type PropertyBatchResult
    processed As Long
    updated As Long
    skipped As Long
    failed As Long
    skipped003 As Long
    logPath As String
End Type

Private Type DXFBatchResult
    processed As Long
    exported As Long
    skipped As Long
    failed As Long
    logPath As String
End Type

Private Type PartAnalysis
    HasSheetMetal       As Boolean
    HasBends            As Boolean
    HasRolledBack       As Boolean
    flatPatternName     As String
    MultipleConfigs     As Boolean
End Type

Private Enum PartResult
    prExported = 0
    prFailed = 1
    prSkipped = 2
End Enum

Sub main()

    Set swApp = Application.SldWorks

    Dim opts As BatchOptions
    If Not ShowCombinedDialog(opts) Then Exit Sub

    opts.sourceFolder = NormalizeFolder(opts.sourceFolder)
    opts.exportFolder = NormalizeFolder(opts.exportFolder)

    If Not FolderExists(opts.sourceFolder) Then
        swApp.SendMsgToUser2 "Source folder does not exist:" & vbCrLf & opts.sourceFolder, swMbStop, swMbOk
        Exit Sub
    End If

    If opts.doDXFExport Then
        If Trim$(opts.exportFolder) = "" Then
            swApp.SendMsgToUser2 "Please provide a DXF export folder.", swMbStop, swMbOk
            Exit Sub
        End If
        If Not FolderExists(opts.exportFolder) Then
            swApp.SendMsgToUser2 "DXF export folder does not exist:" & vbCrLf & opts.exportFolder, swMbStop, swMbOk
            Exit Sub
        End If
    End If

    If opts.doPropertyUpdate Then
        If Trim$(opts.modelInitials) = "" Or Trim$(opts.drawingInitials) = "" Then
            swApp.SendMsgToUser2 "Please enter both model initials and drawing initials.", swMbStop, swMbOk
            Exit Sub
        End If
    End If

    If Not opts.doPropertyUpdate And Not opts.doDXFExport Then
        swApp.SendMsgToUser2 "Nothing selected to run.", swMbWarning, swMbOk
        Exit Sub
    End If

    Dim confirmMsg As String
    confirmMsg = "Source Folder:" & vbCrLf & opts.sourceFolder & vbCrLf & vbCrLf & _
                 "Property Update: " & IIf(opts.doPropertyUpdate, "YES", "NO") & vbCrLf & _
                 "DXF Export: " & IIf(opts.doDXFExport, "YES", "NO") & vbCrLf & _
                 "Skip '003-' files: " & IIf(opts.skip003, "YES", "NO") & vbCrLf & vbCrLf

    If opts.doPropertyUpdate Then
        confirmMsg = confirmMsg & "Model DrawnBy: " & opts.modelInitials & vbCrLf & _
                                  "Drawing DwgDrawnBy: " & opts.drawingInitials & vbCrLf & vbCrLf
    End If

    If opts.doDXFExport Then
        confirmMsg = confirmMsg & "DXF Export Folder:" & vbCrLf & opts.exportFolder & vbCrLf & vbCrLf
    End If

    confirmMsg = confirmMsg & "Proceed?"

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Combined Batch Tool") <> vbYes Then Exit Sub

    Dim propResult As PropertyBatchResult
    Dim dxfResult As DXFBatchResult
    Dim summary As String

    If opts.doPropertyUpdate Then
        propResult = RunPropertyUpdateBatch(opts)
        summary = summary & "PROPERTY UPDATE" & vbCrLf & _
                  "Processed: " & propResult.processed & vbCrLf & _
                  "Updated: " & propResult.updated & vbCrLf & _
                  "Skipped: " & propResult.skipped & vbCrLf & _
                  "Failed: " & propResult.failed & vbCrLf
        If propResult.skipped003 > 0 Then
            summary = summary & "Skipped 003- files: " & propResult.skipped003 & vbCrLf
        End If
        summary = summary & "Log: " & propResult.logPath & vbCrLf & vbCrLf
    End If

    If opts.doDXFExport Then
        dxfResult = RunDXFExportBatch(opts)
        summary = summary & "DXF EXPORT" & vbCrLf & _
                  "Processed: " & dxfResult.processed & vbCrLf & _
                  "Exported: " & dxfResult.exported & vbCrLf & _
                  "Skipped: " & dxfResult.skipped & vbCrLf & _
                  "Failed: " & dxfResult.failed & vbCrLf & _
                  "Log: " & dxfResult.logPath & vbCrLf
    End If

    swApp.SendMsgToUser2 summary, swMbInformation, swMbOk

End Sub

Private Function ShowCombinedDialog(ByRef opts As BatchOptions) As Boolean

    ShowCombinedDialog = False

    Dim htaPath As String
    htaPath = Environ("TEMP") & "\SW_CombinedBatchTool.hta"

    Dim resultPath As String
    resultPath = Environ("TEMP") & "\SW_CombinedBatchTool_result.txt"

    ' If params file already exists, skip HTA dialog (called from external automation)
    If Dir(resultPath) <> "" Then GoTo ReadResult

    On Error Resume Next
    Kill resultPath
    On Error GoTo 0

    Dim f As Integer
    f = FreeFile
    Open htaPath For Output As #f

    Print #f, "<html>"
    Print #f, "<head>"
    Print #f, "<title>SolidWorks Batch Utility</title>"
    Print #f, "<HTA:APPLICATION ID=""BatchTool"" APPLICATIONNAME=""SolidWorks Batch Utility"" BORDER=""thin"" BORDERSTYLE=""normal"" INNERBORDER=""no"" CAPTION=""yes"" MAXIMIZEBUTTON=""no"" MINIMIZEBUTTON=""no"" SYSMENU=""yes"" SCROLL=""yes"" SINGLEINSTANCE=""yes"" WINDOWSTATE=""normal"">"
    Print #f, "<style>"
    Print #f, "body{font-family:Segoe UI,Tahoma,sans-serif;background:#eef2f7;margin:0;padding:18px;color:#1f2937;}"
    Print #f, ".wrap{background:#ffffff;border:1px solid #d7dee8;border-radius:14px;padding:18px 18px 14px 18px;box-shadow:0 10px 25px rgba(0,0,0,0.08);}"
    Print #f, ".hero{margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid #e5e7eb;}"
    Print #f, ".hero h1{margin:0;font-size:18pt;color:#111827;}"
    Print #f, ".hero p{margin:6px 0 0 0;font-size:9pt;color:#6b7280;}"
    Print #f, ".grid{display:block;}"
    Print #f, ".card{border:1px solid #e5e7eb;border-radius:12px;padding:12px 12px 10px 12px;margin-bottom:12px;background:#fbfcfe;}"
    Print #f, ".card h2{margin:0 0 10px 0;font-size:11pt;color:#111827;}"
    Print #f, ".hint{font-size:8.5pt;color:#6b7280;margin:-4px 0 8px 0;}"
    Print #f, "label{display:block;margin:8px 0 4px 0;font-weight:600;font-size:9pt;color:#374151;}"
    Print #f, "input[type=text]{width:100%;padding:8px 10px;font-size:9pt;border:1px solid #c7d0db;border-radius:8px;box-sizing:border-box;background:#fff;}"
    Print #f, "input[type=text]:focus{outline:none;border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,0.12);}"
    Print #f, ".row{display:flex;gap:12px;}"
    Print #f, ".col{flex:1;}"
    Print #f, ".toggle{display:flex;align-items:center;gap:8px;margin:2px 0 10px 0;padding:8px 10px;background:#f3f6fb;border:1px solid #dbe4f0;border-radius:8px;}"
    Print #f, ".toggle label{margin:0;font-weight:600;}"
    Print #f, ".foot{display:flex;justify-content:space-between;align-items:center;border-top:1px solid #e5e7eb;padding-top:12px;margin-top:4px;}"
    Print #f, ".footnote{font-size:8.5pt;color:#6b7280;max-width:420px;}"
    Print #f, "button{padding:8px 18px;font-size:9pt;border-radius:8px;border:1px solid #9aa8b8;cursor:pointer;}"
    Print #f, ".run{background:#2563eb;color:#fff;border-color:#1d4ed8;font-weight:600;}"
    Print #f, ".run:hover{background:#1d4ed8;}"
    Print #f, ".cancel{background:#f3f4f6;color:#111827;}"
    Print #f, "</style>"
    Print #f, "<script language=""VBScript"">"
    Print #f, "Sub Window_OnLoad"
    Print #f, "  window.resizeTo 920, 820"
    Print #f, "  window.moveTo (screen.width-920)/2, (screen.height-820)/2"
    Print #f, "  document.getElementById(""txtSource"").focus"
    Print #f, "  UpdateState"
    Print #f, "End Sub"
    Print #f, "Sub UpdateState"
    Print #f, "  Dim propOn, dxfOn"
    Print #f, "  propOn = document.getElementById(""chkProps"").checked"
    Print #f, "  dxfOn = document.getElementById(""chkDXF"").checked"
    Print #f, "  If propOn Then"
    Print #f, "    document.getElementById(""propFields"").style.display = ""block"""
    Print #f, "  Else"
    Print #f, "    document.getElementById(""propFields"").style.display = ""none"""
    Print #f, "  End If"
    Print #f, "  If dxfOn Then"
    Print #f, "    document.getElementById(""dxfFields"").style.display = ""block"""
    Print #f, "  Else"
    Print #f, "    document.getElementById(""dxfFields"").style.display = ""none"""
    Print #f, "  End If"
    Print #f, "End Sub"
    Print #f, "Sub btnRun_Click"
    Print #f, "  Dim fso, tf"
    Print #f, "  Set fso = CreateObject(""Scripting.FileSystemObject"")"
    Print #f, "  Set tf = fso.CreateTextFile(""" & Replace(resultPath, "\", "\\") & """, True)"
    Print #f, "  tf.WriteLine document.getElementById(""txtSource"").value"
    Print #f, "  tf.WriteLine document.getElementById(""txtExport"").value"
    Print #f, "  tf.WriteLine document.getElementById(""txtModel"").value"
    Print #f, "  tf.WriteLine document.getElementById(""txtDraw"").value"
    Print #f, "  If document.getElementById(""chkSkip"").checked Then"
    Print #f, "    tf.WriteLine ""YES"""
    Print #f, "  Else"
    Print #f, "    tf.WriteLine ""NO"""
    Print #f, "  End If"
    Print #f, "  If document.getElementById(""chkProps"").checked Then"
    Print #f, "    tf.WriteLine ""YES"""
    Print #f, "  Else"
    Print #f, "    tf.WriteLine ""NO"""
    Print #f, "  End If"
    Print #f, "  If document.getElementById(""chkDXF"").checked Then"
    Print #f, "    tf.WriteLine ""YES"""
    Print #f, "  Else"
    Print #f, "    tf.WriteLine ""NO"""
    Print #f, "  End If"
    Print #f, "  tf.Close"
    Print #f, "  self.close"
    Print #f, "End Sub"
    Print #f, "Sub btnCancel_Click"
    Print #f, "  self.close"
    Print #f, "End Sub"
    Print #f, "Sub CheckEnter()"
    Print #f, "  If window.event.keyCode = 13 Then btnRun_Click"
    Print #f, "End Sub"
    Print #f, "</script>"
    Print #f, "</head>"
    Print #f, "<body onkeypress=""CheckEnter()"">"
    Print #f, "<div class=""wrap"">"
    Print #f, "  <div class=""hero"">"
    Print #f, "    <h1>SolidWorks Batch Utility</h1>"
    Print #f, "    <p>Run property cleanup and DXF export from one place.</p>"
    Print #f, "  </div>"
    Print #f, "  <div class=""card"">"
    Print #f, "    <h2>1. Source folder</h2>"
    Print #f, "    <p class=""hint"">All SolidWorks files are scanned here. DXF export uses .sldprt files from this same folder.</p>"
    Print #f, "    <label for=""txtSource"">Source folder path</label>"
    Print #f, "    <input type=""text"" id=""txtSource"" value="""">"
    Print #f, "    <div class=""toggle""><input type=""checkbox"" id=""chkSkip""><label for=""chkSkip"">Skip files starting with '003-'</label></div>"
    Print #f, "  </div>"
    Print #f, "  <div class=""card"">"
    Print #f, "    <h2>2. Tasks to run</h2>"
    Print #f, "    <div class=""toggle""><input type=""checkbox"" id=""chkProps"" checked onclick=""UpdateState()""><label for=""chkProps"">Update custom properties and drawing revision table</label></div>"
    Print #f, "    <div id=""propFields"">"
    Print #f, "      <div class=""row"">"
    Print #f, "        <div class=""col""><label for=""txtModel"">DrawnBy initials for parts / assemblies</label><input type=""text"" id=""txtModel"" value=""""></div>"
    Print #f, "        <div class=""col""><label for=""txtDraw"">DwgDrawnBy initials for drawings</label><input type=""text"" id=""txtDraw"" value=""""></div>"
    Print #f, "      </div>"
    Print #f, "    </div>"
    Print #f, "    <div class=""toggle""><input type=""checkbox"" id=""chkDXF"" checked onclick=""UpdateState()""><label for=""chkDXF"">Export sheet metal parts to DXF</label></div>"
    Print #f, "    <div id=""dxfFields"">"
    Print #f, "      <label for=""txtExport"">DXF destination folder</label>"
    Print #f, "      <input type=""text"" id=""txtExport"" value="""">"
    Print #f, "    </div>"
    Print #f, "  </div>"
    Print #f, "  <div class=""foot"">"
    Print #f, "    <div class=""footnote"">Logs are written to the source folder for property updates and to the DXF folder for export results.</div>"
    Print #f, "    <div><button class=""cancel"" onclick=""btnCancel_Click()"">Cancel</button> <button class=""run"" onclick=""btnRun_Click()"">Run Batch</button></div>"
    Print #f, "  </div>"
    Print #f, "</div>"
    Print #f, "</body>"
    Print #f, "</html>"

    Close #f

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & htaPath & """", 1, True

    If Dir(resultPath) = "" Then Exit Function

ReadResult:
    Dim lines(1 To 7) As String
    Dim lineNum As Long
    lineNum = 0

    f = FreeFile
    Open resultPath For Input As #f
    Do While Not EOF(f) And lineNum < 7
        lineNum = lineNum + 1
        Line Input #f, lines(lineNum)
    Loop
    Close #f

    On Error Resume Next
    Kill htaPath
    Kill resultPath
    On Error GoTo 0

    If lineNum < 7 Then Exit Function

    opts.sourceFolder = Trim$(lines(1))
    opts.exportFolder = Trim$(lines(2))
    opts.modelInitials = Trim$(lines(3))
    opts.drawingInitials = Trim$(lines(4))
    opts.skip003 = (UCase$(Trim$(lines(5))) = "YES")
    opts.doPropertyUpdate = (UCase$(Trim$(lines(6))) = "YES")
    opts.doDXFExport = (UCase$(Trim$(lines(7))) = "YES")

    If opts.sourceFolder = "" Then Exit Function

    ShowCombinedDialog = True

End Function

Private Function RunPropertyUpdateBatch(ByRef opts As BatchOptions) As PropertyBatchResult

    Dim result As PropertyBatchResult
    Dim jPaths() As String
    Dim jNames() As String
    Dim jTypes() As Long
    Dim jIsDrw() As Boolean
    Dim jobCount As Long

    Dim fName As String
    fName = Dir(opts.sourceFolder & "*.*")

    Do While fName <> ""
        Dim uName As String
        uName = UCase$(fName)

        If opts.skip003 And Left$(fName, 4) = "003-" Then
            result.skipped003 = result.skipped003 + 1
            GoTo NextScanFile
        End If

        Dim dt As Long
        Dim isDrw As Boolean
        dt = -1

        If Right$(uName, 7) = ".SLDPRT" Then
            dt = swDocPART
        ElseIf Right$(uName, 7) = ".SLDASM" Then
            dt = swDocASSEMBLY
        ElseIf Right$(uName, 7) = ".SLDDRW" Then
            dt = swDocDRAWING
            isDrw = True
        End If

        If dt >= 0 Then
            jobCount = jobCount + 1
            ReDim Preserve jPaths(1 To jobCount)
            ReDim Preserve jNames(1 To jobCount)
            ReDim Preserve jTypes(1 To jobCount)
            ReDim Preserve jIsDrw(1 To jobCount)
            jPaths(jobCount) = opts.sourceFolder & fName
            jNames(jobCount) = fName
            jTypes(jobCount) = dt
            jIsDrw(jobCount) = isDrw
        End If

NextScanFile:
        fName = Dir
    Loop

    result.logPath = opts.sourceFolder & "Combined_Property_Update_Log.txt"

    If jobCount = 0 Then
        WriteSimpleLog result.logPath, "No SolidWorks files found to update."
        RunPropertyUpdateBatch = result
        Exit Function
    End If

    Dim updateLog As String, skipLog As String, failLog As String
    Dim overallStart As Single
    overallStart = Timer

    Dim i As Long
    For i = 1 To jobCount
        result.processed = result.processed + 1

        Dim filePath As String, fileName As String, docType As Long, isDrawing As Boolean
        filePath = jPaths(i)
        fileName = jNames(i)
        docType = jTypes(i)
        isDrawing = jIsDrw(i)

        Dim isRO As Boolean
        On Error Resume Next
        isRO = ((GetAttr(filePath) And vbReadOnly) = vbReadOnly)
        On Error GoTo 0

        If isRO Then
            skipLog = skipLog & "SKIP - Read-only: " & fileName & vbCrLf
            result.skipped = result.skipped + 1
            GoTo NextJob
        End If

        CloseIfOpen filePath

        Dim openErrors As Long, openWarnings As Long
        Dim swModel As SldWorks.ModelDoc2
        Set swModel = swApp.OpenDoc6(filePath, docType, swOpenDocOptions_Silent, "", openErrors, openWarnings)

        If swModel Is Nothing Then
            failLog = failLog & "FAIL - Could not open: " & fileName & " (error code: " & openErrors & ")" & vbCrLf
            result.failed = result.failed + 1
            GoTo NextJob
        End If

        Dim partLog As String
        Dim didChange As Boolean

        If isDrawing Then
            If Not EnforceDrawingProperties(swModel, opts.drawingInitials, partLog, didChange) Then
                failLog = failLog & "FAIL - Property enforcement failed: " & fileName & vbCrLf & partLog
                result.failed = result.failed + 1
                swApp.CloseDoc swModel.GetTitle
                GoTo NextJob
            End If

            If Not UpdateDrawingRevisionTable(swModel, fileName, partLog, didChange) Then
                failLog = failLog & "FAIL - Revision table failed: " & fileName & vbCrLf & partLog
                result.failed = result.failed + 1
                swApp.CloseDoc swModel.GetTitle
                GoTo NextJob
            End If
        Else
            If Not SetOrCreateCustomPropertyIfNeeded(swModel, "DrawnBy", opts.modelInitials, didChange) Then
                failLog = failLog & "FAIL - Could not set DrawnBy in: " & fileName & vbCrLf
                result.failed = result.failed + 1
                swApp.CloseDoc swModel.GetTitle
                GoTo NextJob
            End If
        End If

        If didChange Then
            Dim saveErrors As Long, saveWarnings As Long
            swModel.SetSaveFlag
            Call swModel.Save3(swSaveAsOptions_Silent, saveErrors, saveWarnings)

            If saveErrors <> 0 Then
                failLog = failLog & "FAIL - Could not save: " & fileName & " (save error: " & saveErrors & ")" & vbCrLf
                result.failed = result.failed + 1
            Else
                If isDrawing Then
                    updateLog = updateLog & "UPDATED - " & fileName & " | DwgDrawnBy=" & opts.drawingInitials & " | Revision=A | Description=INITIAL RELEASE" & vbCrLf
                Else
                    updateLog = updateLog & "UPDATED - " & fileName & " | DrawnBy=" & opts.modelInitials & vbCrLf
                End If
                result.updated = result.updated + 1
            End If
        Else
            skipLog = skipLog & "SKIP - No change needed: " & fileName & vbCrLf
            result.skipped = result.skipped + 1
        End If

        swApp.CloseDoc swModel.GetTitle

NextJob:
    Next i

    Dim elapsed As Double
    elapsed = ElapsedSeconds(overallStart)

    Dim lf As Integer
    lf = FreeFile
    Open result.logPath For Output As #lf
    Print #lf, "COMBINED PROPERTY UPDATE LOG"
    Print #lf, "Started: " & Now
    Print #lf, "Folder: " & opts.sourceFolder
    Print #lf, "DrawnBy initials: " & opts.modelInitials
    Print #lf, "DwgDrawnBy initials: " & opts.drawingInitials
    Print #lf, String(70, "=")
    Print #lf, ""
    Print #lf, "FAILURES"
    Print #lf, String(70, "-")
    If failLog <> "" Then Print #lf, failLog Else Print #lf, "None" & vbCrLf
    Print #lf, "SKIPS"
    Print #lf, String(70, "-")
    If skipLog <> "" Then Print #lf, skipLog Else Print #lf, "None" & vbCrLf
    Print #lf, "UPDATES"
    Print #lf, String(70, "-")
    If updateLog <> "" Then Print #lf, updateLog Else Print #lf, "None" & vbCrLf
    Print #lf, "SUMMARY"
    Print #lf, String(70, "-")
    Print #lf, "Processed: " & result.processed
    Print #lf, "Updated: " & result.updated
    Print #lf, "Skipped: " & result.skipped
    Print #lf, "Failed: " & result.failed
    Print #lf, "Skipped 003-: " & result.skipped003
    Print #lf, "Elapsed: " & FormatSeconds(elapsed)
    Close #lf

    RunPropertyUpdateBatch = result

End Function

Private Function RunDXFExportBatch(ByRef opts As BatchOptions) As DXFBatchResult

    Dim result As DXFBatchResult
    result.logPath = opts.exportFolder & "Combined_DXF_Export_Log.txt"

    Dim overallStart As Single
    overallStart = Timer

    Dim fileName As String
    fileName = Dir(opts.sourceFolder & "*.sldprt")

    Dim failLog As String, warnLog As String, okLog As String

    Do While fileName <> ""
        If opts.skip003 And Left$(fileName, 4) = "003-" Then
            result.skipped = result.skipped + 1
            warnLog = warnLog & "FILE: " & fileName & vbCrLf & "  SKIP - Starts with 003-." & vbCrLf & vbCrLf
            GoTo NextDXFFile
        End If

        result.processed = result.processed + 1

        Dim partLog As String
        Dim partResultValue As PartResult
        partResultValue = ProcessPart(fileName, opts.sourceFolder, opts.exportFolder, partLog)

        Select Case partResultValue
            Case prExported
                okLog = okLog & partLog & vbCrLf
                result.exported = result.exported + 1
            Case prFailed
                failLog = failLog & partLog & vbCrLf
                result.failed = result.failed + 1
            Case prSkipped
                warnLog = warnLog & partLog & vbCrLf
                result.skipped = result.skipped + 1
        End Select

NextDXFFile:
        fileName = Dir
    Loop

    Dim overallElapsed As Double
    overallElapsed = ElapsedSeconds(overallStart)

    Dim avgPerPart As Double
    If result.processed > 0 Then avgPerPart = overallElapsed / result.processed

    Dim logFile As Integer
    logFile = FreeFile
    Open result.logPath For Output As #logFile
    Print #logFile, "COMBINED DXF EXPORT LOG"
    Print #logFile, "Started: " & Now
    Print #logFile, "Source Folder: " & opts.sourceFolder
    Print #logFile, "Destination Folder: " & opts.exportFolder
    Print #logFile, String(70, "=")
    Print #logFile, ""
    Print #logFile, "ERRORS / FAILURES"
    Print #logFile, String(70, "-")
    If failLog <> "" Then Print #logFile, failLog Else Print #logFile, "None" & vbCrLf
    Print #logFile, "SKIPS / WARNINGS"
    Print #logFile, String(70, "-")
    If warnLog <> "" Then Print #logFile, warnLog Else Print #logFile, "None" & vbCrLf
    Print #logFile, "SUCCESSFUL EXPORTS"
    Print #logFile, String(70, "-")
    If okLog <> "" Then Print #logFile, okLog Else Print #logFile, "None" & vbCrLf
    Print #logFile, "SUMMARY"
    Print #logFile, String(70, "-")
    Print #logFile, "Processed: " & result.processed
    Print #logFile, "Exported: " & result.exported
    Print #logFile, "Failed: " & result.failed
    Print #logFile, "Skipped: " & result.skipped
    Print #logFile, "Average Time / Part: " & FormatSeconds(avgPerPart)
    Print #logFile, "Total Time: " & FormatSeconds(overallElapsed)
    Close #logFile

    RunDXFExportBatch = result

End Function

Private Sub CloseIfOpen(ByVal filePath As String)
    On Error Resume Next
    Dim vDocs As Variant
    vDocs = swApp.GetDocuments
    If Not IsEmpty(vDocs) Then
        Dim d As Long
        For d = 0 To UBound(vDocs)
            Dim tmpDoc As SldWorks.ModelDoc2
            Set tmpDoc = vDocs(d)
            If Not tmpDoc Is Nothing Then
                If LCase$(tmpDoc.GetPathName) = LCase$(filePath) Then
                    swApp.CloseDoc tmpDoc.GetTitle
                    Exit For
                End If
            End If
        Next d
    End If
    On Error GoTo 0
End Sub

Private Sub WriteSimpleLog(ByVal logPath As String, ByVal textOut As String)
    Dim f As Integer
    f = FreeFile
    Open logPath For Output As #f
    Print #f, textOut
    Close #f
End Sub

Function SetOrCreateCustomPropertyIfNeeded(ByVal swModel As SldWorks.ModelDoc2, _
                                           ByVal propName As String, _
                                           ByVal propValue As String, _
                                           ByRef didChange As Boolean) As Boolean

    On Error GoTo EH

    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")

    Dim valOut As String
    Dim resolvedValOut As String
    Dim wasResolved As Boolean
    Dim linkToProp As Boolean

    valOut = ""
    resolvedValOut = ""

    Call swCustPropMgr.Get6(propName, False, valOut, resolvedValOut, wasResolved, linkToProp)

    If StrComp(Trim$(resolvedValOut), Trim$(propValue), vbTextCompare) = 0 Or _
       StrComp(Trim$(valOut), Trim$(propValue), vbTextCompare) = 0 Then
        SetOrCreateCustomPropertyIfNeeded = True
        Exit Function
    End If

    Dim addResult As Long
    addResult = swCustPropMgr.Add3(propName, swCustomInfoText, propValue, swCustomPropertyDeleteAndAdd)

    If addResult >= 0 Then
        didChange = True
        SetOrCreateCustomPropertyIfNeeded = True
    Else
        SetOrCreateCustomPropertyIfNeeded = (swCustPropMgr.Set2(propName, propValue) <> 0)
        If SetOrCreateCustomPropertyIfNeeded Then didChange = True
    End If

    Exit Function

EH:
    SetOrCreateCustomPropertyIfNeeded = False

End Function

Function EnforceDrawingProperties(ByVal swModel As SldWorks.ModelDoc2, _
                                  ByVal dwgDrawnByValue As String, _
                                  ByRef partLog As String, _
                                  ByRef didChange As Boolean) As Boolean

    On Error GoTo EH

    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")

    If swCustPropMgr Is Nothing Then
        partLog = partLog & "  Could not get CustomPropertyManager." & vbCrLf
        EnforceDrawingProperties = False
        Exit Function
    End If

    Dim existingValues(1 To 15) As String
    Dim existingExpressions(1 To 15) As String
    Dim propIdx As Long

    For propIdx = 1 To 15
        Dim pName As String
        pName = GetDrawingPropertyName(propIdx)

        Dim valOut As String
        Dim resolvedValOut As String
        Dim wasResolved As Boolean
        Dim linkToProp As Boolean

        valOut = ""
        resolvedValOut = ""

        Call swCustPropMgr.Get6(pName, False, valOut, resolvedValOut, wasResolved, linkToProp)

        existingExpressions(propIdx) = valOut
        existingValues(propIdx) = resolvedValOut
    Next propIdx

    Dim vNames As Variant
    vNames = swCustPropMgr.GetNames

    Dim alreadyCorrect As Boolean
    alreadyCorrect = True

    If IsEmpty(vNames) Then
        alreadyCorrect = False
    Else
        If UBound(vNames) - LBound(vNames) + 1 <> 15 Then
            alreadyCorrect = False
        Else
            Dim chkIdx As Long
            For chkIdx = 0 To UBound(vNames)
                If StrComp(CStr(vNames(chkIdx)), GetDrawingPropertyName(chkIdx + 1), vbTextCompare) <> 0 Then
                    alreadyCorrect = False
                    Exit For
                End If
            Next chkIdx
        End If
    End If

    Dim dwgDrawnByCurrent As String
    dwgDrawnByCurrent = Trim$(existingValues(7))
    If dwgDrawnByCurrent = "" Then dwgDrawnByCurrent = Trim$(existingExpressions(7))

    Dim dwgDrawnByMatch As Boolean
    dwgDrawnByMatch = (StrComp(dwgDrawnByCurrent, Trim$(dwgDrawnByValue), vbTextCompare) = 0)

    If alreadyCorrect And dwgDrawnByMatch Then
        EnforceDrawingProperties = True
        Exit Function
    End If

    If Not IsEmpty(vNames) Then
        Dim delIdx As Long
        For delIdx = LBound(vNames) To UBound(vNames)
            swCustPropMgr.Delete2 CStr(vNames(delIdx))
        Next delIdx
    End If

    For propIdx = 1 To 15
        pName = GetDrawingPropertyName(propIdx)

        Dim valueToSet As String
        If propIdx = 7 Then
            valueToSet = dwgDrawnByValue
        Else
            valueToSet = existingExpressions(propIdx)
        End If

        Dim addResult As Long
        addResult = swCustPropMgr.Add3(pName, swCustomInfoText, valueToSet, swCustomPropertyOnlyIfNew)

        If addResult < 0 Then
            partLog = partLog & "  WARN - Could not add property: " & pName & " (result: " & addResult & ")" & vbCrLf
        End If
    Next propIdx

    didChange = True
    EnforceDrawingProperties = True
    Exit Function

EH:
    partLog = partLog & "  Exception in EnforceDrawingProperties: " & Err.Description & vbCrLf
    EnforceDrawingProperties = False

End Function

Function GetDrawingPropertyName(ByVal idx As Long) As String
    Select Case idx
        Case 1:  GetDrawingPropertyName = "SWFormatSize"
        Case 2:  GetDrawingPropertyName = "Revision"
        Case 3:  GetDrawingPropertyName = "Description"
        Case 4:  GetDrawingPropertyName = "Material"
        Case 5:  GetDrawingPropertyName = "Finish"
        Case 6:  GetDrawingPropertyName = "DrawnBy"
        Case 7:  GetDrawingPropertyName = "DwgDrawnBy"
        Case 8:  GetDrawingPropertyName = "Bend Deduction"
        Case 9:  GetDrawingPropertyName = "Top Die"
        Case 10: GetDrawingPropertyName = "Bottom Die"
        Case 11: GetDrawingPropertyName = "Tol X"
        Case 12: GetDrawingPropertyName = "Tol X.X"
        Case 13: GetDrawingPropertyName = "Tol X.XX"
        Case 14: GetDrawingPropertyName = "Tol X.XXX"
        Case 15: GetDrawingPropertyName = Chr$(84) & Chr$(111) & Chr$(108) & Chr$(32) & Chr$(176)
        Case Else: GetDrawingPropertyName = ""
    End Select
End Function

Function UpdateDrawingRevisionTable(ByVal swModel As SldWorks.ModelDoc2, _
                                    ByVal fileName As String, _
                                    ByRef partLog As String, _
                                    ByRef didChange As Boolean) As Boolean

    On Error GoTo EH

    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel

    Dim vSheetNames As Variant
    vSheetNames = swDraw.GetSheetNames

    If IsEmpty(vSheetNames) Then
        partLog = partLog & "  No sheets found." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    Dim firstSheetName As String
    firstSheetName = CStr(vSheetNames(LBound(vSheetNames)))
    Call swDraw.ActivateSheet(firstSheetName)

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        partLog = partLog & "  Could not get current sheet." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    Dim swRevTable As SldWorks.RevisionTableAnnotation
    Dim insertedNewTable As Boolean
    Set swRevTable = swSheet.RevisionTable

    If swRevTable Is Nothing Then
        Dim tPath As String
        tPath = "\\npsvr05\FOXFAB\FOXFAB_DATA\ENGINEERING\SOLIDWORKS\Foxfab Templates\Revision Table v1.1.sldrevtbt"

        Set swRevTable = swSheet.InsertRevisionTable(True, 0#, 0#, swBOMConfigurationAnchor_TopRight, tPath)
        If swRevTable Is Nothing Then
            partLog = partLog & "  InsertRevisionTable failed using template." & vbCrLf
            UpdateDrawingRevisionTable = False
            Exit Function
        End If

        insertedNewTable = True
        partLog = partLog & "  INFO - Revision table inserted from template." & vbCrLf
    End If

    Dim swTable As SldWorks.TableAnnotation
    Set swTable = swRevTable

    Dim descCol As Long, r As Long, c As Long
    descCol = -1
    For r = 0 To swTable.RowCount - 1
        For c = 0 To swTable.ColumnCount - 1
            If UCase$(Trim$(swTable.Text2(r, c, False))) = "DESCRIPTION" Then
                descCol = c
                Exit For
            End If
        Next c
        If descCol >= 0 Then Exit For
    Next r

    If descCol < 0 Then
        partLog = partLog & "  Could not find DESCRIPTION column." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    If Not insertedNewTable Then
        Dim rowIdx As Long, revId As Long
        For rowIdx = swTable.RowCount - 1 To 2 Step -1
            revId = swRevTable.GetIdForRowNumber(rowIdx)
            If revId <> -1 Then swRevTable.DeleteRevision revId, True
        Next rowIdx
    End If

    Dim newRevId As Long
    newRevId = swRevTable.AddRevision("A")

    If newRevId < 0 Then
        partLog = partLog & "  Could not add revision A." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    Dim newRow As Long
    newRow = swRevTable.GetRowNumberForId(newRevId)
    If newRow < 0 Then
        partLog = partLog & "  Could not find new revision row." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    swTable.Text2(newRow, descCol, True) = "INITIAL RELEASE"
    didChange = True
    UpdateDrawingRevisionTable = True
    Exit Function

EH:
    partLog = partLog & "  Exception: " & Err.Description & vbCrLf
    UpdateDrawingRevisionTable = False

End Function

Private Function ProcessPart(ByVal fileName As String, _
                             ByVal sourceFolder As String, _
                             ByVal destFolder As String, _
                             ByRef partLog As String) As PartResult

    Dim partStart As Single
    partStart = Timer

    Dim filePath As String
    filePath = sourceFolder & fileName

    Dim fileBaseName As String
    fileBaseName = Left$(fileName, InStrRev(fileName, ".") - 1)

    Dim dxfPath As String
    dxfPath = destFolder & fileBaseName & ".dxf"

    Dim preOpenDocs As Object
    Set preOpenDocs = GetOpenDocDict()

    partLog = "FILE: " & fileName & vbCrLf

    Dim errors As Long
    Dim warnings As Long
    Dim swModel As SldWorks.ModelDoc2
    Dim swPart As SldWorks.PartDoc

    Set swModel = swApp.OpenDoc6(filePath, swDocPART, swOpenDocOptions_Silent, "", errors, warnings)

    If swModel Is Nothing Then
        partLog = partLog & "  FAIL - Could not open file (OpenDoc6 error: " & CStr(errors) & ")" & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prFailed
        GoTo Cleanup
    End If

    Set swPart = swModel

    If Not FastRebuild(swModel, partLog, True) Then
        partLog = partLog & "  FAIL - Initial rebuild failed." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prFailed
        GoTo Cleanup
    End If

    Dim info As PartAnalysis
    info = AnalyzePart(swModel)

    If Not info.HasSheetMetal Then
        partLog = partLog & "  SKIP - Not a sheet metal part." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prSkipped
        GoTo Cleanup
    End If

    If info.MultipleConfigs Then
        partLog = partLog & "  SKIP - Multiple user configurations found." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prSkipped
        GoTo Cleanup
    End If

    If info.flatPatternName = "" And info.HasBends Then
        partLog = partLog & "  FAIL - No Flat-Pattern feature found." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prFailed
        GoTo Cleanup
    End If

    If info.HasRolledBack Then
        If SilentRebuild(swModel, True) Then
            partLog = partLog & "  INFO - Rolled forward, continuing." & vbCrLf
        Else
            partLog = partLog & "  INFO - Rollback bar already at end." & vbCrLf
        End If
    End If

    If info.HasBends Then
        If Not FlattenPart(swModel, info.flatPatternName, partLog) Then
            partLog = partLog & "  FAIL - Could not flatten part." & vbCrLf
            AppendElapsed partLog, partStart
            ProcessPart = prFailed
            GoTo Cleanup
        End If

        If Not FastRebuild(swModel, partLog, True) Then
            partLog = partLog & "  FAIL - Rebuild failed after flatten." & vbCrLf
            AppendElapsed partLog, partStart
            ProcessPart = prFailed
            GoTo Cleanup
        End If

        Dim flatErrText As String
        If CheckFeatureErrors(swModel, flatErrText, info.flatPatternName) Then
            partLog = partLog & "  FAIL - Feature tree error(s) after flatten:" & vbCrLf & flatErrText
            AppendElapsed partLog, partStart
            ProcessPart = prFailed
            GoTo Cleanup
        End If

        partLog = partLog & "  INFO - Flat pattern validation passed." & vbCrLf
    Else
        partLog = partLog & "  INFO - No bends detected, skipping flat pattern validation." & vbCrLf
    End If

    Dim alignArr(11) As Double
    alignArr(0) = 0#: alignArr(1) = 0#: alignArr(2) = 0#
    alignArr(3) = 1#: alignArr(4) = 0#: alignArr(5) = 0#
    alignArr(6) = 0#: alignArr(7) = 1#: alignArr(8) = 0#
    alignArr(9) = 0#: alignArr(10) = 0#: alignArr(11) = 1#

    Dim smOptions As Long
    smOptions = 71

    If swPart.ExportToDWG2(dxfPath, filePath, 1, True, alignArr, False, False, smOptions, Empty) Then
        partLog = partLog & "  OK - Exported to: " & dxfPath & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prExported
    Else
        partLog = partLog & "  FAIL - ExportToDWG2 returned False." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prFailed
    End If

Cleanup:
    On Error Resume Next
    If Not swModel Is Nothing Then
        swApp.CloseDoc swModel.GetTitle
        Set swModel = Nothing
        Set swPart = Nothing
    End If
    CloseExtraDocs preOpenDocs
    Set preOpenDocs = Nothing
    On Error GoTo 0

End Function

Private Function AnalyzePart(ByVal swModel As SldWorks.ModelDoc2) As PartAnalysis

    Dim result As PartAnalysis
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FirstFeature

    Do While Not swFeat Is Nothing
        Dim featType As String
        featType = LCase$(swFeat.GetTypeName2)

        Select Case featType
            Case "sheetmetal"
                result.HasSheetMetal = True
            Case "flatpattern"
                If result.flatPatternName = "" Then result.flatPatternName = swFeat.Name
            Case "edgeflange", "sketchbend", "foldfeature", "jog", "loftbend", "baseflangewall", "hem", "miterflange", "crossbreak", "smbaseflangewall"
                result.HasBends = True
        End Select

        If swFeat.IsRolledBack Then result.HasRolledBack = True
        Set swFeat = swFeat.GetNextFeature
    Loop

    On Error Resume Next
    Dim vConfNames As Variant
    vConfNames = swModel.GetConfigurationNames
    If Not IsEmpty(vConfNames) Then
        Dim userCount As Long, i As Long
        For i = LBound(vConfNames) To UBound(vConfNames)
            Dim upperName As String
            upperName = UCase$(Trim$(CStr(vConfNames(i))))
            If InStr(upperName, "SM-FLAT-PATTERN") = 0 And InStr(upperName, "FLAT-PATTERN") = 0 Then
                userCount = userCount + 1
            End If
        Next i
        result.MultipleConfigs = (userCount > 1)
    End If
    On Error GoTo 0

    AnalyzePart = result

End Function

Private Function CheckFeatureErrors(ByVal swModel As SldWorks.ModelDoc2, _
                                    ByRef errorText As String, _
                                    ByVal flatPatternName As String) As Boolean

    CheckFeatureErrors = False
    errorText = ""

    On Error Resume Next

    Dim flatFeat As SldWorks.Feature
    Set flatFeat = swModel.FeatureByName(flatPatternName)

    If flatFeat Is Nothing Then
        errorText = "    - Flat-Pattern feature not found: " & flatPatternName & vbCrLf
        CheckFeatureErrors = True
        Exit Function
    End If

    Dim errCode As Long
    Dim isWarning As Boolean
    errCode = flatFeat.GetErrorCode2(isWarning)

    If Err.Number = 0 Then
        If errCode <> 0 And Not isWarning Then
            CheckFeatureErrors = True
            errorText = "    - Feature: " & flatFeat.Name & " | Type: " & flatFeat.GetTypeName2 & " | ErrorCode: " & CStr(errCode) & " | IsWarning: " & CStr(isWarning) & vbCrLf
        End If
    End If

    Err.Clear
    On Error GoTo 0

End Function

Private Function FlattenPart(ByVal swModel As SldWorks.ModelDoc2, _
                             ByVal flatPatternName As String, _
                             ByRef partLog As String) As Boolean

    On Error GoTo EH

    Dim flatFeat As SldWorks.Feature
    Set flatFeat = swModel.FeatureByName(flatPatternName)

    If flatFeat Is Nothing Then
        partLog = partLog & "  ERROR - Flat-Pattern feature not found: " & flatPatternName & vbCrLf
        FlattenPart = False
        Exit Function
    End If

    partLog = partLog & "  INFO - Flattening: " & flatFeat.Name & vbCrLf

    swModel.ClearSelection2 True
    If flatFeat.Select2(False, 0) Then
        swModel.EditUnsuppress2
        partLog = partLog & "  INFO - Flattened via EditUnsuppress2." & vbCrLf
        FlattenPart = True
        Exit Function
    End If

    partLog = partLog & "  INFO - Select2 failed, using SetSuppression2 fallback." & vbCrLf

    Dim confNames(0) As String
    confNames(0) = swModel.ConfigurationManager.ActiveConfiguration.Name

    If flatFeat.SetSuppression2(swUnSuppressFeature, swSpecifyConfiguration, confNames) = False Then
        partLog = partLog & "  ERROR - SetSuppression2 fallback also failed." & vbCrLf
        FlattenPart = False
        Exit Function
    End If

    FlattenPart = True
    Exit Function

EH:
    partLog = partLog & "  ERROR - Exception during flatten: " & Err.Description & vbCrLf
    FlattenPart = False

End Function

Private Function FastRebuild(ByVal swModel As SldWorks.ModelDoc2, _
                             ByRef partLog As String, _
                             Optional ByVal topLevelOnly As Boolean = True) As Boolean

    On Error GoTo EH

    Dim ok As Boolean
    ok = swModel.ForceRebuild3(topLevelOnly)

    If ok Then
        partLog = partLog & "  INFO - Rebuild succeeded." & vbCrLf
    Else
        partLog = partLog & "  WARN - ForceRebuild3 returned False." & vbCrLf
    End If

    FastRebuild = True
    Exit Function

EH:
    partLog = partLog & "  ERROR - Exception during rebuild: " & Err.Description & vbCrLf
    FastRebuild = False

End Function

Private Function SilentRebuild(ByVal swModel As SldWorks.ModelDoc2, _
                               Optional ByVal topLevelOnly As Boolean = True) As Boolean
    On Error GoTo EH
    swModel.ForceRebuild3 topLevelOnly
    SilentRebuild = True
    Exit Function
EH:
    SilentRebuild = False
End Function

Private Function GetOpenDocDict() As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim vDocs As Variant
    vDocs = swApp.GetDocuments
    If IsEmpty(vDocs) Then
        Set GetOpenDocDict = dict
        Exit Function
    End If

    Dim i As Long
    For i = 0 To UBound(vDocs)
        Dim swDoc As SldWorks.ModelDoc2
        Set swDoc = vDocs(i)
        If Not swDoc Is Nothing Then
            Dim docPath As String
            docPath = LCase$(swDoc.GetPathName)
            If docPath <> "" And Not dict.Exists(docPath) Then dict.Add docPath, True
        End If
    Next i

    Set GetOpenDocDict = dict

End Function

Private Sub CloseExtraDocs(ByVal baselineDict As Object)

    On Error Resume Next
    If baselineDict Is Nothing Then Exit Sub

    Dim vDocs As Variant
    vDocs = swApp.GetDocuments
    If IsEmpty(vDocs) Then Exit Sub

    Dim i As Long
    For i = 0 To UBound(vDocs)
        Dim swDoc As SldWorks.ModelDoc2
        Set swDoc = vDocs(i)
        If Not swDoc Is Nothing Then
            Dim thisPath As String
            thisPath = LCase$(swDoc.GetPathName)
            If thisPath <> "" Then
                If Not baselineDict.Exists(thisPath) Then swApp.CloseDoc swDoc.GetTitle
            End If
        End If
    Next i
    On Error GoTo 0

End Sub

Private Sub AppendElapsed(ByRef partLog As String, ByVal partStart As Single)
    partLog = partLog & "  Elapsed: " & FormatSeconds(ElapsedSeconds(partStart)) & vbCrLf
End Sub

Private Function NormalizeFolder(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath <> "" Then
        If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    End If
    NormalizeFolder = folderPath
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

Private Function ElapsedSeconds(ByVal startT As Single) As Double
    Dim t As Double
    t = Timer - startT
    If t < 0 Then t = t + 86400
    ElapsedSeconds = t
End Function

Private Function FormatSeconds(ByVal totalSec As Double) As String
    Dim h As Long, m As Long, s As Long
    h = Int(totalSec / 3600)
    m = Int((totalSec - h * 3600) / 60)
    s = Int(totalSec - h * 3600 - m * 60)
    FormatSeconds = Format$(h, "00") & ":" & Format$(m, "00") & ":" & Format$(s, "00")
End Function


