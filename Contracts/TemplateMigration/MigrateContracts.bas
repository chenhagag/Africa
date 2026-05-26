Attribute VB_Name = "MigrateContracts"
' ============================================================
' MigrateContracts - Convert old SP-bound CONTRACTS (with data) to clean add-in format
'
' Unlike MigrateTemplates which sets placeholder text, this module
' PRESERVES the existing values in content controls.
'
' HOW TO USE:
' 1. Open Word (any blank document)
' 2. Alt+F11 (or Developer > Visual Basic)
' 3. File > Import File > select this .bas file
' 4. Run "MigrateSingleContract" to test on the active document, or
'    Run "MigrateAllContracts" to process all .docx in Old Contracts folder
'
' INPUT:  Old Contracts\   (place filled-in contracts here)
' OUTPUT: New Contracts\   (converted contracts saved here)
' ============================================================

' --- Base path (change if needed) ---
Const BASE_PATH As String = "C:\SPFX\Africa\GIT-Africa\Africa\Contracts\TemplateMigration\"

' ============================================================
' Tag mappings: old tag -> new tag
' ============================================================
Private Function GetNewTag(oldTag As String) As String
    Select Case oldTag
        Case "cntTzadB_x002C__x0020_cntTzadC_x002C__x0020_cntTzadD"
            GetNewTag = "cntTzadB"
        Case "cmtTzadAName"
            GetNewTag = "cntPartyAName"
        Case "cntLocalAuth"
            GetNewTag = "cntMunicipality"
        Case "cntContractType"
            GetNewTag = "cntTemplateName"
        Case "_dlc_DocId"
            GetNewTag = "cntContractNumber"
        Case "cntJobDesc"
            GetNewTag = "cntWorkDescription"
        Case Else
            GetNewTag = oldTag  ' keep as-is
    End Select
End Function

' ============================================================
' Tag -> Hebrew label (matches FIELD_CATALOG in taskpane.ts)
' ============================================================
Private Function GetLabel(tag As String) As String
    Select Case tag
        Case "cntContractNumber": GetLabel = ChrW$(1502) & ChrW$(1505) & ChrW$(1508) & ChrW$(1512) & " " & ChrW$(1495) & ChrW$(1493) & ChrW$(1494) & ChrW$(1492)
        Case "cntContractVersion": GetLabel = ChrW$(1490) & ChrW$(1512) & ChrW$(1505) & ChrW$(1514) & " " & ChrW$(1495) & ChrW$(1493) & ChrW$(1494) & ChrW$(1492)
        Case "cntTemplateName": GetLabel = ChrW$(1514) & ChrW$(1489) & ChrW$(1504) & ChrW$(1497) & ChrW$(1514)
        Case "cntProjectName": GetLabel = ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1492) & ChrW$(1508) & ChrW$(1512) & ChrW$(1493) & ChrW$(1497) & ChrW$(1511) & ChrW$(1496)
        Case "cntSite": GetLabel = ChrW$(1488) & ChrW$(1514) & ChrW$(1512)
        Case "cntSupllierName": GetLabel = ChrW$(1505) & ChrW$(1508) & ChrW$(1511)
        Case "cntMunicipality": GetLabel = ChrW$(1512) & ChrW$(1513) & ChrW$(1493) & ChrW$(1514) & " " & ChrW$(1502) & ChrW$(1511) & ChrW$(1493) & ChrW$(1502) & ChrW$(1497) & ChrW$(1514)
        Case "cntWorkDescription": GetLabel = ChrW$(1514) & ChrW$(1497) & ChrW$(1488) & ChrW$(1493) & ChrW$(1512) & " " & ChrW$(1492) & ChrW$(1506) & ChrW$(1489) & ChrW$(1493) & ChrW$(1491) & ChrW$(1492)
        Case "cntSignDate": GetLabel = ChrW$(1514) & ChrW$(1488) & ChrW$(1512) & ChrW$(1497) & ChrW$(1498) & " " & ChrW$(1495) & ChrW$(1514) & ChrW$(1497) & ChrW$(1502) & ChrW$(1492) & " " & ChrW$(1492) & ChrW$(1495) & ChrW$(1493) & ChrW$(1494) & ChrW$(1492)
        Case "cntStartDate": GetLabel = ChrW$(1502) & ChrW$(1493) & ChrW$(1506) & ChrW$(1491) & " " & ChrW$(1492) & ChrW$(1514) & ChrW$(1495) & ChrW$(1500) & ChrW$(1492)
        Case "cntDurationMonths": GetLabel = ChrW$(1502) & ChrW$(1505) & ChrW$(1508) & ChrW$(1512) & " " & ChrW$(1495) & ChrW$(1493) & ChrW$(1491) & ChrW$(1513) & ChrW$(1497) & ChrW$(1501)
        Case "cntExpectedEndDate": GetLabel = ChrW$(1502) & ChrW$(1493) & ChrW$(1506) & ChrW$(1491) & " " & ChrW$(1505) & ChrW$(1497) & ChrW$(1493) & ChrW$(1501) & " " & ChrW$(1510) & ChrW$(1508) & ChrW$(1493) & ChrW$(1497)
        Case "cntStatus": GetLabel = ChrW$(1505) & ChrW$(1496) & ChrW$(1496) & ChrW$(1493) & ChrW$(1505)
        Case "cntTzadA": GetLabel = ChrW$(1510) & ChrW$(1491) & " " & ChrW$(1488) & "'" & " " & ChrW$(1489) & ChrW$(1495) & ChrW$(1493) & ChrW$(1494) & ChrW$(1492)
        Case "cntTzadB": GetLabel = ChrW$(1510) & ChrW$(1491) & ChrW$(1491) & ChrW$(1497) & ChrW$(1501) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1508) & ChrW$(1497) & ChrW$(1501)
        Case "cntPartyAName": GetLabel = ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1510) & ChrW$(1491) & " " & ChrW$(1488) & "'"
        Case "cmtTzadAPercent": GetLabel = ChrW$(1488) & ChrW$(1495) & ChrW$(1493) & ChrW$(1494) & " " & ChrW$(1492) & ChrW$(1513) & ChrW$(1514) & ChrW$(1514) & ChrW$(1508) & ChrW$(1493) & ChrW$(1514) & " " & ChrW$(1510) & ChrW$(1491) & " " & ChrW$(1488) & "'"
        Case "cntMadadTypeTitle": GetLabel = ChrW$(1505) & ChrW$(1493) & ChrW$(1490) & " " & ChrW$(1502) & ChrW$(1491) & ChrW$(1491)
        Case "cntIsKnownTitle": GetLabel = ChrW$(1502) & ChrW$(1491) & ChrW$(1491) & " " & ChrW$(1489) & ChrW$(1490) & ChrW$(1497) & ChrW$(1503) & "/" & ChrW$(1497) & ChrW$(1491) & ChrW$(1493) & ChrW$(1506)
        Case "cntMadadBase": GetLabel = ChrW$(1502) & ChrW$(1491) & ChrW$(1491) & " " & ChrW$(1489) & ChrW$(1505) & ChrW$(1497) & ChrW$(1505)
        Case "cntMadadPoints": GetLabel = ChrW$(1504) & ChrW$(1511) & ChrW$(1493) & ChrW$(1491) & ChrW$(1493) & ChrW$(1514) & " " & ChrW$(1502) & ChrW$(1491) & ChrW$(1491)
        Case "cntCustomField1": GetLabel = ChrW$(1513) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1507) & " 1"
        Case "cntCustomField2": GetLabel = ChrW$(1513) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1507) & " 2"
        Case "cntCustomField3": GetLabel = ChrW$(1513) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1507) & " 3"
        Case "cntCustomField4": GetLabel = ChrW$(1513) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1507) & " 4"
        Case "cntCustomField5": GetLabel = ChrW$(1513) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1507) & " 5"
        Case "cntCustomField6": GetLabel = ChrW$(1513) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1507) & " 6"
        Case "cntCustomField7": GetLabel = ChrW$(1513) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1507) & " 7"
        Case "cntCustomField8": GetLabel = ChrW$(1513) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1505) & ChrW$(1507) & " 8"
        Case Else
            GetLabel = tag
    End Select
End Function

' ============================================================
' Convert a single content control: remove binding, update tag/title, KEEP value
' ============================================================
Private Function ConvertControl(cc As ContentControl) As Boolean
    Dim oldTag As String
    Dim newTag As String
    Dim label As String
    Dim stp As String
    Dim existingText As String

    On Error GoTo ErrHandler

    oldTag = Trim(cc.tag)
    If oldTag = "" Then
        ConvertControl = False
        Exit Function
    End If

    stp = "1-map"
    newTag = GetNewTag(oldTag)
    label = GetLabel(newTag)

    stp = "2-unlock"
    cc.LockContentControl = False
    cc.LockContents = False

    stp = "3-unmap"
    If cc.XMLMapping.IsMapped Then
        cc.XMLMapping.Delete
    End If

    stp = "4-saverange"
    ' Save the existing text and range before deleting the control
    existingText = cc.Range.Text
    Dim targetRange As Range
    Set targetRange = cc.Range

    stp = "5-delete"
    cc.Delete False  ' delete control, keep content

    stp = "6-addcc"
    ' targetRange still points to the same text after deletion
    Dim newCC As ContentControl
    On Error Resume Next
    Set newCC = targetRange.ContentControls.Add(wdContentControlRichText)
    If Err.Number <> 0 Then
        Err.Clear
        Set newCC = targetRange.ContentControls.Add(wdContentControlText)
    End If

    ' If still failing, check for parent CC (nested controls) and remove it
    If newCC Is Nothing Then
        Dim parentCC As ContentControl
        Set parentCC = targetRange.ParentContentControl
        If Not parentCC Is Nothing Then
            Err.Clear
            stp = "6b-removeparent"
            parentCC.LockContentControl = False
            parentCC.LockContents = False
            If parentCC.XMLMapping.IsMapped Then parentCC.XMLMapping.Delete
            parentCC.Delete False  ' delete parent, keep content
            ' Now try again — range is free
            stp = "6c-retry"
            Set newCC = targetRange.ContentControls.Add(wdContentControlRichText)
            If Err.Number <> 0 Then
                Err.Clear
                Set newCC = targetRange.ContentControls.Add(wdContentControlText)
            End If
        End If
    End If
    On Error GoTo ErrHandler

    If newCC Is Nothing Then
        Debug.Print "  DIAG tag=" & oldTag & " rangeStart=" & targetRange.Start & _
                    " rangeEnd=" & targetRange.End & " text=" & Left(targetRange.Text, 50)
        GoTo ErrHandler
    End If

    stp = "7-props"
    newCC.tag = newTag
    newCC.Title = label

    ' NOTE: We do NOT replace the text — the existing value is preserved
    ' The range text was kept by cc.Delete False and is now inside newCC

    Debug.Print "  OK: " & oldTag & " -> " & newTag & " val=""" & Left(existingText, 30) & """"

    ConvertControl = True
    Exit Function

ErrHandler:
    Debug.Print "FAIL at " & stp & " tag=" & oldTag & " err#" & Err.Number & ": " & Err.Description
    ConvertControl = False
End Function

' ============================================================
' Process a single document: convert all controls in body + headers/footers
' ============================================================
Private Function ProcessContract(doc As Document) As String
    Dim migrated As Long
    Dim skipped As Long
    Dim cc As ContentControl

    ' Process body controls
    Dim i As Long
    For i = doc.ContentControls.Count To 1 Step -1
        Set cc = doc.ContentControls(i)
        If ConvertControl(cc) Then
            migrated = migrated + 1
        Else
            skipped = skipped + 1
        End If
    Next i

    ' Process headers and footers
    Dim sec As Section
    Dim hf As HeaderFooter
    For Each sec In doc.Sections
        For Each hf In sec.Headers
            For i = hf.Range.ContentControls.Count To 1 Step -1
                Set cc = hf.Range.ContentControls(i)
                If ConvertControl(cc) Then
                    migrated = migrated + 1
                Else
                    skipped = skipped + 1
                End If
            Next i
        Next hf
        For Each hf In sec.Footers
            For i = hf.Range.ContentControls.Count To 1 Step -1
                Set cc = hf.Range.ContentControls(i)
                If ConvertControl(cc) Then
                    migrated = migrated + 1
                Else
                    skipped = skipped + 1
                End If
            Next i
        Next hf
    Next sec

    ProcessContract = migrated & " converted, " & skipped & " skipped"
End Function

' ============================================================
' PUBLIC: Migrate the currently open contract (for testing)
' ============================================================
Sub MigrateSingleContract()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim result As String
    result = ProcessContract(doc)

    ' Save to New Contracts
    Dim outputPath As String
    outputPath = BASE_PATH & "New Contracts\" & doc.Name

    ' Create output folder if needed
    If Dir(BASE_PATH & "New Contracts\", vbDirectory) = "" Then
        MkDir BASE_PATH & "New Contracts\"
    End If

    doc.SaveAs2 outputPath, wdFormatDocumentDefault

    MsgBox "Contract migration complete!" & vbCrLf & vbCrLf & _
           "File: " & doc.Name & vbCrLf & _
           "Result: " & result & vbCrLf & _
           "Saved to: " & outputPath, vbInformation, "Migrate Contract"
End Sub

' ============================================================
' PUBLIC: Migrate all contracts in Old Contracts folder
' ============================================================
Sub MigrateAllContracts()
    Dim newDir As String
    newDir = BASE_PATH & "New Contracts\"

    If Dir(newDir, vbDirectory) = "" Then
        MkDir newDir
    End If

    ' Work on all currently open documents (except Normal.dotm)
    Dim docCount As Long
    docCount = Application.Documents.Count

    If docCount = 0 Then
        MsgBox "No documents are open." & vbCrLf & vbCrLf & _
               "Open the contracts you want to migrate, then run this macro again.", vbExclamation
        Exit Sub
    End If

    ' Collect document names first (processing changes the Documents collection)
    Dim docNames() As String
    Dim docPaths() As String
    Dim fileCount As Long
    fileCount = 0
    Dim d As Document
    For Each d In Application.Documents
        If LCase(d.Name) <> "normal.dotm" And LCase(d.Name) <> "normal.dot" Then
            fileCount = fileCount + 1
            ReDim Preserve docNames(1 To fileCount)
            ReDim Preserve docPaths(1 To fileCount)
            docNames(fileCount) = d.Name
            docPaths(fileCount) = d.FullName
        End If
    Next d

    If fileCount = 0 Then
        MsgBox "No contract documents found open.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Found " & fileCount & " open documents to migrate:" & vbCrLf & vbCrLf & _
              Join(docNames, vbCrLf) & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, "Migrate All Contracts") = vbNo Then
        Exit Sub
    End If

    Dim doc As Document
    Dim result As String
    Dim summary As String
    Dim outputPath As String
    Dim i As Long
    Dim succeeded As Long
    Dim failed As Long

    For i = 1 To fileCount
        Debug.Print "[" & i & "/" & fileCount & "] " & docNames(i)

        On Error Resume Next
        Set doc = Application.Documents(docNames(i))
        If Err.Number <> 0 Or doc Is Nothing Then
            summary = summary & "ERROR: " & docNames(i) & " - not found" & vbCrLf
            failed = failed + 1
            Err.Clear
            GoTo NextContract
        End If

        doc.Activate
        DoEvents

        result = ProcessContract(doc)
        If Err.Number <> 0 Then
            summary = summary & "ERROR: " & docNames(i) & " - " & Err.Description & vbCrLf
            failed = failed + 1
            Err.Clear
            GoTo NextContract
        End If

        outputPath = newDir & docNames(i)
        doc.SaveAs2 outputPath, wdFormatDocumentDefault
        If Err.Number <> 0 Then
            summary = summary & "ERROR saving: " & docNames(i) & " - " & Err.Description & vbCrLf
            failed = failed + 1
            Err.Clear
            GoTo NextContract
        End If

        doc.Close SaveChanges:=False
        If Err.Number <> 0 Then Err.Clear

        DoEvents

        summary = summary & docNames(i) & " - " & result & vbCrLf
        succeeded = succeeded + 1

        Set doc = Nothing
NextContract:
        On Error GoTo 0
    Next i

    MsgBox "Contract migration complete!" & vbCrLf & vbCrLf & _
           "Succeeded: " & succeeded & vbCrLf & _
           "Failed: " & failed & vbCrLf & vbCrLf & _
           "Full log: " & logPath & vbCrLf & vbCrLf & _
           summary, vbInformation, "Migrate All Contracts"
End Sub
