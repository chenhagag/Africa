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

    ' If still failing, remove all parent CCs up the chain (multi-level nesting)
    Dim retryCount As Long
    retryCount = 0
    Do While newCC Is Nothing And retryCount < 10
        Dim parentCC As ContentControl
        Set parentCC = Nothing
        On Error Resume Next
        Set parentCC = targetRange.ParentContentControl
        On Error GoTo 0
        If parentCC Is Nothing Then Exit Do

        Err.Clear
        stp = "6b-removeparent-" & retryCount
        On Error Resume Next
        parentCC.LockContentControl = False
        parentCC.LockContents = False
        If parentCC.XMLMapping.IsMapped Then parentCC.XMLMapping.Delete
        parentCC.Delete False
        Err.Clear

        stp = "6c-retry-" & retryCount
        Set newCC = targetRange.ContentControls.Add(wdContentControlRichText)
        If Err.Number <> 0 Or newCC Is Nothing Then
            Err.Clear
            Set newCC = targetRange.ContentControls.Add(wdContentControlText)
            If Err.Number <> 0 Then Err.Clear
        End If

        retryCount = retryCount + 1
    Loop
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
' ============================================================
' CSV Export helper functions (integrated from ExportContractData)
' ============================================================
Private Function GetAllExportTags() As Variant
    Dim t(0 To 32) As String
    t(0) = "cntContractNumber"
    t(1) = "cntContractVersion"
    t(2) = "cntTemplateName"
    t(3) = "cntProjectName"
    t(4) = "cntSite"
    t(5) = "cntSupllierName"
    t(6) = "cntMunicipality"
    t(7) = "cntWorkDescription"
    t(8) = "cntSignDate"
    t(9) = "cntStartDate"
    t(10) = "cntDurationMonths"
    t(11) = "cntExpectedEndDate"
    t(12) = "cntStatus"
    t(13) = "cntTzadA"
    t(14) = "cntTzadB"
    t(15) = "cntPartyAName"
    t(16) = "cmtTzadAPercent"
    t(17) = "cntCostCompMethod"
    t(18) = "cntCostContractScope"
    t(19) = "cntCostCurrency"
    t(20) = "cntCostIndexType"
    t(21) = "cntCostBaseIndexDate"
    t(22) = "cntCostIndexMode"
    t(23) = "cntCostIndexPoints"
    t(24) = "cntCostPaymentTerms"
    t(25) = "cntCustomField1"
    t(26) = "cntCustomField2"
    t(27) = "cntCustomField3"
    t(28) = "cntCustomField4"
    t(29) = "cntCustomField5"
    t(30) = "cntCustomField6"
    t(31) = "cntCustomField7"
    t(32) = "cntCustomField8"
    GetAllExportTags = t
End Function

Private Function TagToSPColumn(tag As String) As String
    Select Case tag
        Case "cntContractNumber": TagToSPColumn = "ContractNumber"
        Case "cntContractVersion": TagToSPColumn = "contractVersion"
        Case "cntTemplateName": TagToSPColumn = "ContractTemplate"
        Case "cntProjectName": TagToSPColumn = "project"
        Case "cntSite": TagToSPColumn = "SiteName"
        Case "cntSupllierName": TagToSPColumn = "supplierName"
        Case "cntMunicipality": TagToSPColumn = "Municipality"
        Case "cntWorkDescription": TagToSPColumn = "WorkDescription"
        Case "cntSignDate": TagToSPColumn = "signDate"
        Case "cntStartDate": TagToSPColumn = "StartDate"
        Case "cntDurationMonths": TagToSPColumn = "DurationMonths"
        Case "cntExpectedEndDate": TagToSPColumn = "ExpectedEndDate"
        Case "cntStatus": TagToSPColumn = "status"
        Case "cntTzadA": TagToSPColumn = "recipient"
        Case "cntTzadB": TagToSPColumn = "otherSides"
        Case "cntPartyAName": TagToSPColumn = "partyAName"
        Case "cmtTzadAPercent": TagToSPColumn = "PartyAContactNamePercent"
        Case "cntCostCompMethod": TagToSPColumn = "CostCompMethod"
        Case "cntCostContractScope": TagToSPColumn = "CostContractScope"
        Case "cntCostCurrency": TagToSPColumn = "CostCurrency"
        Case "cntCostIndexType": TagToSPColumn = "CostIndexType"
        Case "cntCostBaseIndexDate": TagToSPColumn = "CostBaseIndexDate"
        Case "cntCostIndexMode": TagToSPColumn = "CostIndexMode"
        Case "cntCostIndexPoints": TagToSPColumn = "CostIndexPoints"
        Case "cntCostPaymentTerms": TagToSPColumn = "CostPaymentTerms"
        Case "cntCustomField1": TagToSPColumn = "customField1"
        Case "cntCustomField2": TagToSPColumn = "customField2"
        Case "cntCustomField3": TagToSPColumn = "customField3"
        Case "cntCustomField4": TagToSPColumn = "customField4"
        Case "cntCustomField5": TagToSPColumn = "customField5"
        Case "cntCustomField6": TagToSPColumn = "customField6"
        Case "cntCustomField7": TagToSPColumn = "customField7"
        Case "cntCustomField8": TagToSPColumn = "customField8"
        Case Else: TagToSPColumn = tag
    End Select
End Function

Private Function IsMultiLineTag(tag As String) As Boolean
    Select Case tag
        Case "cntTzadA", "cntTzadB"
            IsMultiLineTag = True
        Case Else
            IsMultiLineTag = False
    End Select
End Function

Private Function GetCCValue(doc As Document, tag As String) As String
    Dim cc As ContentControl
    Dim multiLine As Boolean
    multiLine = IsMultiLineTag(tag)

    For Each cc In doc.ContentControls
        If cc.tag = tag Then
            GetCCValue = CleanExportText(cc.Range.Text, multiLine)
            Exit Function
        End If
    Next cc

    Dim sec As Section
    Dim hf As HeaderFooter
    For Each sec In doc.Sections
        For Each hf In sec.Headers
            For Each cc In hf.Range.ContentControls
                If cc.tag = tag Then
                    GetCCValue = CleanExportText(cc.Range.Text, multiLine)
                    Exit Function
                End If
            Next cc
        Next hf
        For Each hf In sec.Footers
            For Each cc In hf.Range.ContentControls
                If cc.tag = tag Then
                    GetCCValue = CleanExportText(cc.Range.Text, multiLine)
                    Exit Function
                End If
            Next cc
        Next hf
    Next sec

    GetCCValue = ""
End Function

Private Function CleanExportText(txt As String, Optional keepLineBreaks As Boolean = False) As String
    Dim result As String
    result = txt
    result = Replace(result, Chr(7), "")

    If keepLineBreaks Then
        result = Replace(result, vbCr & vbLf, vbLf)
        result = Replace(result, vbCr, vbLf)
        result = Replace(result, Chr(11), vbLf)
        result = Replace(result, Chr(13), vbLf)
    Else
        result = Replace(result, vbCr, " ")
        result = Replace(result, vbLf, " ")
        result = Replace(result, Chr(11), " ")
        result = Replace(result, Chr(13), " ")
        Do While InStr(result, "  ") > 0
            result = Replace(result, "  ", " ")
        Loop
    End If

    result = Trim(result)

    If LCase(result) = "click or tap here to enter text." Or _
       LCase(result) = "click or tap here to enter text" Or _
       LCase(result) Like "[[]*[]]" Then
        CleanExportText = ""
        Exit Function
    End If

    CleanExportText = result
End Function

' ============================================================
' Fix duplicated contract numbers like "CONT-5-668CONT-10-2195" -> "CONT-10-2195"
' Takes the last CONT-xx-xxxx occurrence
' ============================================================
Private Function FixContractNumber(val As String) As String
    Dim pos As Long
    Dim lastPos As Long
    lastPos = 0
    pos = InStr(1, val, "CONT-")
    Do While pos > 0
        lastPos = pos
        pos = InStr(pos + 1, val, "CONT-")
    Loop
    If lastPos > 1 Then
        ' There were multiple CONT- occurrences, take the last one
        FixContractNumber = Mid(val, lastPos)
    Else
        FixContractNumber = val
    End If
End Function

Private Function CsvEscape(val As String) As String
    CsvEscape = """" & Replace(val, """", """""") & """"
End Function

Private Function BuildCsvRowWithName(doc As Document, fileName As String, folderSiteName As String) As String
    Dim tags As Variant
    tags = GetAllExportTags()
    Dim row As String
    row = CsvEscape(fileName)
    Dim i As Long
    Dim val As String
    For i = LBound(tags) To UBound(tags)
        val = GetCCValue(doc, CStr(tags(i)))
        ' Fallback to old tags if new tag not found
        If val = "" Then
            Select Case CStr(tags(i))
                Case "cntContractNumber": val = GetCCValue(doc, "_dlc_DocId")
                Case "cntTemplateName": val = GetCCValue(doc, "cntContractType")
                Case "cntPartyAName": val = GetCCValue(doc, "cmtTzadAName")
                Case "cntMunicipality": val = GetCCValue(doc, "cntLocalAuth")
                Case "cntWorkDescription": val = GetCCValue(doc, "cntJobDesc")
                Case "cntTzadB": val = GetCCValue(doc, "cntTzadB_x002C__x0020_cntTzadC_x002C__x0020_cntTzadD")
            End Select
        End If
        If CStr(tags(i)) = "cntContractNumber" Then val = FixContractNumber(val)
        ' Override site name with folder name (always more reliable)
        If CStr(tags(i)) = "cntSite" And folderSiteName <> "" Then val = folderSiteName
        row = row & "," & CsvEscape(val)
    Next i
    BuildCsvRowWithName = row
End Function

Private Function BuildCsvHeader() As String
    Dim tags As Variant
    tags = GetAllExportTags()
    Dim header As String
    header = CsvEscape("FileName")
    Dim i As Long
    For i = LBound(tags) To UBound(tags)
        header = header & "," & CsvEscape(TagToSPColumn(CStr(tags(i))))
    Next i
    BuildCsvHeader = header
End Function

' ============================================================
' PUBLIC: Migrate all contracts AND export data to CSV
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

    ' Prepare CSV
    Dim csvData As String
    csvData = BuildCsvHeader() & vbCrLf

    Dim doc As Document
    Dim result As String
    Dim summary As String
    Dim outputPath As String
    Dim contractNum As String
    Dim baseName As String
    Dim newFileName As String
    Dim folderSiteName As String
    Dim parts() As String
    Dim i As Long
    Dim succeeded As Long
    Dim failed As Long

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    For i = 1 To fileCount
        Debug.Print "[" & i & "/" & fileCount & "] " & docNames(i)

        On Error Resume Next
        Set doc = Nothing
        Set doc = Application.Documents(docNames(i))
        If Err.Number <> 0 Or doc Is Nothing Then
            summary = summary & "ERROR: " & docNames(i) & " - not found" & vbCrLf
            failed = failed + 1
            Err.Clear
            GoTo NextContract
        End If
        On Error GoTo 0

        doc.Activate
        DoEvents

        ' Process contract with error handling - continue even if one doc fails
        On Error Resume Next
        result = ProcessContract(doc)
        If Err.Number <> 0 Then
            summary = summary & "ERROR migrating: " & docNames(i) & " - " & Err.Description & vbCrLf
            Debug.Print "ERROR migrating: " & docNames(i) & " - " & Err.Description
            failed = failed + 1
            Err.Clear
            On Error Resume Next
            doc.Close SaveChanges:=False
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            Set doc = Nothing
            GoTo NextContract
        End If
        On Error GoTo 0

        ' Build output filename: append contract number to avoid duplicates
        ' Try new tag first, fall back to old tag
        contractNum = GetCCValue(doc, "cntContractNumber")
        If contractNum = "" Then contractNum = GetCCValue(doc, "_dlc_DocId")
        contractNum = FixContractNumber(contractNum)
        baseName = Left(docNames(i), Len(docNames(i)) - 5) ' remove .docx

        ' Build new filename with contract number
        If contractNum <> "" Then
            newFileName = baseName & " - " & contractNum & ".docx"
        Else
            newFileName = docNames(i)
        End If

        ' Extract site name from parent folder of original file path
        ' e.g. "...\Old Contracts\נרקיסים\חוזה.docx" -> "נרקיסים"
        folderSiteName = ""
        If InStr(docPaths(i), "\") > 0 Then
            parts = Split(docPaths(i), "\")
            If UBound(parts) >= 1 Then
                folderSiteName = parts(UBound(parts) - 1)
                ' Don't use folder name if it's a known non-site folder
                If LCase(folderSiteName) = "old contracts" Or _
                   LCase(folderSiteName) = "new contracts" Or _
                   LCase(folderSiteName) = "desktop" Then
                    folderSiteName = ""
                End If
            End If
        End If

        ' Export data to CSV row with the NEW filename (matching what goes to SP)
        csvData = csvData & BuildCsvRowWithName(doc, newFileName, folderSiteName) & vbCrLf

        outputPath = newDir & newFileName

        On Error Resume Next
        doc.SaveAs2 outputPath, wdFormatDocumentDefault
        If Err.Number <> 0 Then
            summary = summary & "ERROR saving: " & docNames(i) & " - " & Err.Description & vbCrLf
            failed = failed + 1
            Err.Clear
            doc.Close SaveChanges:=False
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            Set doc = Nothing
            GoTo NextContract
        End If

        doc.Close SaveChanges:=False
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

        DoEvents

        summary = summary & docNames(i) & " -> " & newFileName & " - " & result & vbCrLf
        succeeded = succeeded + 1

        Set doc = Nothing
NextContract:
        On Error GoTo 0
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Write CSV file (UTF-8)
    Dim csvPath As String
    csvPath = BASE_PATH & "ExportedData.csv"
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText csvData
    stream.SaveToFile csvPath, 2
    stream.Close
    Set stream = Nothing

    MsgBox "Contract migration complete!" & vbCrLf & vbCrLf & _
           "Succeeded: " & succeeded & vbCrLf & _
           "Failed: " & failed & vbCrLf & vbCrLf & _
           "CSV exported to: " & csvPath & vbCrLf & vbCrLf & _
           summary, vbInformation, "Migrate All Contracts"
End Sub
