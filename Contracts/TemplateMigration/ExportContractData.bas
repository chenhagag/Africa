Attribute VB_Name = "ExportContractData"
' ============================================================
' ExportContractData - Extract content control values from migrated contracts
' and export to CSV for bulk SharePoint column update
'
' HOW TO USE:
' 1. Open all migrated contracts in Word
' 2. Alt+F8 > ExportAllContracts
' 3. CSV saved to TemplateMigration\ExportedData.csv
'
' The CSV can then be used by UpdateSharePoint.ps1 to update
' the SharePoint library columns.
' ============================================================

Const BASE_PATH As String = "C:\SPFX\Africa\GIT-Africa\Africa\Contracts\TemplateMigration\"
Const CSV_FILE As String = "ExportedData.csv"

' ============================================================
' All tags we want to extract (matching the Add-in's TAGS + fields)
' ============================================================
Private Function GetAllTags() As Variant
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
    GetAllTags = t
End Function

' ============================================================
' Map CC tag -> SharePoint column internal name
' (matches the saveToHelper field names in taskpane.ts)
' ============================================================
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

' ============================================================
' Check if a tag is a multi-line field (preserve line breaks)
' ============================================================
Private Function IsMultiLineTag(tag As String) As Boolean
    Select Case tag
        Case "cntTzadA", "cntTzadB"
            IsMultiLineTag = True
        Case Else
            IsMultiLineTag = False
    End Select
End Function

' ============================================================
' Get the value of a content control by tag from a document
' Searches body + headers + footers
' ============================================================
Private Function GetCCValue(doc As Document, tag As String) As String
    Dim cc As ContentControl
    Dim multiLine As Boolean
    multiLine = IsMultiLineTag(tag)

    ' Search body
    For Each cc In doc.ContentControls
        If cc.tag = tag Then
            GetCCValue = CleanText(cc.Range.Text, multiLine)
            Exit Function
        End If
    Next cc

    ' Search headers and footers
    Dim sec As Section
    Dim hf As HeaderFooter
    For Each sec In doc.Sections
        For Each hf In sec.Headers
            For Each cc In hf.Range.ContentControls
                If cc.tag = tag Then
                    GetCCValue = CleanText(cc.Range.Text, multiLine)
                    Exit Function
                End If
            Next cc
        Next hf
        For Each hf In sec.Footers
            For Each cc In hf.Range.ContentControls
                If cc.tag = tag Then
                    GetCCValue = CleanText(cc.Range.Text, multiLine)
                    Exit Function
                End If
            Next cc
        Next hf
    Next sec

    GetCCValue = ""
End Function

' ============================================================
' Clean text for CSV: remove line breaks, escape quotes
' ============================================================
Private Function CleanText(txt As String, Optional keepLineBreaks As Boolean = False) As String
    Dim result As String
    result = txt

    ' Remove cell markers always
    result = Replace(result, Chr(7), "")

    If keepLineBreaks Then
        ' For multi-line fields: normalize line breaks but keep them
        result = Replace(result, vbCr & vbLf, vbLf)
        result = Replace(result, vbCr, vbLf)
        result = Replace(result, Chr(11), vbLf)
        result = Replace(result, Chr(13), vbLf)
    Else
        ' For single-line fields: replace all line breaks with space
        result = Replace(result, vbCr, " ")
        result = Replace(result, vbLf, " ")
        result = Replace(result, Chr(11), " ")
        result = Replace(result, Chr(13), " ")
        ' Trim multiple spaces
        Do While InStr(result, "  ") > 0
            result = Replace(result, "  ", " ")
        Loop
    End If

    result = Trim(result)

    ' Filter out placeholder text
    If LCase(result) = "click or tap here to enter text." Or _
       LCase(result) = "click or tap here to enter text" Or _
       LCase(result) Like "[[]*[]]" Then
        CleanText = ""
        Exit Function
    End If

    CleanText = result
End Function

' ============================================================
' Escape a value for CSV (double-quote wrapping)
' ============================================================
Private Function CsvEscape(val As String) As String
    ' Always wrap in quotes to handle commas and Hebrew text
    CsvEscape = """" & Replace(val, """", """""") & """"
End Function

' ============================================================
' PUBLIC: Export data from a single open document (for testing)
' ============================================================
Sub ExportSingleContract()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim tags As Variant
    tags = GetAllTags()

    Dim msg As String
    msg = "Content Controls in: " & doc.Name & vbCrLf & vbCrLf

    Dim i As Long
    For i = LBound(tags) To UBound(tags)
        Dim val As String
        val = GetCCValue(doc, CStr(tags(i)))
        If val <> "" Then
            msg = msg & TagToSPColumn(CStr(tags(i))) & " = " & Left(val, 80) & vbCrLf
        End If
    Next i

    MsgBox msg, vbInformation, "Export Preview"
End Sub

' ============================================================
' PUBLIC: Export data from all open contracts to CSV
' ============================================================
Sub ExportAllContracts()
    Dim tags As Variant
    tags = GetAllTags()

    ' Collect document names
    Dim docNames() As String
    Dim fileCount As Long
    fileCount = 0
    Dim d As Document
    For Each d In Application.Documents
        If LCase(d.Name) <> "normal.dotm" And LCase(d.Name) <> "normal.dot" Then
            fileCount = fileCount + 1
            ReDim Preserve docNames(1 To fileCount)
            docNames(fileCount) = d.Name
        End If
    Next d

    If fileCount = 0 Then
        MsgBox "No documents are open.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Found " & fileCount & " open documents:" & vbCrLf & vbCrLf & _
              Join(docNames, vbCrLf) & vbCrLf & vbCrLf & _
              "Export data to CSV?", vbYesNo + vbQuestion, "Export Contract Data") = vbNo Then
        Exit Sub
    End If

    ' Build CSV header
    Dim header As String
    header = CsvEscape("FileName")
    Dim i As Long
    For i = LBound(tags) To UBound(tags)
        header = header & "," & CsvEscape(TagToSPColumn(CStr(tags(i))))
    Next i

    ' Build CSV rows
    Dim allRows As String
    allRows = header & vbCrLf

    Dim doc As Document
    Dim row As String
    Dim val As String
    Dim succeeded As Long
    Dim j As Long

    For j = 1 To fileCount
        On Error Resume Next
        Set doc = Application.Documents(docNames(j))
        If Err.Number <> 0 Or doc Is Nothing Then
            Debug.Print "ERROR: Could not access " & docNames(j)
            Err.Clear
            GoTo NextDoc
        End If
        On Error GoTo 0

        row = CsvEscape(doc.Name)

        For i = LBound(tags) To UBound(tags)
            val = GetCCValue(doc, CStr(tags(i)))
            row = row & "," & CsvEscape(val)
        Next i

        allRows = allRows & row & vbCrLf
        succeeded = succeeded + 1
        Debug.Print "Exported: " & doc.Name

NextDoc:
        On Error GoTo 0
    Next j

    ' Write CSV file (UTF-8 with BOM for Hebrew support)
    Dim csvPath As String
    csvPath = BASE_PATH & CSV_FILE

    Dim fNum As Integer
    fNum = FreeFile

    ' Use ADODB.Stream for proper UTF-8 encoding
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2  ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText allRows
    stream.SaveToFile csvPath, 2  ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing

    MsgBox "Export complete!" & vbCrLf & vbCrLf & _
           "Documents exported: " & succeeded & "/" & fileCount & vbCrLf & _
           "CSV file: " & csvPath, vbInformation, "Export Contract Data"
End Sub
