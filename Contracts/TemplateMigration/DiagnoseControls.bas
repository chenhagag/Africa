Attribute VB_Name = "DiagnoseControls"
' ============================================================
' DiagnoseControls - Lists all content controls in the active document
'
' HOW TO USE:
' 1. Open an OLD template in Word
' 2. Press Alt+F11 to open the VBA editor
' 3. File > Import File > select this .bas file
' 4. Press F5 to run (or close editor and run from Macros menu)
' 5. A message box will show all content controls found
' ============================================================

Sub DiagnoseContentControls()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim cc As ContentControl
    Dim msg As String
    Dim i As Long

    msg = "Document: " & doc.Name & vbCrLf
    msg = msg & "Total Content Controls: " & doc.ContentControls.Count & vbCrLf & vbCrLf

    If doc.ContentControls.Count = 0 Then
        msg = msg & "NO CONTENT CONTROLS FOUND!"
        MsgBox msg, vbExclamation, "Diagnose Controls"
        Exit Sub
    End If

    ' List each control
    i = 1
    For Each cc In doc.ContentControls
        msg = msg & "#" & i & ": "
        msg = msg & "Tag=""" & cc.Tag & """ "
        msg = msg & "Title=""" & cc.Title & """ "
        msg = msg & "Type=" & cc.Type & " "

        ' Check if it has XML mapping (SharePoint binding)
        If Not cc.XMLMapping Is Nothing Then
            If cc.XMLMapping.IsMapped Then
                msg = msg & "[SP-BOUND: " & cc.XMLMapping.XPath & "]"
            End If
        End If

        msg = msg & vbCrLf
        i = i + 1
    Next cc

    ' Also check headers and footers
    Dim sec As Section
    Dim hf As HeaderFooter
    Dim hfCount As Long
    hfCount = 0

    For Each sec In doc.Sections
        For Each hf In sec.Headers
            hfCount = hfCount + hf.Range.ContentControls.Count
            For Each cc In hf.Range.ContentControls
                msg = msg & "#" & i & " [HEADER]: "
                msg = msg & "Tag=""" & cc.Tag & """ "
                msg = msg & "Title=""" & cc.Title & """"
                msg = msg & vbCrLf
                i = i + 1
            Next cc
        Next hf
        For Each hf In sec.Footers
            hfCount = hfCount + hf.Range.ContentControls.Count
            For Each cc In hf.Range.ContentControls
                msg = msg & "#" & i & " [FOOTER]: "
                msg = msg & "Tag=""" & cc.Tag & """ "
                msg = msg & "Title=""" & cc.Title & """"
                msg = msg & vbCrLf
                i = i + 1
            Next cc
        Next hf
    Next sec

    msg = msg & vbCrLf & "Header/Footer controls: " & hfCount

    ' Show result - if too long for MsgBox, also write to Immediate Window
    Debug.Print msg

    If Len(msg) > 1000 Then
        MsgBox "Found " & doc.ContentControls.Count & " controls in body + " & hfCount & " in headers/footers." & vbCrLf & vbCrLf & "Full list printed to Immediate Window (Ctrl+G in VBA Editor).", vbInformation, "Diagnose Controls"
    Else
        MsgBox msg, vbInformation, "Diagnose Controls"
    End If
End Sub
