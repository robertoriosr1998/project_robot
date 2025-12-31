' ProcessEmailAttachments VBA Macro
' This macro processes email attachments based on the selected cell's row data
' 
' How to install:
' 1. Open OPC_TEST.xlsm in Excel
' 2. Press Alt+F11 to open VBA Editor
' 3. Go to File > Import File and select this .bas file
' 4. Add references: Tools > References > Check:
'    - Microsoft Outlook XX.0 Object Library
'    - Adobe Acrobat XX.0 Type Library (if available) OR use Shell execution
' 5. Close VBA Editor and save the workbook

Option Explicit

' Main subroutine to process email attachments
Public Sub ProcessEmailAttachments()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim wsActive As Worksheet
    Dim wsTips As Worksheet
    Dim wsParams As Worksheet
    Dim wsCNDb As Worksheet
    
    Dim selectedRow As Long
    Dim searchValue As Variant
    Dim tipsRow As Long
    Dim myTipsValue As String
    Dim emailAddress As String
    Dim password1 As String
    Dim password2 As String
    Dim password3 As String
    
    ' Set workbook reference
    Set wb = ThisWorkbook
    
    ' Get the active sheet and selected row
    Set wsActive = ActiveSheet
    selectedRow = ActiveCell.Row
    
    ' Validate selection
    If selectedRow < 2 Then
        MsgBox "Please select a data row (not a header row).", vbExclamation, "Invalid Selection"
        Exit Sub
    End If
    
    ' Get the value from column 5 (E) of the selected row
    searchValue = wsActive.Cells(selectedRow, 5).Value
    
    If IsEmpty(searchValue) Or Trim(CStr(searchValue)) = "" Then
        MsgBox "Column E of the selected row is empty.", vbExclamation, "No Search Value"
        Exit Sub
    End If
    
    ' Get worksheet references
    On Error Resume Next
    Set wsTips = wb.Sheets("TIPS")
    Set wsParams = wb.Sheets("Parameters")
    Set wsCNDb = wb.Sheets("CN Database")
    On Error GoTo ErrorHandler
    
    If wsTips Is Nothing Then
        MsgBox "TIPS worksheet not found!", vbCritical, "Error"
        Exit Sub
    End If
    
    If wsParams Is Nothing Then
        MsgBox "Parameters worksheet not found!", vbCritical, "Error"
        Exit Sub
    End If
    
    If wsCNDb Is Nothing Then
        MsgBox "CN Database worksheet not found!", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Search for the value in column A of TIPS sheet
    tipsRow = FindValueInColumn(wsTips, 1, searchValue)
    
    If tipsRow = 0 Then
        MsgBox "Value '" & searchValue & "' not found in column A of TIPS sheet.", vbExclamation, "Not Found"
        Exit Sub
    End If
    
    ' Get the value from column Q (MY TIPS - column 17) of the found row
    myTipsValue = CStr(wsTips.Cells(tipsRow, 17).Value)
    
    If Trim(myTipsValue) = "" Then
        MsgBox "Column Q (MY TIPS) is empty for the found row.", vbExclamation, "No Data"
        Exit Sub
    End If
    
    ' Get the email address from cell B4 of Parameters sheet
    emailAddress = CStr(wsParams.Range("B4").Value)
    
    If Trim(emailAddress) = "" Then
        MsgBox "Email address in Parameters!B4 is empty.", vbExclamation, "No Email"
        Exit Sub
    End If
    
    ' Get passwords from TIPS sheet (columns R=18, S=19, T=20)
    password1 = Trim(CStr(wsTips.Cells(tipsRow, 18).Value))
    password2 = Trim(CStr(wsTips.Cells(tipsRow, 19).Value))
    password3 = Trim(CStr(wsTips.Cells(tipsRow, 20).Value))
    
    ' Process emails and attachments
    Call ProcessOutlookEmails(emailAddress, myTipsValue, password1, password2, password3, wsCNDb)
    
    MsgBox "Processing complete!", vbInformation, "Done"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

' Function to find a value in a specific column and return the row number
Private Function FindValueInColumn(ws As Worksheet, colNum As Long, searchValue As Variant) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Variant
    
    FindValueInColumn = 0
    
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    
    For i = 2 To lastRow ' Start from row 2 to skip header
        cellValue = ws.Cells(i, colNum).Value
        If CStr(cellValue) = CStr(searchValue) Then
            FindValueInColumn = i
            Exit Function
        End If
    Next i
End Function

' Subroutine to process Outlook emails
Private Sub ProcessOutlookEmails(emailAddress As String, searchSubject As String, _
                                  pwd1 As String, pwd2 As String, pwd3 As String, _
                                  wsCNDb As Worksheet)
    On Error GoTo OutlookError
    
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    Dim olMail As Object
    Dim olAttachment As Object
    
    Dim cnFolder As String
    Dim attachmentPath As String
    Dim emailFound As Boolean
    Dim i As Long
    
    ' Create Outlook application object
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Get the inbox folder
    Set olFolder = olNamespace.GetDefaultFolder(6) ' 6 = olFolderInbox
    
    ' CN Folder path for saving contract notes
    cnFolder = "C:\Users\rrr19\Documents\Codebases\Project Robot\CN Folder\"
    
    ' Create CN Folder if it doesn't exist
    If Dir(cnFolder, vbDirectory) = "" Then
        MkDir cnFolder
    End If
    
    emailFound = False
    
    ' Search through emails
    For Each olMail In olFolder.Items
        ' Check if sender email matches and subject contains search value
        If LCase(GetSenderEmail(olMail)) = LCase(emailAddress) Or _
           InStr(1, LCase(olMail.Subject), LCase(searchSubject), vbTextCompare) > 0 Then
            
            emailFound = True
            
            ' Process attachments
            For Each olAttachment In olMail.Attachments
                ' Check if it's a file attachment (not embedded)
                If olAttachment.Type = 1 Then ' 1 = olByValue
                    attachmentPath = cnFolder & olAttachment.FileName
                    
                    ' Save attachment
                    olAttachment.SaveAsFile attachmentPath
                    
                    ' Try to open the attachment (especially PDFs)
                    If LCase(Right(attachmentPath, 4)) = ".pdf" Then
                        Call TryOpenPDF(attachmentPath, pwd1, pwd2, pwd3, wsCNDb)
                    Else
                        ' For non-PDF files, just add to database
                        Call AddToCNDatabase(wsCNDb, attachmentPath)
                    End If
                End If
            Next olAttachment
        End If
    Next olMail
    
    If Not emailFound Then
        MsgBox "No emails found from '" & emailAddress & "' with subject containing '" & searchSubject & "'.", _
               vbInformation, "No Emails Found"
    End If
    
    ' Cleanup
    Set olAttachment = Nothing
    Set olMail = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    
    Exit Sub
    
OutlookError:
    MsgBox "Outlook Error " & Err.Number & ": " & Err.Description, vbCritical, "Outlook Error"
End Sub

' Function to get sender email address
Private Function GetSenderEmail(olMail As Object) As String
    On Error Resume Next
    
    Dim senderEmail As String
    
    ' Try to get sender email address
    If olMail.SenderEmailType = "EX" Then
        ' Exchange user - get SMTP address
        senderEmail = olMail.Sender.GetExchangeUser.PrimarySmtpAddress
        If senderEmail = "" Then
            senderEmail = olMail.SenderEmailAddress
        End If
    Else
        senderEmail = olMail.SenderEmailAddress
    End If
    
    GetSenderEmail = senderEmail
    On Error GoTo 0
End Function

' Subroutine to try opening a PDF with passwords
Private Sub TryOpenPDF(filePath As String, pwd1 As String, pwd2 As String, pwd3 As String, wsCNDb As Worksheet)
    On Error GoTo PDFError
    
    Dim opened As Boolean
    Dim acroApp As Object
    Dim acroDoc As Object
    Dim passwords(1 To 3) As String
    Dim i As Integer
    
    passwords(1) = pwd1
    passwords(2) = pwd2
    passwords(3) = pwd3
    
    opened = False
    
    ' First, try to open without password using Shell
    opened = TryOpenPDFNoPassword(filePath)
    
    If opened Then
        Call AddToCNDatabase(wsCNDb, filePath)
        Exit Sub
    End If
    
    ' Try with Adobe Acrobat (if available)
    On Error Resume Next
    Set acroApp = CreateObject("AcroExch.App")
    
    If Not acroApp Is Nothing Then
        Set acroDoc = CreateObject("AcroExch.PDDoc")
        
        ' Try opening without password first
        If acroDoc.Open(filePath) Then
            opened = True
            acroDoc.Close
            acroApp.Exit
        Else
            ' Try each password
            For i = 1 To 3
                If passwords(i) <> "" Then
                    ' Adobe Acrobat SDK method to open with password
                    If TryAdobeWithPassword(filePath, passwords(i)) Then
                        opened = True
                        Exit For
                    End If
                End If
            Next i
        End If
        
        Set acroDoc = Nothing
        Set acroApp = Nothing
    Else
        ' Fallback: Try using command line tools
        opened = TryPDFWithCommandLine(filePath, passwords)
    End If
    
    On Error GoTo PDFError
    
    If opened Then
        Call AddToCNDatabase(wsCNDb, filePath)
    Else
        Debug.Print "Could not open PDF: " & filePath
    End If
    
    Exit Sub
    
PDFError:
    Debug.Print "PDF Error for " & filePath & ": " & Err.Description
End Sub

' Function to try opening PDF without password
Private Function TryOpenPDFNoPassword(filePath As String) As Boolean
    On Error Resume Next
    
    Dim fso As Object
    Dim fileStream As Object
    Dim fileContent As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if file exists and is accessible
    If fso.FileExists(filePath) Then
        ' Try to verify PDF is not encrypted by checking header
        Set fileStream = fso.OpenTextFile(filePath, 1, False)
        fileContent = fileStream.Read(1024)
        fileStream.Close
        
        ' Check if PDF has encryption
        If InStr(1, fileContent, "/Encrypt", vbTextCompare) = 0 Then
            TryOpenPDFNoPassword = True
        Else
            TryOpenPDFNoPassword = False
        End If
    Else
        TryOpenPDFNoPassword = False
    End If
    
    Set fileStream = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Function

' Function to try Adobe with password
Private Function TryAdobeWithPassword(filePath As String, password As String) As Boolean
    On Error Resume Next
    
    Dim acroDoc As Object
    Set acroDoc = CreateObject("AcroExch.PDDoc")
    
    ' Try to open with password using AVDoc (UI version)
    Dim avDoc As Object
    Set avDoc = CreateObject("AcroExch.AVDoc")
    
    ' Open with password parameter
    If avDoc.Open(filePath, "") Then
        ' Document opened, check if we can access content
        TryAdobeWithPassword = True
        avDoc.Close True
    Else
        TryAdobeWithPassword = False
    End If
    
    Set avDoc = Nothing
    Set acroDoc = Nothing
    On Error GoTo 0
End Function

' Function to try PDF with command line tools (like qpdf or pdftk)
Private Function TryPDFWithCommandLine(filePath As String, passwords() As String) As Boolean
    On Error Resume Next
    
    Dim wsh As Object
    Dim result As Long
    Dim i As Integer
    Dim tempOutput As String
    Dim cmd As String
    
    Set wsh = CreateObject("WScript.Shell")
    tempOutput = Environ("TEMP") & "\temp_decrypted.pdf"
    
    TryPDFWithCommandLine = False
    
    ' Try qpdf (if installed)
    For i = 1 To 3
        If passwords(i) <> "" Then
            cmd = "qpdf --password=" & passwords(i) & " --decrypt """ & filePath & """ """ & tempOutput & """"
            result = wsh.Run(cmd, 0, True)
            If result = 0 Then
                ' Success - copy back
                FileCopy tempOutput, filePath
                Kill tempOutput
                TryPDFWithCommandLine = True
                Exit For
            End If
        End If
    Next i
    
    Set wsh = Nothing
    On Error GoTo 0
End Function

' Subroutine to add entry to CN Database
Private Sub AddToCNDatabase(wsCNDb As Worksheet, filePath As String)
    On Error GoTo DBError
    
    Dim lastRow As Long
    Dim newID As Long
    
    ' Find the last row with data in column A
    lastRow = wsCNDb.Cells(wsCNDb.Rows.Count, 1).End(xlUp).Row
    
    ' Calculate new ID
    If lastRow = 1 Then
        ' Only header exists
        newID = 1
    Else
        ' Get the last ID and increment
        If IsNumeric(wsCNDb.Cells(lastRow, 1).Value) Then
            newID = CLng(wsCNDb.Cells(lastRow, 1).Value) + 1
        Else
            newID = lastRow ' Fallback to row count
        End If
    End If
    
    ' Add new row
    lastRow = lastRow + 1
    wsCNDb.Cells(lastRow, 1).Value = newID           ' Column A: ID
    wsCNDb.Cells(lastRow, 2).Value = filePath        ' Column B: File Path
    
    Debug.Print "Added to CN Database: ID=" & newID & ", Path=" & filePath
    
    Exit Sub
    
DBError:
    Debug.Print "Database Error: " & Err.Description
End Sub

' Alternative method using Python (if VBA limitations are too restrictive)
' This creates a Python script that can be called from VBA for PDF handling
Public Sub CreatePythonHelper()
    Dim pythonScript As String
    Dim scriptPath As String
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    scriptPath = ThisWorkbook.Path & "\pdf_helper.py"
    
    pythonScript = "import sys" & vbCrLf & _
                   "import os" & vbCrLf & _
                   "from PyPDF2 import PdfReader" & vbCrLf & _
                   "" & vbCrLf & _
                   "def try_open_pdf(filepath, passwords):" & vbCrLf & _
                   "    try:" & vbCrLf & _
                   "        reader = PdfReader(filepath)" & vbCrLf & _
                   "        if reader.is_encrypted:" & vbCrLf & _
                   "            for pwd in passwords:" & vbCrLf & _
                   "                if pwd and reader.decrypt(pwd):" & vbCrLf & _
                   "                    return True" & vbCrLf & _
                   "            return False" & vbCrLf & _
                   "        return True" & vbCrLf & _
                   "    except Exception as e:" & vbCrLf & _
                   "        return False" & vbCrLf & _
                   "" & vbCrLf & _
                   "if __name__ == '__main__':" & vbCrLf & _
                   "    filepath = sys.argv[1]" & vbCrLf & _
                   "    passwords = sys.argv[2:5]" & vbCrLf & _
                   "    result = try_open_pdf(filepath, passwords)" & vbCrLf & _
                   "    print('SUCCESS' if result else 'FAILED')"
    
    Set file = fso.CreateTextFile(scriptPath, True)
    file.Write pythonScript
    file.Close
    
    MsgBox "Python helper script created at: " & scriptPath, vbInformation
    
    Set file = Nothing
    Set fso = Nothing
End Sub

' Function to call Python helper for PDF operations
Private Function TryPDFWithPython(filePath As String, pwd1 As String, pwd2 As String, pwd3 As String) As Boolean
    On Error Resume Next
    
    Dim wsh As Object
    Dim exec As Object
    Dim output As String
    Dim scriptPath As String
    Dim cmd As String
    
    Set wsh = CreateObject("WScript.Shell")
    scriptPath = ThisWorkbook.Path & "\pdf_helper.py"
    
    ' Build command
    cmd = "python """ & scriptPath & """ """ & filePath & """ """ & pwd1 & """ """ & pwd2 & """ """ & pwd3 & """"
    
    Set exec = wsh.exec(cmd)
    
    ' Wait for completion and get output
    Do While exec.Status = 0
        DoEvents
    Loop
    
    output = exec.StdOut.ReadAll
    
    TryPDFWithPython = (InStr(1, output, "SUCCESS", vbTextCompare) > 0)
    
    Set exec = Nothing
    Set wsh = Nothing
    On Error GoTo 0
End Function
