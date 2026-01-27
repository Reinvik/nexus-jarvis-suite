' SAP Automation Bot
' Checks for negative stock in LX02 and created documents in the monitor.
' Sends email via Outlook if conditions are met.

Option Explicit

Dim SapGuiAuto, application, connection, session
Dim WScriptObj
Dim foundNegatives, foundDocuments
Dim logMessage

foundNegatives = False
foundDocuments = False
logMessage = ""

' ==========================================
' 1. Connect to SAP
' ==========================================
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    MsgBox "Could not attach to SAP GUI. Please ensure SAP is open and logged in.", vbCritical, "SAP Bot Error"
    WScript.Quit
End If
On Error GoTo 0

Set application = SapGuiAuto.GetScriptingEngine
If Not IsObject(connection) Then Set connection = application.Children(0)
If Not IsObject(session) Then Set session = connection.Children(0)

If IsObject(WScript) Then
    WScript.ConnectObject session,     "on"
    WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize

' ==========================================
' 2. Check LX02 for Negative Stock (920)
' ==========================================
Function CheckNegatives()
    On Error Resume Next
    
    ' Navigate to LX02
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nLX02" ' Use /n to ensure new transaction
    session.findById("wnd[0]").sendVKey 0
    
    ' Set Filters
    session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").text = "920"
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "SGSJ"
    
    ' Execute
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Check Grid Results
    Dim grid, rowCount, i, stockVal
    Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell")
    
    If Not grid Is Nothing Then
        rowCount = grid.RowCount
        If rowCount > 0 Then
            ' Iterate rows to check for negatives
            ' Assuming "VERME" (Available) or "GESME" (Total) is the column to check.
            ' The user didn't specify, but usually negative stock is checked in available or total.
            ' We will check "VERME" as per the user's column selection list which included it.
            
            For i = 0 To rowCount - 1
                stockVal = grid.GetCellValue(i, "VERME")
                If IsNumeric(stockVal) Then
                    If CDbl(stockVal) < 0 Then
                        CheckNegatives = True
                        logMessage = logMessage & "Found negative stock in row " & (i+1) & ": " & stockVal & vbCrLf
                        Exit Function ' Found one, that's enough to trigger
                    End If
                End If
            Next
        End If
    End If
    
    CheckNegatives = False
    On Error GoTo 0
End Function

' ==========================================
' 3. Check Document Monitor
' ==========================================
Function CheckDocuments()
    On Error Resume Next
    
    ' Navigate to main menu to ensure clean state for tree navigation
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    ' Execute Tree Navigation (User Snippet)
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00021"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00022"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00026"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "Favo"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00026"
    
    ' Enter parameters
    session.findById("wnd[0]/usr/txtP_IVNUM-LOW").text = "23802"
    session.findById("wnd[0]/usr/txtP_IVNUM-HIGH").text = "999999999"
    
    ' Press Execute
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' ---------------------------------------------------------
    ' VALIDATION LOGIC:
    ' 1. Check Status Bar for "No existen datos" (Image 0 from user)
    ' 2. Check if we are on the result screen (Image 1 from user)
    ' ---------------------------------------------------------
    
    Dim sbarText
    sbarText = ""
    On Error Resume Next
    sbarText = session.findById("wnd[0]/sbar").text
    On Error GoTo 0
    
    If InStr(1, sbarText, "No existen datos", 1) > 0 Or InStr(1, sbarText, "No data", 1) > 0 Then
        ' Explicit "No data" message found
        CheckDocuments = False
        logMessage = logMessage & "Monitor: No documents found (Status bar: " & sbarText & ")." & vbCrLf
    Else
        ' No error message. Check if we are still on the selection screen.
        ' If we are on the selection screen and no error, it might be a different issue, but usually "No data" appears.
        ' Let's try to find an element that ONLY exists on the result screen (e.g. the grid/list).
        ' In Image 1, it looks like a list. We don't have the ID, but we can check if the selection field is GONE.
        
        Dim selectionField
        Set selectionField = Nothing
        On Error Resume Next
        Set selectionField = session.findById("wnd[0]/usr/txtP_IVNUM-LOW")
        On Error GoTo 0
        
        If Not selectionField Is Nothing Then
            ' Still on selection screen, and maybe missed the status bar message?
            ' Assume False to be safe if we haven't moved.
            CheckDocuments = False
            logMessage = logMessage & "Monitor: No documents found (Remained on selection screen)." & vbCrLf
        Else
            ' Selection field is gone, so we must be on the result screen.
            CheckDocuments = True
            logMessage = logMessage & "Monitor: Documents found (Screen transition detected)." & vbCrLf
        End If
    End If

End Function

' ==========================================
' 4. Send Email
' ==========================================
Sub SendEmail(subject, bodyContent)
    Dim OutlookApp, MailItem
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0) ' 0 = olMailItem
    
    ' Construct HTML Body with Nexus Jarvis System styling
    ' Using HTML Entities for special characters to avoid encoding issues
    Dim htmlBody
    htmlBody = "<html><head><style>" & _
               "body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; background-color: #f4f4f4; padding: 20px; }" & _
               ".container { background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); max-width: 600px; margin: 0 auto; overflow: hidden; }" & _
               ".header { background-color: #2c3e50; color: #ffffff; padding: 20px; text-align: center; }" & _
               ".header h1 { margin: 0; font-size: 24px; letter-spacing: 1px; }" & _
               ".content { padding: 30px; }" & _
               ".alert-box { background-color: #e74c3c; color: #ffffff; padding: 15px; border-radius: 4px; margin-bottom: 20px; }" & _
               ".alert-title { font-weight: bold; font-size: 18px; margin-bottom: 5px; display: block; }" & _
               ".details { background-color: #ecf0f1; padding: 15px; border-radius: 4px; font-family: Consolas, monospace; font-size: 12px; color: #555; }" & _
               ".footer { background-color: #34495e; color: #ecf0f1; text-align: center; padding: 10px; font-size: 12px; }" & _
               "</style></head><body>" & _
               "<div class='container'>" & _
               "<div class='header'><h1>NEXUS JARVIS SYSTEM</h1></div>" & _
               "<div class='content'>" & _
               "<div class='alert-box'>" & _
               "<span class='alert-title'>Alerta de Control de Existencias</span>" & _
               bodyContent & _
               "</div>" & _
               "<p>Se requiere su atenci&oacute;n inmediata en las situaciones detectadas.</p>" & _
               "<div class='details'><strong>Log T&eacute;cnico:</strong><br/>" & Replace(logMessage, vbCrLf, "<br/>") & "</div>" & _
               "</div>" & _
               "<div class='footer'>Generado autom&aacute;ticamente por Nexus Jarvis System | CIAL Alimentos</div>" & _
               "</div></body></html>"

    With MailItem
        .To = "controldeexistencias@cial.cl"
        .Subject = subject
        .HTMLBody = htmlBody
        .Send
    End With
    
    Set MailItem = Nothing
    Set OutlookApp = Nothing
End Sub

' ==========================================
' Main Execution
' ==========================================

' Run LX02 Check
If CheckNegatives() Then
    foundNegatives = True
End If

' Run Document Monitor Check
If CheckDocuments() Then
    foundDocuments = True
End If

' Send Email if needed
If foundNegatives Or foundDocuments Then
    Dim emailSubject, emailContent
    emailSubject = "[NEXUS] Alerta de Stock/Documentos SAP"
    emailContent = ""
    
    If foundNegatives Then
        emailContent = emailContent & "<p><strong>&bull; Stock Negativo Detectado:</strong> Existen materiales con stock negativo en el almac&eacute;n 920.</p>"
    End If
    
    If foundDocuments Then
        emailContent = emailContent & "<p><strong>&bull; Documentos Pendientes:</strong> Se han encontrado documentos creados en el monitor.</p>"
    End If
    
    Call SendEmail(emailSubject, emailContent)
    MsgBox "Alerta enviada correctamente via Nexus Jarvis System.", vbInformation, "Nexus Jarvis System"
Else
    MsgBox "Sistema verificado: Sin novedades.", vbInformation, "Nexus Jarvis System"
End If
