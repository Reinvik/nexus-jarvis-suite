' SAP Automation Bot - SIMULATION MODE
' This script simulates finding negative stock and documents to test alerts.

Option Explicit

Dim SapGuiAuto, application, connection, session
Dim WScriptObj
Dim foundNegatives, foundDocuments
Dim logMessage

foundNegatives = False
foundDocuments = False
logMessage = ""

' ==========================================
' 1. Connect to SAP (Kept for realism)
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
    
    ' Navigate to LX02 (Real navigation)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nLX02"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").text = "920"
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "SGSJ"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' SIMULATION: Force finding negatives
    CheckNegatives = True
    logMessage = logMessage & "SIMULATION: Found negative stock in row 1: -50.00" & vbCrLf
    logMessage = logMessage & "SIMULATION: Found negative stock in row 5: -12.00" & vbCrLf
    
    On Error GoTo 0
End Function

' ==========================================
' 3. Check Document Monitor
' ==========================================
Function CheckDocuments()
    On Error Resume Next
    
    ' Navigate to Monitor (Real navigation)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00021"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00022"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00026"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "Favo"
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00026"
    session.findById("wnd[0]/usr/txtP_IVNUM-LOW").text = "23802"
    session.findById("wnd[0]/usr/txtP_IVNUM-HIGH").text = "999999999"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' SIMULATION: Force finding documents
    CheckDocuments = True
    logMessage = logMessage & "SIMULATION: Documents found in monitor (Forced)." & vbCrLf

End Function

' ==========================================
' 4. Send Email
' ==========================================
Sub SendEmail(subject, bodyContent)
    Dim OutlookApp, MailItem
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0) ' 0 = olMailItem
    
    ' Construct HTML Body with Nexus Jarvis System styling
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
               "<div class='header'><h1>NEXUS JARVIS SYSTEM (SIMULACRO)</h1></div>" & _
               "<div class='content'>" & _
               "<div class='alert-box'>" & _
               "<span class='alert-title'>Alerta de Control de Existencias (PRUEBA)</span>" & _
               bodyContent & _
               "</div>" & _
               "<p>Este es un correo de prueba generado por el modo de simulaci&oacute;n.</p>" & _
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
    emailSubject = "[NEXUS] SIMULACRO - Alerta de Stock/Documentos SAP"
    emailContent = ""
    
    If foundNegatives Then
        emailContent = emailContent & "<p><strong>&bull; Stock Negativo Detectado:</strong> Existen materiales con stock negativo en el almac&eacute;n 920.</p>"
    End If
    
    If foundDocuments Then
        emailContent = emailContent & "<p><strong>&bull; Documentos Pendientes:</strong> Se han encontrado documentos creados en el monitor.</p>"
    End If
    
    Call SendEmail(emailSubject, emailContent)
    MsgBox "Simulacro completado. Se ha enviado un correo de prueba.", vbInformation, "Nexus Jarvis Simulation"
Else
    MsgBox "Error en simulacro.", vbCritical, "Nexus Jarvis Simulation"
End If
