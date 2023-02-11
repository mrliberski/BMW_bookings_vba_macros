Sub CMP() 
'Generates email to CMP team to confirm SAP system was prepared for routine massive MRP run
'Because of standfard content it will not be displayed but will send out automatically (.send)

Dim messages
messages = "WARNING! " & vbCrLf & vbCrLf & _
            "This script will notify CMP team that the MRP for plant Redditch has been ran." & vbCrLf & vbCrLf & _
            "Are you sure you want to proceed?"
    
answer = MsgBox(messages, vbCritical, vbYesNo, "Warning")

If answer = vbYes Then

    Dim OlApp As Object
    Dim NewMail As Object
    Dim EmailBody As String
    Dim StrSignature As String
    Dim sPath As String
    Dim strString
    Dim strResult
    Dim seqLead As Integer
    Dim Address As String
    Dim Fso As Object
    Dim fsFolder As Object
    Dim fsFile As Object

    Address = Application.ActiveWorkbook.FullName
    
    Set OlApp = CreateObject("Outlook.Application")
    Set NewMail = OlApp.CreateItem(0)
    With NewMail
        .display
    End With
    Signature = NewMail.htmlbody

'Email content
    Dim newBody
        newBody = " <html><head><style>body {color: #3d3d40;font-size:10pt;font-family:Calibri;}</style></head><body>"
        newBody = newBody + "<h4>RDD MRP&nbsp;</h4>"
        newBody = newBody & "Dear all, this is to confirm that MRP for Plant RDD has been ran and you are ok to schedule.<br> Thank you.<br> "
        newBody = newBody & "This is automated message: " & Now()

'attempt to grab a default signature
 GetEmailSignature()

    With NewMail
        .importance = 2
        .To = "UK_MaterialPlanning@grupoantolin.com"
        .cc = "pawel.liberski@grupoantolin.com"
        .BCC = ""
        .Subject = "RDD MRP RUN Confirmation - " & Date
        .htmlbody = newBody & "</BODY>" & vbNewLine & Signature
        .display
        .send
    End With
    
    Set NewMail = Nothing
    Set OlApp = Nothing

End If

End Sub

Sub SendAllPlans()
' This email will be send out to all internal recipients. 
' Macro will fetch all pdf files containing shipping plan tables, saved in previous steps

    Dim OlApp As Object
    Dim NewMail As Object
    Dim EmailBody As String
    Dim StrSignature As String
    Dim sPath As String
    Dim strString
    Dim strResult
    Dim seqLead As Integer
    Dim Address As String
    Dim Fso As Object
    Dim fsFolder As Object
    Dim fsFile As Object

    Address = Application.ActiveWorkbook.FullName
    
    Set OlApp = CreateObject("Outlook.Application")
    Set NewMail = OlApp.CreateItem(0)
        With NewMail
            .display
        End With
    Signature = NewMail.htmlbody

    Dim newBody
        newBody = " <html><head><style>body {color: #3d3d40;font-size:10pt;font-family:Calibri;}</style></head><body>"
        newBody = newBody + "<h3>Redditch Shipping Plan Updated&nbsp;</h3>"
        newBody = newBody & "<h4>" & Now() & "</h4>"
        'newBody = newBody & "Generated on " & Now()

'attempt to grab a default signature
 GetEmailSignature()

'define where attachments are saved
    Dim FileName As Variant
        FileName = Dir(AccessPath & "\")
        Debug.Print "filename: " & FileName

'generate email and add all atachments available
    With NewMail
        .importance = 2
        .To = EmailToList
        .cc = EmailCCList
        .Subject = "Redditch Shipping Plans Update " & Date
            While FileName <> ""
                Debug.Print "Attempting to add " & AccessPath & "\" & FileName
                .Attachments.Add (AccessPath & "\" & FileName)
                FileName = Dir
            Wend
                Debug.Print "Looks like files were added OK"
        .htmlbody = newBody & "</BODY>" & vbNewLine & Signature
        .display
    End With
    
'clean up
    Set NewMail = Nothing
    Set OlApp = Nothing
        Debug.Print "Mail was displayed"
End Sub
