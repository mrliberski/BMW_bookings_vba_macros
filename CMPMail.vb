Sub CMP() 
'Generates email to CMP team to confirm SAP system was prepared for routine massive MRP run
'Because of standfard content it will not be displayed but will send out automatically (.send)

Dim messages
messages = "WARNING! " & vbCrLf & vbCrLf & _
            "This script will notify CMP team that the MRP for plant Redditch has been ran." & vbCrLf & vbCrLf & _
            "Are you sure you want to proceed?"
    
answer = MsgBox(messages, vbCritical, vbYesNo, "Warning")

If answer = vbYes Then

    CreateEmail()

    'Email body content
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