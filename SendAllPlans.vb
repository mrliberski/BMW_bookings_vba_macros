Sub SendAllPlans()
' This email will be send out to all internal recipients. 
' Macro will fetch all pdf files containing shipping plan tables, saved in previous steps

CreateEmail()

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