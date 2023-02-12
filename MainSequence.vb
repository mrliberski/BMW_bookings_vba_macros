'Variables will fetch data from 'ThisWorkbook' on startup
Public AccessPath
Public EmailToList As String
Public EmailCCList As String
Public sPath As String

'// Attempt to grab Outlook default signature for email footer I am using two different functions simultanously because had some issues with different configs in past
'// If one fails the other one normally works
Function GetEmailSignature()
    If Dir(sPath) <> "" Then
        StrSignature = GetSignature(sPath)
    Else
        StrSignature = ""
    End If
End Function

'// Attempt to grab Outlook default signature for email footer
Function GetSignature(FPath As String) As String
    Dim Fso As Object
    Dim TSet As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set TSet = Fso.GetFile(FPath).OpenAsTextStream(1, -2)
    GetSignature = TSet.ReadAll
    TSet.Close
End Function

'repetitive part of code to create email in Outllok
Function CreateEmail()
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

End Function