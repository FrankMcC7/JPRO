Option Explicit

' Declare WithEvents to trap NewInspector events
Private WithEvents insp As Outlook.Inspectors

' Initialize event handling at startup
Private Sub Application_Startup()
    Set insp = Application.Inspectors
End Sub

' Fires when a new compose or reply window opens
Private Sub insp_NewInspector(ByVal Inspector As Outlook.Inspector)
    If TypeOf Inspector.CurrentItem Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Inspector.CurrentItem
        If Left(mail.Subject, 9) <> "[SecMail:]" Then
            mail.Subject = "[SecMail:]" & mail.Subject
        End If
    End If
End Sub

' Fires just before sending any item
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    If TypeOf Item Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Item
        If Left(mail.Subject, 9) <> "[SecMail:]" Then
            mail.Subject = "[SecMail:]" & mail.Subject
        End If
    End If
End Sub