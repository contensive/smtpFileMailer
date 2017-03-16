
Option Strict On
Option Explicit On

Imports System.Net
Imports System.Net.Mail


Namespace Contensive.SmtpFileMailerDotNet
    Public Class mainClass
        Public Sub send(toAddress As String, fromAddress As String, subject As String, body As String, smtpServer As String, attachmentFilename As String)
            Dim message As New MailMessage(fromAddress, toAddress, subject, body)
            Dim smtpClient As New SmtpClient(smtpServer)
            Dim attachment As New Attachment(attachmentFilename)
            message.Attachments.Add(attachment)
            'smtpClient.Port = 587
            'smtpClient.Credentials = New System.Net.NetworkCredential("your mail@gmail.com", "your password")
            'smtpClient.EnableSsl = True
            smtpClient.Send(message)
        End Sub
    End Class
End Namespace