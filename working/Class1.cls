VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "smtpFileMailerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Public mailServer As String
Public sendTo As String
Public from As String
Public subject As String
Public messagetext As String
Public attachmentFilePath As String


Public Sub send()
    Form1.Show
    Form1.mailServer = mailServer
    Form1.sendTo = sento
    Form1.from = from
    Form1.subject = subject
    Form1.messagetext = messagetext
    Form1.attachmentFilePath = attachmentFilePath
    Call Form1.send
    
End Sub
