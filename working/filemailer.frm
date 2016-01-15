VERSION 5.00
Object = "{53337413-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "fileml50.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3360
      Top             =   6480
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   350
      Left            =   4080
      TabIndex        =   8
      Top             =   700
      Width           =   1455
   End
   Begin VB.TextBox txtText 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   5415
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   310
      Left            =   4200
      TabIndex        =   6
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ListBox lstAttachedFiles 
      Height          =   645
      ItemData        =   "filemailer.frx":0000
      Left            =   1200
      List            =   "filemailer.frx":0002
      TabIndex        =   5
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   310
      Left            =   4200
      TabIndex        =   7
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
   End
   Begin FILEMAILERLibCtl.FileMailer FileMailer1 
      Left            =   1920
      Top             =   6360
      AttachmentCount =   0
      BCc             =   ""
      Cc              =   ""
      Date            =   ""
      FirewallHost    =   ""
      FirewallPassword=   ""
      FirewallPort    =   80
      FirewallType    =   0
      FirewallUser    =   ""
      From            =   ""
      MailServer      =   ""
      MessageDate     =   ""
      MessageText     =   ""
      ReplyTo         =   ""
      SendTo          =   ""
      Subject         =   ""
      Timeout         =   0
      WinsockLoaded   =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Mail Server:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Label l1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attachments:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Subject:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "To:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "This is a demo of the FileMailer control.  Input text in the appropriate fields and click on the 'Send' button."
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mailServer As String
Public sendTo As String
Public from As String
Public subject As String
Public messagetext As String
Public attachmentFilePath As String
'
'
'
Public Sub send()
    FileMailer1.mailServer = mailServer
    FileMailer1.sendTo = sendTo
    FileMailer1.from = from
    FileMailer1.subject = subject
    FileMailer1.messagetext = messagetext
    FileMailer1.AttachmentCount = 1
    FileMailer1.Attachments(1) = attachmentFilePath
'    FileMailer1.mailServer = "s1.kma.net"
'    FileMailer1.sendTo = "jay@contensive.com"
'    'FileMailer1.SendTo = "5712913472@rcfax.com"
'    FileMailer1.from = "billing@contensive.com"
'    FileMailer1.subject = "send from mailer object"
'    FileMailer1.messagetext = "This is the text of the email"
'    FileMailer1.AttachmentCount = 1
'    FileMailer1.Attachments(1) = "c:\temp\test2.html"
    FileMailer1.send
End Sub
