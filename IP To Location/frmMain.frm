VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Global Location"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   1860
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Locate Me"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   1860
      Width           =   3015
   End
   Begin VB.TextBox txtLocalHost 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1380
      Width           =   3015
   End
   Begin VB.TextBox txtDomain 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   780
      Width           =   3015
   End
   Begin VB.TextBox txtExtIP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   180
      Width           =   3015
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   6720
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Geolocation provided by IPligence Community Edition. http://www.ipligence.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   8
      Top             =   2520
      Width           =   6675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Global Location :"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Long Address :"
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "External IP Address :"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1995
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CancelRequest As Boolean

Private Sub Command1_Click()
    
    Dim sTmp As String: sTmp = Empty
    
    Command2.Enabled = True
    
    txtExtIP = ""
    txtDomain = ""
    txtLocalHost = ""
    txtDomain = ""
    
    Me.Caption = "Acquiering External IP Address."
    sTmp = Inet.OpenURL("http://www.ShowMyIP.com/xml", icString)
    
    Do
        DoEvents 'start a loop
        If CancelRequest = True Then Inet.Cancel: Exit Do
    Loop Until Not Inet.StillExecuting 'keep doing 'nothing' until the inet control has gathered all of the source code
    
    If sTmp <> "" And CancelRequest = False Then
        txtExtIP.Text = GetIP(sTmp) 'parse the IP from the source code, and place it into the appropriate textbox
    End If
    
    sTmp = Empty 'clean up

    If Trim(txtExtIP.Text) <> "" Then
        Me.Caption = "Locating..."
        txtDomain.Text = DotToLong(txtExtIP.Text)
        txtLocalHost.Text = LongToLocation(CDbl(txtDomain.Text))
        Me.Caption = "Global Location Acquired."
    Else
        Me.Caption = "IP Acquasition Failed."
    End If
    
    CancelRequest = False
    Command2.Enabled = False
    
End Sub


Private Sub Command2_Click()

    CancelRequest = True
    Command2.Enabled = False
    
End Sub
