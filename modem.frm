VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODEM"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm1 
      Left            =   60
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   480
      Left            =   165
      TabIndex        =   2
      Top             =   4590
      Width           =   3900
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   3930
      Left            =   60
      TabIndex        =   0
      Top             =   585
      Width           =   4125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MODEM PROPERTIES"
      Height          =   450
      Left            =   30
      TabIndex        =   1
      Top             =   285
      Width           =   4050
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Hide Me
Form1.Show
End Sub

Private Sub Form_Load()
Show
MSComm1.CommPort = 3
MSComm1.PortOpen = True
MSComm1.Output = "ATI4" & vbCr
MSComm1.InBufferCount = 0
MSComm1.InputLen = 0
Do
Loop While MSComm1.InBufferCount <= 0
List1.AddItem ("Modem Info: " & MSComm1.Input)
MSComm1.Output = "ATI6" & vbCr
MSComm1.InBufferCount = 0
Do
Loop While MSComm1.InBufferCount <= 0
List1.AddItem ("Supported Standards: " & MSComm1.Input)
End Sub
