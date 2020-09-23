VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrxFcan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scan Proxys"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   3045
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame4 
      Caption         =   "Control"
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   2760
      Width           =   3015
      Begin VB.CommandButton CmdScan 
         Caption         =   "Scan"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Port"
         Height          =   615
         Left            =   1800
         TabIndex        =   10
         Top             =   120
         Width           =   975
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Text            =   "80"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Shape ShapeLead 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillStyle       =   6  'Cross
         Height          =   615
         Left            =   120
         Shape           =   2  'Oval
         Top             =   185
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proxy File"
      Height          =   2655
      Left            =   0
      TabIndex        =   7
      Top             =   4200
      Width           =   3015
      Begin VB.ListBox lstIP 
         Height          =   2205
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proxys Found"
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   3015
      Begin VB.ListBox lstWin 
         Height          =   1425
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Wingate"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin MSWinsockLib.Winsock Sck1 
      Left            =   5040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFile 
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   165
   End
   Begin VB.TextBox TxtFile 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmPrxFcan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '##############################################################
    
            'I spent a long time fixing this code up and
           'nothing would make me feel better than to know
    'that people can make use of it and do better. Special thanks
          'to Kris & flandersucf who was a big help in to make
      'up the application! with out them none of this was working
               'This is Free source Kepp me in touch
                    'and send me new version :)
                '    will be happy happy thanks
                '        xvicx@hotmail.com
                
    '#############################################################
Dim GoIP As String
Dim X1, X2, i As Integer

Sub WinScan()
  X2 = 0
PB1.Max = lstIP.ListCount - 1
PB1.Value = 0

For i = 1 To lstIP.ListCount - 1
     Debug.Print X2 & "  :  " & lstIP.List(X2)
     PB1.Value = PB1.Value + 1
     Sck1.Connect lstIP.List(X2), txtPort
Do
     Select Case Sck1.State
         Case 7, 8, 9, 0
             Exit Do
     End Select
     DoEvents
Loop

If Sck1.State = 7 Then
   lstWin.AddItem lstIP.List(X2)
   ShapeLead.BackColor = &HFF00&
End If
   X2 = X2 + 1
   Sck1.Close
   ShapeLead.BackColor = &HFF&
    Next
    
Debug.Print "Wingate Done"
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFile_Click()
CD1.Filter = "All Files (*.*) | *.*"
CD1.ShowOpen
TxtFile.Text = CD1.FileName
End Sub

Private Sub CmdScan_Click()
On Error Resume Next
CmdScan.Enabled = False
TxtFile.Enabled = False

  lstIP.Clear  'clear the list
  X1 = -1
  X2 = -1
  
  
  Open TxtFile.Text For Input As #1  'open the file
  Do Until EOF(1) = True  'go until the end of the file
    Input #1, GoIP
    X1 = X1 + 1
      If GoIP = "" Then
      Else
        lstIP.AddItem GoIP, X1  'add all the lines into the lstip listbox
      End If
  Loop
  
  Close #1 'close the file


Call WinScan

TxtFile.Enabled = True
CmdScan.Enabled = True
End Sub

Private Sub Form_Load()
TxtFile = App.Path & "\Proxy.txt"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Sck1.Close
Unload Me
End
End Sub

Private Sub Sck1_Connect()
ShapeLead.BackColor = &HFF00&
End Sub


