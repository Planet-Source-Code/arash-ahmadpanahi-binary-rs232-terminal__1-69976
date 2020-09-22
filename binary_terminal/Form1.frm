VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ktbit.com"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7140
      TabIndex        =   14
      Top             =   7440
      Width           =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10545
      Top             =   8070
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OPEN PORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8145
      TabIndex        =   9
      Top             =   7440
      Width           =   1560
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   9765
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   7590
      Width           =   1050
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":08CA
      Left            =   11700
      List            =   "Form1.frx":08EF
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   7500
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "RTS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10905
      TabIndex        =   5
      Top             =   7650
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "DTR"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10890
      TabIndex        =   4
      Top             =   7410
      Value           =   1  'Checked
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   9240
      Top             =   7980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Export"
      Filter          =   "Text File|*.TXT"
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6120
      TabIndex        =   3
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   2
      Top             =   7425
      Width           =   975
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12726
      _Version        =   393217
      BackColor       =   16711680
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0937
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   7215
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12726
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":09C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9900
      Top             =   7965
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   19200
      SThreshold      =   1
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "www.ktbit.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Left            =   165
      TabIndex        =   13
      Top             =   8010
      Width           =   12630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DSRHolding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   2880
      TabIndex        =   12
      Top             =   7560
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTSHolding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   1275
      TabIndex        =   11
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CDHolding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   4635
      TabIndex        =   10
      Top             =   7560
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Commport"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9765
      TabIndex        =   8
      Top             =   7395
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
MSComm1.DTREnable = IIf(Check1.Value = 1, True, False)
End Sub

Private Sub Check2_Click()
MSComm1.RTSEnable = IIf(Check2.Value = 1, True, False)
End Sub

Private Sub Combo1_Click()
MSComm1.Settings = Combo1.Text & "N,8,1"
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Command1.Caption = "OPEN PORT" Then
    A = Replace(Combo2.Text, "COM", "")
    MSComm1.CommPort = Val(A)
    MSComm1.PortOpen = True
    Timer1.Enabled = True
    If Err Then MsgBox Error(Err): Exit Sub
Else
    MSComm1.PortOpen = False
    Timer1.Enabled = False
    If Err Then MsgBox Error(Err): Exit Sub
End If
Command1.Caption = IIf(Command1.Caption = "OPEN PORT", "CLOSE PORT", "OPEN PORT")
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
End Sub

Private Sub Command3_Click()
On Error Resume Next
CMD.ShowSave
If Err Then Exit Sub
Text1.SaveFile CMD.FileName, 1
End Sub

Private Sub Command4_Click()
On Error Resume Next
CMD.ShowSave
If Err Then Exit Sub
Text2.SaveFile CMD.FileName, 1
End Sub

Private Sub Form_Load()
For I = 1 To 100
Combo2.AddItem "COM" & Trim(Str(I))
Next
Combo2.ListIndex = 0
'MSComm1.PortOpen = True
Combo1.ListIndex = 7
End Sub

Private Sub MSComm1_OnComm()
Static J, X
If MSComm1.CommEvent = comEvReceive Then
    A = MSComm1.Input
    Text1.SelColor = vbWhite
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = A
    Text1.SelStart = Len(Text1.Text)
    
    For I = 1 To Len(A)
        b = Mid(A, I, 1)
        'Debug.Print IIf(Len(Hex(Asc(B))) = 1, "0" & Hex(Asc(B)), Hex(Asc(B))); " ";
        M = IIf(Len(Hex(Asc(b))) = 1, "0" & Hex(Asc(b)), Hex(Asc(b))) & " "
        Text2.SelColor = vbWhite
        Text2.SelStart = Len(Text2.Text)
        Text2.SelText = M
        X = X & IIf(Asc(b) < 32, ".", b)
        J = J + 1
        If J = 16 Then
            J = 0
            Text2.SelStart = Len(Text2.Text)
            Text2.SelText = ": " & X & vbCrLf: X = ""
        End If
    Next
    
    Text2.SelStart = Len(Text2.Text)
    If A = vbCr Then Debug.Print
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
MSComm1.Output = Chr(KeyAscii)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If MSComm1.CDHolding = True Then
    Label2.ForeColor = RGB(255, 255, 0)
    
Else
    Label2.ForeColor = RGB(128, 128, 128)
End If
    
If MSComm1.CTSHolding = True Then
    Label3.ForeColor = RGB(255, 255, 0)
Else
    Label3.ForeColor = RGB(128, 128, 128)
End If

If MSComm1.DSRHolding = True Then
    Label4.ForeColor = RGB(255, 255, 0)
Else
    Label4.ForeColor = RGB(128, 128, 128)
End If
End Sub
