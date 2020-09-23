VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Console + HotTracking Example"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add More"
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2340
      Index           =   0
      Left            =   240
      ScaleHeight     =   2340
      ScaleWidth      =   4845
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   4845
      Begin VB.Frame fraSample1 
         Caption         =   "Sample 1"
         Height          =   1785
         Left            =   210
         TabIndex        =   10
         Top             =   255
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2340
      Index           =   1
      Left            =   240
      ScaleHeight     =   2340
      ScaleWidth      =   4845
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   4845
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   8
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2340
      Index           =   2
      Left            =   240
      ScaleHeight     =   2340
      ScaleWidth      =   4845
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   4845
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   6
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2340
      Index           =   3
      Left            =   240
      ScaleHeight     =   2340
      ScaleWidth      =   4845
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   4845
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.CheckBox chkHotTracking 
      Caption         =   "Hot Tracking"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   2805
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4948
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 1"
            Key             =   "Group1"
            Object.ToolTipText     =   "Set Options for Group 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 2"
            Key             =   "Group2"
            Object.ToolTipText     =   "Set Options for Group 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 3"
            Key             =   "Group3"
            Object.ToolTipText     =   "Set Options for Group 3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 4"
            Key             =   "Group4"
            Object.ToolTipText     =   "Set Options for Group 4"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=
'|| ::Author::          :  TonY Myers           ||
'|| ::Project Name::    :  Console Example      ||
'|| ::Complied          :  25/6/02              ||
'|| ::Created           :  24/6/02 - 01:54am    ||
'|| ::Comments          :  If this code has been||
'|| found on psc.om before i have posted,       ||
'|| sorry :)                                    ||
'||                                             ||
'|| ::NOTES             :    N/A                ||
'||                                             ||
'=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=

Option Explicit
Dim i As Integer
Dim Counter As Integer

Private Sub chkHotTracking_Click()
On Error Resume Next

If tbsOptions.HotTracking = True Then
tbsOptions.HotTracking = False
ElseIf tbsOptions.HotTracking = False Then
tbsOptions.HotTracking = True
End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next

Counter = Counter + 1
WriteConsole "Test " & Counter, &H8000000F
End Sub

Private Sub cmdQuit_Click()
On Error Resume Next

Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Counter = 0
WriteConsole "Welcome to this console example, if an example has been up loaded to psc.com, sorry for wasting your time :)", vbRed
End Sub

Private Sub tbsOptions_Click()
Dim i As Integer
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Visible = True
            picOptions(i).Enabled = True
        Else
            picOptions(i).Visible = False
            picOptions(i).Enabled = False
        End If
    Next
End Sub
