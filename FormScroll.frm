VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Scrolling on Normal Form"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   285
      Left            =   4320
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   2760
      Width           =   285
   End
   Begin VB.HScrollBar hs 
      Height          =   240
      Left            =   4200
      TabIndex        =   11
      Top             =   1320
      Width           =   870
   End
   Begin VB.VScrollBar vs 
      Height          =   960
      Left            =   4320
      TabIndex        =   10
      Top             =   1680
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   -45
      ScaleHeight     =   4140
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   0
      Width           =   4740
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   3120
         Width           =   3525
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   195
         Left            =   3000
         TabIndex        =   8
         Top             =   2640
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   3000
         TabIndex        =   7
         Top             =   2880
         Width           =   1365
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   3900
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   2640
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   1365
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   3930
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   3930
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         MaxLength       =   16
         TabIndex        =   1
         Top             =   840
         Width           =   3885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Now Try Resizing the form to make the scroll bars visible"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Scrolling on Normal Form code by
'Kailash Nadh, 15 yrs, India
'kailashbn@satyam.net.in
'http://kbn.rom.cd
'For 7 super Softwares in VB less that 500 Kb,
'visit my site (3D Clock, Talking Calculator, Quiz master..
Option Explicit

Private Sub Form_Load()
'Assign Normal values to the scrollbars
vs.Value = 0
vs.LargeChange = 100
vs.SmallChange = 50
vs.Min = 0
'~~~~~~~~~~~~~~~~~~~
hs.Value = 0
hs.LargeChange = 100
hs.SmallChange = 50
hs.Min = 0
End Sub

Private Sub Form_Resize()
'Make changes to the scrollbar accordint to form's size
If Picture1.Height > Me.Height Then vs.Visible = True Else vs.Visible = False

If Picture1.Width > Me.Width Then hs.Visible = True Else hs.Visible = False

'Show/Hide the Small Picturebox
If hs.Visible = False And vs.Visible = False Then
Picture2.Visible = False
Else
Picture2.Visible = True
End If

hs.Width = Me.ScaleWidth: vs.Height = Me.ScaleHeight
hs.Top = Me.ScaleHeight - hs.Height
vs.Left = Me.ScaleWidth - vs.Width
hs.Left = 0
vs.Top = 0

Picture2.Top = Me.ScaleHeight - hs.Height
Picture2.Left = Me.ScaleWidth - vs.Width

vs.Height = vs.Height - Picture2.ScaleHeight
hs.Width = hs.Width - Picture2.ScaleWidth

hs.Max = (Picture1.Width - Me.Width) + vs.Width
vs.Max = (Picture1.Height - Me.Height) + hs.Height
End Sub

'scroll the form (picture box) when the scrollbars
'are dragged
Private Sub vs_Change()
Picture1.Top = -vs.Value
End Sub

Private Sub hs_Scroll()
Picture1.Left = -hs.Value
End Sub

Private Sub vs_Scroll()
Picture1.Top = -vs.Value
End Sub

Private Sub hs_Change()
Picture1.Left = -hs.Value
End Sub
