VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmPaint 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8550
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13995
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
   ScaleHeight     =   8550
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   1588
      ButtonWidth     =   1032
      ButtonHeight    =   1429
      Wrappable       =   0   'False
      ImageList       =   "ImgList"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cut"
            Key             =   "ImCut"
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            ImageKey        =   "ImCut"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Copy"
            Key             =   "ImCopy"
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            ImageKey        =   "ImCut"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Paste"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            ImageKey        =   "ImCut"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdSaffron 
      BackColor       =   &H00C0E0FF&
      Height          =   480
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdDarkRed 
      BackColor       =   &H00000080&
      Height          =   480
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdLightRed 
      BackColor       =   &H008080FF&
      Height          =   480
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdRed 
      BackColor       =   &H000000FF&
      Height          =   480
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdLightOrange 
      BackColor       =   &H0080C0FF&
      Height          =   480
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdLightPink 
      BackColor       =   &H00FF80FF&
      Height          =   480
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdLightYellow 
      BackColor       =   &H0080FFFF&
      Height          =   480
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdLightGreen 
      BackColor       =   &H0080FF80&
      Height          =   480
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdLightBlue 
      BackColor       =   &H00FFC0C0&
      Height          =   480
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdMoreColor 
      Caption         =   "More Color"
      Height          =   360
      Left            =   12840
      TabIndex        =   11
      Top             =   4200
      Width           =   990
   End
   Begin RichTextLib.RichTextBox RichTextBox 
      Height          =   6975
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   12303
      _Version        =   393217
      TextRTF         =   $"frmPaint.frx":0000
   End
   Begin VB.CommandButton cmdPurple 
      BackColor       =   &H00FF8080&
      Height          =   480
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdBlue 
      BackColor       =   &H00FF0000&
      Height          =   480
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdGreen 
      BackColor       =   &H0000FF00&
      Height          =   480
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdYellow 
      BackColor       =   &H0000FFFF&
      Height          =   480
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdPink 
      BackColor       =   &H00FF00FF&
      Height          =   480
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdOrange 
      BackColor       =   &H000080FF&
      Height          =   480
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdGrey 
      BackColor       =   &H8000000A&
      Height          =   480
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdWhite 
      BackColor       =   &H8000000E&
      Height          =   480
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1440
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBlack 
      BackColor       =   &H80000009&
      Caption         =   "Black"
      Height          =   480
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   0
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPaint.frx":007B
            Key             =   "ImCut"
            Object.Tag             =   "ImCut"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
   End
   Begin VB.Menu MenuHome 
      Caption         =   "Home"
   End
   Begin VB.Menu MenuView 
      Caption         =   "View"
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBlack_Click()
    RichTextBox.SelColor = cmdBlack.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdBlue_Click()
    RichTextBox.SelColor = cmdBlue.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdDarkRed_Click()
    RichTextBox.SelColor = cmdDarkRed.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdGreen_Click()
    RichTextBox.SelColor = cmdGreen.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdGrey_Click()
    RichTextBox.SelColor = cmdGrey.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdLightBlue_Click()
    RichTextBox.SelColor = cmdLightBlue.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdLightGreen_Click()
    RichTextBox.SelColor = cmdLightGreen.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdLightOrange_Click()
    RichTextBox.SelColor = cmdLightOrange.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdLightPink_Click()
    RichTextBox.SelColor = cmdLightPink.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdLightRed_Click()
    RichTextBox.SelColor = cmdLightRed.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdLightYellow_Click()
    RichTextBox.SelColor = cmdLightYellow.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdMoreColor_Click()
    CommonDialog.ShowColor
    RichTextBox.SelColor = CommonDialog.Color
    RichTextBox.SetFocus
End Sub

Private Sub cmdOrange_Click()
    RichTextBox.SelColor = cmdOrange.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdPink_Click()
    RichTextBox.SelColor = cmdPink.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdPurple_Click()
    RichTextBox.SelColor = cmdPurple.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdRed_Click()
    RichTextBox.SelColor = cmdRed.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdSaffron_Click()
    RichTextBox.SelColor = cmdSaffron.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdWhite_Click()
    RichTextBox.SelColor = cmdWhite.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub cmdYellow_Click()
    RichTextBox.SelColor = cmdYellow.BackColor
    RichTextBox.SetFocus
End Sub

Private Sub Form_Load()
    RichTextBox.SelFontSize = 20
End Sub

Private Sub MenuFile_Click()
    CommonDialog.ShowOpen
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    If Button = "Cut" Then
        Clipboard.Clear
        Clipboard.SetText RichTextBox.Text
        RichTextBox.Text = Empty
    ElseIf Button = "Copy" Then
        Clipboard.Clear
        Clipboard.SetText RichTextBox.Text
    ElseIf Button = "Paste" Then
        RichTextBox.Text = RichTextBox.Text & Clipboard.GetText
   End If
End Sub
