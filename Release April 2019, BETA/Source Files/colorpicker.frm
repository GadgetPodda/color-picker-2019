VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Color Picker 2019 - Apr,19-BETA"
   ClientHeight    =   1215
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7305
   Icon            =   "colorpicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox code 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Text            =   "#FFFFFF"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton choose 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Choose Colour"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub choose_Click()
cd.ShowColor
Dim clr As Long
clr = cd.color
code.Text = "#" + Format(Hex(CInt(&HFF& And CLng(clr))), "00") & Format(Hex(CInt((&HFF00& And CLng(clr)) \ 256)), "00") & Format(Hex(CInt((&HFF0000 And CLng(clr)) \ 65536)), "00")
color.BackColor = cd.color
End Sub

Private Sub MenuHelpAbout_Click()
MsgBox "Color Picker 2019 BETA. Contributions can be made on GitHub. Author Account : GadgetPodda (Github) . Support Programmer (Paypal) : gadgetpodda2005@gmail.com", vbOKOnly + vbInformation, "Color Picker 2019"
End Sub
