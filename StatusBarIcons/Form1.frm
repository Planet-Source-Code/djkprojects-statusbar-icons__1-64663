VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StatusBarIcons"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   926
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Visible         =   0   'False
            Text            =   "Panel 1"
            TextSave        =   "Panel 1"
            Key             =   "p1"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Visible         =   0   'False
            Text            =   "Panel 2"
            TextSave        =   "Panel 2"
            Key             =   "p2"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   0
            Text            =   "Panel 3"
            TextSave        =   "Panel 3"
            Key             =   "p3"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "http://djkprojects.webd.pl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   "house"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":031A
            Key             =   "butterfly"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0634
            Key             =   "people"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Copyright (C) 2006 DJK's Projects (DJK)

'Website:    http://djkprojects.webd.pl/
'Contact:    djkprojects@interia.pl

'Autor przyk³adu nie odpowiada za ewentualne szkody wywo³ane dzia³aniem
'poni¿szego kodu.

'Author of this example is not responsible for any damages caused by its use

Private imgSB As New Images

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()

Call imgSB.SetImageListIcons(ImageList1)

With StatusBar1
        .Style = sbrNormal
        .Panels("p1").Picture = ImageList1.ListImages("house").ExtractIcon
        .Panels("p1").Text = "First Panel"
        .Panels("p1").Visible = True

        .Panels("p2").Picture = ImageList1.ListImages("butterfly").ExtractIcon
        .Panels("p2").Text = "Second Panel"
        .Panels("p2").Visible = True

        .Panels("p3").Picture = ImageList1.ListImages("people").ExtractIcon
        .Panels("p3").Text = "Third Panel"
        .Panels("p3").Visible = True
End With

End Sub

Private Sub Label1_Click()
ShellExecute Me.hwnd, vbNullString, "http://djkprojects.webd.pl", vbNullString, "", 1
End Sub
