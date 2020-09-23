VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "What icons is loaded from Visual Basic?"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9780
      TabIndex        =   2
      Top             =   7710
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   5895
      Left            =   8100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmSample.frx":0000
      Top             =   30
      Width           =   2625
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1920
      Left            =   8100
      Picture         =   "frmSample.frx":0006
      ToolTipText     =   "This is the icons as appear on Explorer"
      Top             =   6165
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   $"frmSample.frx":0C08
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   8100
      Width           =   10710
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   5
      Left            =   6660
      Picture         =   "frmSample.frx":0C92
      ToolTipText     =   "A ""Invalid image"" error occur if try to load Wrong6 icon."
      Top             =   4230
      Width           =   270
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Index           =   4
      Left            =   5820
      Picture         =   "frmSample.frx":0DDC
      Top             =   6120
      Width           =   750
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Index           =   3
      Left            =   1290
      Picture         =   "frmSample.frx":578E
      Top             =   6120
      Width           =   750
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Index           =   2
      Left            =   1320
      Picture         =   "frmSample.frx":A6B8
      ToolTipText     =   "This is the image Really loaded"
      Top             =   4110
      Width           =   750
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   1
      Left            =   1290
      Picture         =   "frmSample.frx":FE9A
      ToolTipText     =   "This is the image Really loaded"
      Top             =   2130
      Width           =   510
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   0
      Left            =   1320
      Picture         =   "frmSample.frx":15AAC
      ToolTipText     =   "This is the image Really loaded"
      Top             =   120
      Width           =   510
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2025
      Index           =   5
      Left            =   5250
      Picture         =   "frmSample.frx":1B7FE
      Top             =   4050
      Width           =   2790
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2025
      Index           =   4
      Left            =   4380
      Picture         =   "frmSample.frx":1C7A0
      Top             =   6060
      Width           =   3660
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2025
      Index           =   3
      Left            =   30
      Picture         =   "frmSample.frx":1DB8E
      Top             =   6060
      Width           =   4530
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2025
      Index           =   2
      Left            =   30
      Picture         =   "frmSample.frx":1F1C1
      Top             =   4050
      Width           =   5400
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2025
      Index           =   1
      Left            =   30
      Picture         =   "frmSample.frx":20B84
      Top             =   2040
      Width           =   7140
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2025
      Index           =   0
      Left            =   30
      Picture         =   "frmSample.frx":22B55
      ToolTipText     =   "This is the starting image"
      Top             =   30
      Width           =   8010
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim s As String
    Label1.ForeColor = vbRed
    s = "<Ivalid image!>" & vbCrLf
    s = s & "This is a error that occurs if you try to load a 'very' 32bpp icon with Alpha channel." & vbCrLf & vbCrLf
    
    s = s & "Have you never asked yourselves: why?" & vbCrLf
    s = s & "VB does not support Alpha channel! Therefore the image loaded is dependent on the system and image formats embedded in the icon file." & vbCrLf & vbCrLf
    
    s = s & "This form shows what image is choosen by VB when the icon is loaded. You can see that this choice depends on 'which' image formats the icon file contains." & vbCrLf
    
    s = s & "Note: I have made all the WrongN images based on Wrong1. Next, I have deleted a single format, one by one." & vbCrLf
    s = s & "The image really loaded by VB is shown (black bordered) to the right of 'WrongN' name of each image." & vbCrLf
    s = s & "The last image Wrong6 contains a 32bpp Alpha channel images only. Therefore I have placed a red icon (X) to indicate that this image can't be loaded." & vbCrLf
    
    s = s & "Please note also that the icon files in Explorer window shown the correctly image."
    
    Text1.Text = s
    
End Sub


