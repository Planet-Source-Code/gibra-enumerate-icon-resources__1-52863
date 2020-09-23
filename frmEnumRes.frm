VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEnumRes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enum Icon Resources selective - A very How To Use 32bpp icons (with Alpha channel)"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9045
   Icon            =   "frmEnumRes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   7725
      TabIndex        =   23
      Top             =   3630
      Width           =   1200
   End
   Begin VB.CheckBox chkAllSizeFormat 
      Caption         =   "All Size and Formats"
      Height          =   195
      Left            =   4110
      TabIndex        =   19
      Top             =   3150
      Width           =   1815
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   6450
      TabIndex        =   17
      Top             =   1665
      Width           =   2475
   End
   Begin VB.CommandButton cmdLoadIcons 
      Caption         =   "&Load Icons"
      Enabled         =   0   'False
      Height          =   405
      Left            =   6450
      TabIndex        =   22
      Top             =   3630
      Width           =   1200
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   240
      TabIndex        =   14
      Top             =   1365
      Width           =   3315
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   24
      Top             =   4200
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13335
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color depth"
      Height          =   1845
      Index           =   1
      Left            =   4810
      TabIndex        =   8
      Top             =   1140
      Width           =   1485
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1395
         Index           =   1
         Left            =   210
         ScaleHeight     =   1395
         ScaleWidth      =   1110
         TabIndex        =   9
         Top             =   300
         Width           =   1110
         Begin VB.OptionButton optType 
            Caption         =   "Win XP"
            Height          =   240
            Index           =   3
            Left            =   60
            TabIndex        =   13
            ToolTipText     =   "Windows XP (32bit-Alpha channel) image only"
            Top             =   1110
            Width           =   990
         End
         Begin VB.OptionButton optType 
            Caption         =   "TrueColor"
            Height          =   240
            Index           =   2
            Left            =   60
            TabIndex        =   12
            ToolTipText     =   "TrueColor (24bit) image only"
            Top             =   740
            Width           =   990
         End
         Begin VB.OptionButton optType 
            Caption         =   "256"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   11
            ToolTipText     =   "256 colors (16bit) image only"
            Top             =   370
            Width           =   750
         End
         Begin VB.OptionButton optType 
            Caption         =   "16"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   10
            ToolTipText     =   "16 colors (4bit) image only"
            Top             =   0
            Width           =   750
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Size"
      Height          =   1845
      Index           =   0
      Left            =   3710
      TabIndex        =   2
      Top             =   1140
      Width           =   945
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1395
         Index           =   0
         Left            =   210
         ScaleHeight     =   1395
         ScaleWidth      =   675
         TabIndex        =   3
         Top             =   300
         Width           =   675
         Begin VB.OptionButton optSize 
            Caption         =   "48"
            Height          =   240
            Index           =   3
            Left            =   60
            TabIndex        =   7
            ToolTipText     =   "48x48 format only"
            Top             =   1110
            Width           =   600
         End
         Begin VB.OptionButton optSize 
            Caption         =   "32"
            Height          =   240
            Index           =   2
            Left            =   60
            TabIndex        =   6
            ToolTipText     =   "32x32 format only"
            Top             =   735
            Width           =   600
         End
         Begin VB.OptionButton optSize 
            Caption         =   "24"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   5
            ToolTipText     =   "24x24 format only"
            Top             =   375
            Width           =   600
         End
         Begin VB.OptionButton optSize 
            Caption         =   "16"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   4
            ToolTipText     =   "16x16 format only"
            Top             =   0
            Width           =   600
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7890
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   240
      TabIndex        =   20
      Top             =   3390
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Please note that if choose All Size and Formats some images appear wrong because is draw stretched."
      ForeColor       =   &H8000000D&
      Height          =   585
      Index           =   3
      Left            =   3780
      TabIndex        =   21
      Top             =   3450
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Other formats (if any):"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   6480
      TabIndex        =   16
      Top             =   1140
      Width           =   1500
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Num  Size  w  h "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   270
      Index           =   1
      Left            =   6450
      TabIndex        =   15
      Top             =   1395
      Width           =   2475
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Images founds:"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   3150
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose a library file (EXE/DLL):"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1140
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   7410
      Top             =   3120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Menu mnuBar 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&Open Readme.txt"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&How VB load images from 32bpp Alpha channel icons"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuBar 
      Caption         =   "&Help"
      Index           =   1
      Begin VB.Menu mnuHelp 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmEnumRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : frmEnumRes.frm
' DateTime    : 03/04/2004 16.09
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Project     : EnumResource.vbp
' Purpose     : Load resources ICON from a EXE/DLL library
' Description : This project show how to load Windows XP (32bpp) icons format
'               from executable files.
' Comments    : When we try to load a 32bpp icon - Windows XP format with
'               Alpha Channel - to any graphic control (like Form, PictureBox,
'               Image, ...) VB return a error: "Invalid image."
'               Many icon files embedded more than one image format, like
'               below:
'               16x16 4bpp      16x16 16bpp     16x16 32bpp
'               32x32 4bpp      32x32 16bpp     32x32 32bpp
'               48x48 4bpp      48x48 16bpp     48x48 32bpp
'
'               By loading this icon file (with 9 fotmats) we don't get error
'               because VB choose itself automatically the format (32x32 with
'               16bpp is the more frequent, but is depend from system we use.)
'               Now, suppose that we have a icons library, o variuos icon files,
'               that embedded ONLY 32bpp images (that is Windows XP format and
'               Alpha Channel)?
'               VB refuse to load it! Therefore we can't use it.
'
'               WORKAROUND
'               Of course, with a lot of API functions we can obtain the
'               result. This source code show how to does it.
'
'               For more info about source code see the README.TXT.
'               However, source code is full commented, so you can know
'               better what is does.
'---------------------------------------------------------------------------------------
Option Explicit

Dim sLibraryFile As String  '/ complete Path and Filename

Private Sub chkAllSizeFormat_Click()
    gbAllSizeFormat = chkAllSizeFormat.Value
    
    Dim i As Integer
    
    If gbAllSizeFormat Then
        For i = 0 To 3
            optSize(i).Value = False
            optType(i).Value = False
        Next i
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdLoadIcons_Click()

    If sLibraryFile = "" Then
        MsgBox "Choose a library from the list!", vbInformation
        Exit Sub
    End If
    
    If Not gbAllSizeFormat And Not IsOptionChecked() Then
        MsgBox "Choose All Size and Formats, or a Size and Color depth, first", vbExclamation
        Exit Sub
    End If
    
    Msg "Loading. Please wait..."
    GetIconsFromLibrary sLibraryFile
    Msg "Loading done."
    Me.SetFocus
End Sub

Private Sub Form_Load()
    
    ' By default, search for ALL Size and Format image
    chkAllSizeFormat.Value = 1
    
    ' add some library paths...
    'List1.AddItem "XP_icons_16.dll"    ' I have leave out this libraries
    'List1.AddItem "XP_icons_24.dll"    ' because PSC don't accept DLL!
    'List1.AddItem "XP_icons_48.dll"    ' Sorry for this. If you want, you
    'List1.AddItem "XP_icons_all.dll"   ' can download them from my web site
    
    ' Add other paths or change existing one based on your system!)
    List1.AddItem "C:\Windows\Explorer.exe"
    List1.AddItem "C:\Windows\System32\shell32.dll"
    List1.AddItem "C:\Windows\System32\moricons.dll"
    
    
    Msg "Select a library from the list, then click Load icons."
    

    
End Sub

Private Sub List1_Click()
Dim sPath As String

    ' the first 4 DLL is stored on \DLL\ subfolder
'    If List1.ListIndex < 4 Then
'        sPath = App.Path & "\DLL\" & List1.Text
'    Else
        sPath = List1.Text
'    End If
    
    ' Sure that library exists ?
    If Dir(sPath) = "" Then
        MsgBox "This file don't exists!", vbCritical, "EnumResource"
        On Error Resume Next
        List1.Selected(List1.Text) = False
        Exit Sub
    End If

    ' is a library complete path name
    sLibraryFile = sPath
    Msg sPath
    
    cmdLoadIcons.Enabled = True
    
End Sub


Private Sub mnuFile_Click(Index As Integer)
Const FILE_OPEN = 0
Const FILE_SHOW = 1
Const FILE_EXIT = 3
    Select Case Index
        Case FILE_OPEN
            Shell "Notepad.exe " & App.Path & "\README.TXT", vbNormalFocus
        Case FILE_SHOW
            frmSample.Show , Me
        Case FILE_EXIT
            Unload Me
    End Select
End Sub

Private Sub mnuHelp_Click()
    Dim s As String
        s = "Enumerate Icon Resoruces" & vbCrLf
        s = s & "by Giorgio Brausi (aka GIBRA)" & vbCrLf & vbCrLf
        s = s & "website: http://www.vbcorner.net" & vbCrLf
        s = s & "e-mail:  vbcorner@vbcorner.net" & vbCrLf & vbCrLf
        s = s & "Contact me for any bug or suggestion."
        s = s & "Visit my web site for free menu tools!"
        MsgBox s, vbInformation
        
End Sub

Private Sub optSize_Click(Index As Integer)
    
    Select Case Index
        Case 0
            giSize = 16
        Case 1
            giSize = 24
        Case 2
            giSize = 32
        Case 3
            giSize = 48
    End Select
    
    chkAllSizeFormat.Value = 0
    
End Sub


Private Sub optType_Click(Index As Integer)
    Select Case Index
        Case 0
            giColorDepth = 4
        Case 1
            giColorDepth = 16
        Case 2
            giColorDepth = 24
        Case 3
            giColorDepth = 32
    End Select
    
    chkAllSizeFormat.Value = 0
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure   : GetIconsFromLibrary
' DateTime    : 04/04/2004 17.40
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Get all icons from library, fill the ImageCombo1 with images
'               and other info (resource number, bytes and size
'---------------------------------------------------------------------------------------

Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
    
    ' clear objects
    ImageCombo1.ImageList = Nothing
    ImageCombo1.ComboItems.Clear
    Toolbar1.ImageList = Nothing
    ImageList1.ListImages.Clear
    List2.Clear
    StatusBar1.Panels(2).Text = ""
    
    ' libraries with many icons require a lot of time
    Screen.MousePointer = vbHourglass
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
    
    Screen.MousePointer = vbNormal
    
    If ImageList1.ListImages.Count = 0 Then
        StatusBar1.Panels(2).Text = "No images"
        Exit Sub
    End If
    
    ' add images to ImageCombo1...
    ImageCombo1.ImageList = ImageList1
    For i = 1 To ImageList1.ListImages.Count
        ImageCombo1.ComboItems.Add , , "icon " & ImageList1.ListImages(i).Key, ImageList1.ListImages(i).Index
    Next i
      
    ' ... and to Toolbar1
    Toolbar1.ImageList = ImageList1
    ' find the max numbers of images to load on toolbar
    iCount = IIf(Toolbar1.Buttons.Count > ImageList1.ListImages.Count, ImageList1.ListImages.Count, Toolbar1.Buttons.Count)

    For i = 1 To iCount
        Toolbar1.Buttons(i).Image = ImageList1.ListImages(i).Index
    Next i
    
    ' show how many images is loaded from library file
    StatusBar1.Panels(2).Text = ImageList1.ListImages.Count & " images"
    
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : Msg
' DateTime    : 04/04/2004 17.44
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Show a message to statusbar
'---------------------------------------------------------------------------------------

Public Sub Msg(ByVal s As String)
    StatusBar1.Panels(1).Text = s
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : IsOptionChecked
' DateTime    : 04/04/2004 11.09
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Check if at least one Size and Color depth option is checked
' Descritpion : If we start the searching for a specific size and format
'               we need 'both' the values, otherwise error occur.
' Comments    :
'---------------------------------------------------------------------------------------

Public Function IsOptionChecked() As Boolean
Dim i As Integer, bSize As Boolean, bType As Boolean

    For i = 0 To 3
        If optSize(i) Then
            bSize = True
            Exit For
        End If
    Next i
    
    For i = 0 To 3
        If optType(i) Then
            bType = True
            Exit For
        End If
    Next i
    IsOptionChecked = bSize And bType
    
End Function
