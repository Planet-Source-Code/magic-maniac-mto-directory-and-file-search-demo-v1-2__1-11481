VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory and File Search Demo!"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkSubDirectorys 
      Caption         =   "&Include Sub-Directories"
      Height          =   330
      Left            =   4770
      TabIndex        =   10
      Top             =   4140
      Width           =   1995
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2445
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   4313
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DateTime"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Attr"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   5490
      TabIndex        =   4
      Top             =   5220
      Width           =   1275
   End
   Begin VB.TextBox TxtFilters 
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   4815
      Visible         =   0   'False
      Width           =   6720
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   45
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Text            =   "Directories in Directories"
      Top             =   3105
      Width           =   6720
   End
   Begin VB.TextBox TxtPaths 
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   3735
      Width           =   6720
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   330
      Left            =   4185
      TabIndex        =   3
      Top             =   5220
      Width           =   1275
   End
   Begin VB.Label LblStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Total: 0 Results"
      Height          =   240
      Left            =   45
      TabIndex        =   8
      Top             =   2520
      Width           =   6720
   End
   Begin VB.Label LblType 
      Caption         =   "Search Type"
      Height          =   240
      Left            =   45
      TabIndex        =   7
      Top             =   2835
      Width           =   6720
   End
   Begin VB.Label LblFilters 
      Caption         =   "File Filter"
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   4545
      Visible         =   0   'False
      Width           =   6720
   End
   Begin VB.Label LblPaths 
      Caption         =   "Search in Path"
      Height          =   240
      Left            =   45
      TabIndex        =   5
      Top             =   3465
      Width           =   6720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# -------------------------------------
'# DIRECTORY AND FILE SEARCH DEMO
'#
'# Version 1.2 (BUGFIX VERSION)
'# -------------------------------------
'# CODED BY:
'#
'# MAGiC MANiAC^mTo ( mto@kabelfoon.nl )
'#
'# MORTAL OBSESSiON
'# http://home.kabelfoon.nl/~mto
'# -------------------------------------
'# RELEASED 17-NOV-2000 ON:
'#
'# www.planet-source-code.com
'# -------------------------------------

Private Sub Form_Load()
  On Error Resume Next
  CmbType.ListIndex = GetSetting(App.CompanyName, App.Title, "Type", 0)
  If Err Then
    CmbType.ListIndex = 0
  End If
  TxtPaths.Text = GetSetting(App.CompanyName, App.Title, "Paths", "c:\windows;c:\program files")
  ChkSubDirectorys.Value = GetSetting(App.CompanyName, App.Title, "SubDirs", 0)
  If Err Then
    ChkSubDirectorys.Value = 0
  End If
  TxtFilters.Text = GetSetting(App.CompanyName, App.Title, "Filters", "*.bat;*.com;*.exe")
  On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MsgBox "Don't forget to vote me on www.planet-source-code.com :-)"
  End
End Sub

Private Sub CmbType_Click()
  SaveSetting App.CompanyName, App.Title, "Type", CmbType.ListIndex
  TxtFilters.Visible = CmbType.ListIndex = 1
  LblFilters.Visible = CmbType.ListIndex = 1
End Sub

Private Sub TxtPaths_Change()
  SaveSetting App.CompanyName, App.Title, "Paths", TxtPaths.Text
End Sub

Private Sub ChkSubDirectorys_Click()
  SaveSetting App.CompanyName, App.Title, "SubDirs", ChkSubDirectorys.Value
End Sub

Private Sub TxtFilters_Change()
  SaveSetting App.CompanyName, App.Title, "Filters", TxtFilters.Text
End Sub

Private Sub CmdSearch_Click()
  Dim lTmp1 As Long
  Dim sStr1 As String
  Dim lItem As ListItem
  Dim cCol As tSearch
  CmdSearch.Enabled = False
  Me.MousePointer = vbHourglass
  ListView1.ListItems.Clear
  LblStatus.Alignment = vbLeftJustify
  If CmbType.ListIndex = 0 Then
    If ChkSubDirectorys.Value Then
      LblStatus.Caption = "Please Wait, Searching Sub-Directories..."
      GetSubDirs TxtPaths.Text, vbDirectory, cCol
    Else
      LblStatus.Caption = "Please Wait, Searching Directories..."
      GetDirs TxtPaths.Text, vbDirectory, cCol
    End If
  Else
    If ChkSubDirectorys.Value Then
      LblStatus.Caption = "Please Wait, Searching Sub-Files..."
      GetSubFiles TxtPaths.Text, TxtFilters.Text, vbDirectory, vbArchive, cCol
    Else
      LblStatus.Caption = "Please Wait, Searching Files..."
      GetFiles TxtPaths.Text, TxtFilters.Text, vbArchive, cCol
    End If
  End If
  For lTmp1 = 1 To cCol.Count
    Set lItem = ListView1.ListItems.Add(, , cCol.Path(lTmp1))
    lItem.SubItems(1) = Format(cCol.Size(lTmp1), "###,###,##0")
    lItem.SubItems(2) = Format(cCol.DateTime(lTmp1), "DD-MM-YY HH:MM:SS")
    lItem.SubItems(3) = sAttr(cCol.Attr(lTmp1))
  Next
  LblStatus.Alignment = vbRightJustify
  LblStatus.Caption = "Total: " & ListView1.ListItems.Count & " Results"
  Me.MousePointer = vbDefault
  CmdSearch.Enabled = True
End Sub

Private Sub CmdExit_Click()
  Unload Me
End Sub

Function sAttr(Attr As VbFileAttribute) As String
  Dim sStr1 As String
  sStr1 = ""
  If Attr And vbReadOnly Then sStr1 = "r" Else sStr1 = "-"
  If Attr And vbArchive Then sStr1 = sStr1 + "a" Else sStr1 = sStr1 + "-"
  If Attr And vbHidden Then sStr1 = sStr1 + "h" Else sStr1 = sStr1 + "-"
  If Attr And vbSystem Then sStr1 = sStr1 + "s" Else sStr1 = sStr1 + "-"
  sAttr = sStr1
End Function

