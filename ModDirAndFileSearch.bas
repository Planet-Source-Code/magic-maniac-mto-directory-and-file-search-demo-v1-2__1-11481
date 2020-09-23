Attribute VB_Name = "ModDirAndFileSearch"
'# ---------------------------------------------------
'# DIRECTORY AND FILE SEARCH UTILITY
'#
'# Version 1.2 (BUGFIX VERSION)
'# ---------------------------------------------------
'# DESCRIPTION:
'#
'# search directories, sub-directories and files without
'# having any object on your form.
'# ---------------------------------------------------
'# WHATS NEW:
'#
'# add: size, datetime, attribute.
'# add: sub getsubfiles.
'# fix: duplicated search results will not add again.
'# fix: add a correct char on the right of the string.
'# ---------------------------------------------------
'# CODED BY:
'#
'# MAGiC MANiAC^mTo, ( mto@kabelfoon.nl )
'#
'# MORTAL OBSESSiON:
'# http://home.kabelfoon.nl/~mto
'# ---------------------------------------------------
'# RELEASED 17-NOV-2000 ON:
'#
'# www.planet-source-code.com
'# ---------------------------------------------------

Option Explicit

Public Type tSearch
  Count As Long
  Path As New Collection
  Size As New Collection
  DateTime As New Collection
  Attr As New Collection
End Type

'# Get Directories In Directories...
'#
'# sDir = "c:\windows" or sDir = "c:\windows;c:\windows\system"
'# DirAttr = vbDirectory or DirAttr = vbDirectory + vbHidden
'# cCol = Your tSearch
Public Sub GetDirs(ByVal sDir As String, DirAttr As VbFileAttribute, cCol As tSearch)
  Dim lTmp1 As Long
  Dim sStr1 As String
  Dim sStr2 As String
  Dim sResult() As String
  sStr2 = ""
  For lTmp1 = 0 To sSplit(sDir, "", sResult)
    sResult(lTmp1) = Trim$(sResult(lTmp1))
    If Right$(sResult(lTmp1), 1) <> "\" Then
      sResult(lTmp1) = sResult(lTmp1) + "\"
    End If
    If InStr(sStr2, UCase$(sResult(lTmp1)) + ";") < 1 Then
      sStr2 = sStr2 + UCase$(sResult(lTmp1)) + ";"
      sStr1 = Dir$(sResult(lTmp1) + "*.*", DirAttr)
      While sStr1 <> ""
        DoEvents
        If sStr1 <> "." And sStr1 <> ".." Then
          If (GetAttr(sResult(lTmp1) + sStr1) And vbDirectory) = vbDirectory Then
            cCol.Path.Add sResult(lTmp1) + sStr1
            cCol.Size.Add 0
            cCol.DateTime.Add FileDateTime(sResult(lTmp1) + sStr1)
            cCol.Attr.Add GetAttr(sResult(lTmp1) + sStr1)
          End If
        End If
        sStr1 = Dir
      Wend
    End If
  Next
  cCol.Count = cCol.Path.Count
End Sub

'# Get Sub-Directories In Directories...
'#
'# sDir = "c:\windows" or sDir = "c:\windows;c:\windows\system;ect..."
'# DirAttr = vbDirectory or DirAttr = vbDirectory + vbHidden
'# cCol = Your tSearch
Public Sub GetSubDirs(ByVal sDir As String, DirAttr As VbFileAttribute, cCol As tSearch)
  Dim lTmp1 As Long
  Dim cCol1 As tSearch
  GetDirs sDir, DirAttr, cCol1
  For lTmp1 = 1 To cCol1.Count
    cCol.Path.Add cCol1.Path(lTmp1)
    cCol.Size.Add 0
    cCol.DateTime.Add cCol1.DateTime(lTmp1)
    cCol.Attr.Add cCol1.Attr(lTmp1)
    GetSubDirs cCol1.Path(lTmp1), DirAttr, cCol
  Next
  cCol.Count = cCol.Path.Count
End Sub

'# Get Files In Directories...
'#
'# sDir = "c:\windows" or sDir = "c:\window;c:\windows\system;ect..."
'# sFilter = "*.*" or sFilter = "*.bat;*.com;*.exe;ect..."
'# FileAttr = vbArchive or FileAttr = vbArchive + vbHidden
'# cCol = Your tSearch
Public Sub GetFiles(sDir As String, sFilter As String, FileAttr As VbFileAttribute, cCol As tSearch)
  Dim lTmp1 As Long
  Dim lTmp2 As Long
  Dim lTmp3 As Long
  Dim sStr1 As String
  Dim sStr2 As String
  Dim sStr3 As String
  Dim sResult1() As String
  Dim sResult2() As String
  sStr2 = ""
  For lTmp1 = 0 To sSplit(sDir, "", sResult1)
    sResult1(lTmp1) = Trim$(sResult1(lTmp1))
    If Right$(sResult1(lTmp1), 1) <> "\" Then
      sResult1(lTmp1) = sResult1(lTmp1) + "\"
    End If
    If InStr(sStr2, UCase$(sResult1(lTmp1)) + ";") < 1 Then
      sStr2 = sStr2 + UCase$(sResult1(lTmp1)) + ";"
      sStr3 = ""
      For lTmp2 = 0 To sSplit(sFilter, "", sResult2)
        sResult2(lTmp2) = Trim$(sResult2(lTmp2))
        If InStr(sStr3, UCase$(sResult2(lTmp2)) + ";") < 1 Then
          sStr3 = sStr3 + UCase$(sResult2(lTmp2)) + ";"
          sStr1 = Dir$(sResult1(lTmp1) + sResult2(lTmp2), FileAttr)
          DoEvents
          While sStr1 <> ""
            cCol.Path.Add sResult1(lTmp1) + sStr1
            cCol.Size.Add FileLen(sResult1(lTmp1) + sStr1)
            cCol.DateTime.Add FileDateTime(sResult1(lTmp1) + sStr1)
            cCol.Attr.Add GetAttr(sResult1(lTmp1) + sStr1)
            sStr1 = Dir
          Wend
        End If
      Next
    End If
  Next
  cCol.Count = cCol.Path.Count
End Sub

'# Get Sub-Files In Directories...
'#
'# sDir = "c:\windows" or sDir = "c:\window;c:\windows\system;ect..."
'# sFilter = "*.*" or sFilter = "*.bat;*.com;*.exe;ect..."
'# DirAttr = vbDirectory or DirAttr = vbDirectory + vbHidden
'# FileAttr = vbArchive or FileAttr = vbArchive + vbHidden
'# cCol = Your tSearch
Public Sub GetSubFiles(sDir As String, sFilter As String, DirAttr As VbFileAttribute, FileAttr As VbFileAttribute, cCol As tSearch)
  Dim lTmp1 As Long
  Dim sStr1 As String
  Dim cCol1 As tSearch
  GetSubDirs sDir, DirAttr, cCol1
  sStr1 = ""
  For lTmp1 = 1 To cCol1.Count
    sStr1 = sStr1 + cCol1.Path(lTmp1) + ";"
  Next
  GetFiles sStr1, sFilter, FileAttr, cCol
  cCol.Count = cCol.Path.Count
End Sub

'# Split A String...
'#
'# sSplit = Total Strings...
'# sStr1 = "c:\windows" or sStr1 = "c:\windows;c:\windows\system;ect..."
'# sDelims = ";" or sDelims = ";" + chr$(0) + ect...
'# sResult = Dim sResult() As String
Private Function sSplit(ByVal sStr1 As String, sDelims As String, sResult() As String) As Long
  Dim nResult As Long
  Dim lTmp1 As Long
  Dim lTmp2 As Long
  If sDelims = "" Then
    sDelims = ";" + Chr$(0) + Chr$(9) + Chr$(10) + Chr$(13)
  End If
  If InStr(1, Right$(sStr1, 1), sDelims, vbBinaryCompare) < 1 Then
    sStr1 = sStr1 + Left$(sDelims, 1)
  End If
  nResult = -1
  lTmp1 = 1
  For lTmp2 = 1 To Len(sStr1)
    If InStr(1, sDelims, Mid$(sStr1, lTmp2, 1), vbBinaryCompare) > 0 Then
      nResult = nResult + 1
      ReDim Preserve sResult(0 To nResult) As String
      sResult(nResult) = Mid$(sStr1, lTmp1, lTmp2 - lTmp1)
      lTmp1 = lTmp2 + 1
    End If
  Next
  If lTmp1 < 3 Then
    nResult = -1
  End If
  sSplit = nResult
End Function
