<div align="center">

## FindFile \- Fast, using the API


</div>

### Description

Uses the FindFile, FindNextFile, and SearchPath API functions to quickly find a file on your hard drive. Runs faster than methods which use Dir$.
 
### More Info
 
Filename - Filename to search for.

Path - The path to start searching from.

None, if you want to find out what the API does exactly, read the Win32SDK. It's great for stuff like that!

Returns the full path to the filename, if it found it. Otherwise, returns an empty string ("").

Since the function's recursive, i guess you could hit a stack overflow if you had an obscene number of folders to search through :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Hng](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-hng.md)
**Level**          |Unknown
**User Rating**    |4.3 (91 globes from 21 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-hng-findfile-fast-using-the-api__1-1446/archive/master.zip)

### API Declarations

```
'Lots here.. :)
Public Const MAX_PATH As Long = 260
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Type FileTime
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FileTime
  ftLastAccessTime As FileTime
  ftLastWriteTime As FileTime
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
```


### Source Code

```
Public Function FindFile(ByVal FileName As String, ByVal Path As String) As String
Dim hFile As Long, ts As String, WFD As WIN32_FIND_DATA
Dim result As Long, sAttempt As String, szPath As String
szPath = GetRDP(Path) & "*.*" & Chr$(0)
'Note: Inline function here
'----Starts----
Dim szPath2 As String, szFilename As String, dwBufferLen As Long, szBuffer As String, lpFilePart As String
'Set variables
szPath2 = Path & Chr$(0)
szFilename = FileName & Chr$(0)
szBuffer = String$(MAX_PATH, 0)
dwBufferLen = Len(szBuffer)
'Ask windows if it can find a file matching the filename you gave it.
result = SearchPath(szPath2, szFilename, vbNullString, dwBufferLen, szBuffer, lpFilePart)
'----Ends----
If result Then
  FindFile = StripNull(szBuffer)
  Exit Function
End If
'Start asking windows for files.
hFile = FindFirstFile(szPath, WFD)
Do
  If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
    'Hey look, we've got a directory!
    ts = StripNull(WFD.cFileName)
    If Not (ts = "." Or ts = "..") Then
      'Don't look for hidden or system directories
      If Not (WFD.dwFileAttributes And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) Then
        'Search directory recursively
        sAttempt = FindFile(FileName, GetRDP(Path) & ts)
        If sAttempt <> "" Then
          FindFile = sAttempt
          Exit Do
        End If
      End If
    End If
  End If
  WFD.cFileName = ""
  result = FindNextFile(hFile, WFD)
Loop Until result = 0
FindClose hFile
End Function
Public Function StripNull(ByVal WhatStr As String) As String
  Dim pos As Integer
  pos = InStr(WhatStr, Chr$(0))
  If pos > 0 Then
    StripNull = Left$(WhatStr, pos - 1)
  Else
    StripNull = WhatStr
  End If
End Function
Public Function GetRDP(ByVal sPath As String) As String
'Adds a backslash on the end of a path, if required.
  If sPath = "" Then Exit Function
  If Right$(sPath, 1) = "\" Then GetRDP = sPath: Exit Function
  GetRDP = sPath & "\"
End Function
```

