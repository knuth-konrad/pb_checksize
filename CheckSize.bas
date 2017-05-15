'------------------------------------------------------------------------------
'Purpose  : Compares a file's size to a given value
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 2002
'   Source: -
'  Changed: 14.02.2017
'           - Refactoring for Github
'           04.05.2017
'           - #BREAK On to prevent console context menu change
'           15.05.2017
'           - application manifest added.
'------------------------------------------------------------------------------
#Compile Exe ".\CheckSize.exe"
#Option Version5
#Break On
#Dim All

#Debug Error Off
#Tools Off

%VERSION_MAJOR = 1
%VERSION_MINOR = 0
%VERSION_REVISION = 3

' Version Resource information
#Include ".\CheckSizeRes.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include "win32api.inc"
#Include "sautilcc.inc"

Declare Sub ShowHelp
Declare Function GetFileLength(sFileName As String) As Quad
Declare Function CalcVal (ByVal sValue As String) As Quad
'------------------------------------------------------------------------------
'*** Variabels ***
'------------------------------------------------------------------------------
'==============================================================================

Function PBMain()
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local i As Long            'Loop counter
   Local sCommand As String   'Command line
   Local sFile As String      'File to check
   Local sTemp As String
   Local qudVal As Quad, qudSize As Quad
   Local lCompare As Long

   ' Application intro
   ConHeadline "CheckSize", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2002-2017", $COMPANY_NAME
   Print ""

   Trace New ".\CheckSize.tra"

   '** Parse command line
   sCommand = Command$
   If Len(Trim$(sCommand)) < 1 Then
      ShowHelp
      Function = 100
      Exit Function
   End If

   For i = 1 To ArgC()
      sTemp = ArgV(i)
      If LCase$(Left$(sTemp, 2)) = "/f" Then
         sFile = Trim$(Mid$(sTemp, 4))
         If IsFalse(FileExist(ByCopy sFile)) Then
         'File does not exist
            Function = 50
            Exit Function
         End If
      Else
         Select Case LCase$(Left$(Trim$(sTemp), 2))
         Case "/s"
            qudVal = CalcVal(Mid$(sTemp, 4))
         Case "/c"
            lCompare = Val(Mid$(sTemp, 4))
         Case Else
            ShowHelp
            Function = 100
            Exit Function
         End Select
      End If
   Next i

   If lCompare >= -2 And lCompare <= 2 Then
      qudSize = GetFileLength(sFile)
      StdOut "Comparing: " & sFile & " ";
      Select Case lCompare
      Case 0
         StdOut "=";
         If qudSize = qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case 2
         StdOut ">";
         If qudSize > qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case 1
         StdOut ">=";
         If qudSize >= qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case -2
         StdOut "<";
         If qudSize < qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case -1
         StdOut "<=";
         If qudSize <= qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case Else
         Function = 100
      End Select
      StdOut " " & Extract$(FormatNumber(qudVal), ",") & " bytes."
   Else
      Function = 100
   End If

   StdOut sTemp

End Function
'---------------------------------------------------------------------------

Sub ShowHelp
'------------------------------------------------------------------------------
'Purpose  : Help screen
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   StdOut "CheckSize"
   StdOut "--------"
   StdOut "CheckSize determines if a file's size fits into the given comparison."
   StdOut ""
   StdOut "Usage:   CheckSize /f=<file> /s=<size> /c=<compare argument>"
   StdOut "i.e.     CheckSize /f=f:\data\myfile.exe /s=1000000 /c=1"
   StdOut ""
   StdOut "Parameters"
   StdOut "----------"
   StdOut "/f   = File to check."
   StdOut "/s   = Size to check against. Valid fomats:"
   StdOut "       1000 - figures only = Bytes."
   StdOut "        2kb - Kilobytes where 1kb = 1024 Bytes."
   StdOut "        5mb - Megabytes where 1mb = 1024 Kilobytes."
   StdOut "        3gb - Gigabytes where 1gb = 1024 Megabytes."
   StdOut "        1tb - Terrabytes where 1tb = 1024 Gigabytes."
   StdOut "/c   = Comparison to perform. Valid operators:"
   StdOut "     - -2 file size must be lesser than given size."
   StdOut "     - -1 file size must be lesser than or equal to given size."
   StdOut "     -  0 file size must equal to given size."
   StdOut "     -  1 file size must be greater than or equal to given size."
   StdOut "     -  2 file size must be greater than given size."
   StdOut ""
   StdOut "CheckSize returns the following DOS error levels upon exit, determing success"
   StdOut "or failure of the operation:"
   StdOut "  0  = File size has passed comparison."
   StdOut "  1  = File size has *not* passed comparison."
   StdOut " 50  = File not found/doesn't exist"
   StdOut "100  = Invalid/missing command line parameter"
   StdOut "255  = Other (applciation) error"

End Sub
'---------------------------------------------------------------------------

Function GetFileLength(ByRef sFileName As String) As Quad
'------------------------------------------------------------------------------
'Purpose  : Return a file's size
'
'Prereq.  : -
'Parameter: sFileName   - File name incl. full path
'Returns  : File size in bytes
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local W32FD             As WIN32_FIND_DATA
   Local hFile             As Dword

   ' Safeguard
   If Len(sFileName) = 0 Then
      Exit Function
   End If

   hFile = FindFirstFile(ByVal StrPtr(sFileName), W32FD)
   If hFile <> %INVALID_HANDLE_VALUE Then
      Function = W32FD.nFileSizeHigh * &H0100000000 + W32FD.nFileSizeLow
      FindClose hFile
   End If

End Function
'---------------------------------------------------------------------------

Function CalcVal (ByVal sValue As String) As Quad
'------------------------------------------------------------------------------
'Purpose  : Calculate unit multiplier for file size
'
'Prereq.  : -
'Parameter: sValue   - Size unit
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   sValue = LCase$(sValue)
   Select Case Right$(sValue, 2)
   Case "kb"
      CalcVal = Val(sValue) * 1024&&
   Case "mb"
      CalcVal = Val(sValue) * 1024&&^2
   Case "gb"
      CalcVal = Val(sValue) * 1024&&^3
   Case "tb"
      CalcVal = Val(sValue) * 1024&&^4
   Case Else
      CalcVal = Val(sValue)
   End Select

End Function
'---------------------------------------------------------------------------
