====
VBA-CRC32
====

Calc CRC32 for string or file.

Installation
============

Import file-module in VBA project

Function
============

* CalcStr(Text AS String) AS String
* CalcFile(PathFile AS String, [FileBuffer AS Long = 32768]) AS String

Example
============

Example of calc string:

    Dim Str, StrCRC32 as String
    Str = "Test string"
    StrCRC32 = CRC32.CalcStr(Str)

Example of calc file:

    Dim FileName, StrCRC32 as String
    Dim FileBuffer as Long
    FileName = "C:\Test_file_for_CRC32.txt"
    FileBuffer = 65536
    StrCRC32 = CRC32.CalcFile(FileName, FileBuffer)
    If Not StrCRC32 Then
        Debug.Print "File not found!"
    End If
