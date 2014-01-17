====
VBA-CRC32
====

Calc CRC32 for string or file

Installation
============

Import file-module in VBA project

Example
============

Example of calc string::

    Dim Str, StrCRC32 as String
    Str = "Test string"
    StrCRC32 = CRC32.CalcStr(Str)

Example of calc file::

    Dim FileName, StrCRC32 as String
    FileName = "C:\Test_file_for_CRC32.txt"
    StrCRC32 = CRC32.CalcFile(FileName)
