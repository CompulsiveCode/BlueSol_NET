'modFileAPI - Written by Jesse Yeager.   www.CompulsiveCode.com
'
'This module wraps the Win32api File IO routines.

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

Module modBlueSoleil_FileAPI



    'parameters for the CreateFile function.
    Private Const GENERIC_READ As Integer = &H80000000
    Private Const GENERIC_WRITE As Integer = &H40000000
    Private Const CREATE_ALWAYS As Integer = 2
    Private Const OPEN_EXISTING As Integer = 3
    Private Const FILE_SHARE_READ As Integer = &H1S
    Private Const FILE_SHARE_WRITE As Integer = &H2S
    Private Const FILE_FLAG_OVERLAPPED As Integer = &H40000000
    Private Const FILE_FLAG_NO_BUFFERING As Integer = &H20000000
    Private Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
    Private Const FILE_ATTRIBUTE_DIRECTORY As Integer = &H10

    Private Const SECURITY_ATTRIBUTES_UDTsize As Integer = 12

    Private Const INVALID_HANDLE_VALUE As Integer = -1

    Private Declare Sub CopyMemory_FromBYTEtoINT32 Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Integer, ByRef lpvSource As Byte, ByVal cbCopy As UInt32)

    Private Declare Function CreateFile_NoUDT Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByRef lpSecurityAttributes As Byte, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As IntPtr
    Private Declare Function WriteFile_NoUDT Lib "kernel32" Alias "WriteFile" (ByVal hFile As IntPtr, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Integer, ByRef lpNumberOfBytesWritten As Integer, ByVal lpOverlapped As Integer) As Integer
    Private Declare Function ReadFile_NoUDT Lib "kernel32" Alias "ReadFile" (ByVal hFile As IntPtr, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Integer, ByRef lpNumberOfBytesRead As Integer, ByVal lpOverlapped As Integer) As Integer

    Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As IntPtr) As Integer
    Private Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As IntPtr, ByVal lDistanceToMove As Long, ByRef lpNewPointer As UInt64, ByVal dwMoveMethod As Integer) As Integer


    'parameters for the SetFilePointer function.
    Private Const FILE_BEGIN As Integer = 0
    Private Const FILE_CURRENT As Integer = 1
    Private Const FILE_END As Integer = 2

    Private Declare Function FindFirstFile_NoUDT Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As Byte) As IntPtr
    Private Declare Function FindNextFile_NoUDT Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As IntPtr, lpFindFileData As Byte) As Integer
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As IntPtr) As Integer

    Private Const WIN32_FIND_DATA_UDTsize As Integer = 318
    Private Const W32FD_MAX_PATH As Integer = 260





    <DllImport("kernel32.dll", CharSet:=CharSet.Auto)>
    Private Function GetFileSizeEx(<[In]()> ByVal hFile As IntPtr, <[In](), Out()> ByRef lpFileSize As Long) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function CloseHandle(ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function



    Private Sub UDT_InitSECURITY_ATTRIBUTES(ByRef udtBytes() As Byte)

        ReDim udtBytes(SECURITY_ATTRIBUTES_UDTsize - 1)

        udtBytes(0) = SECURITY_ATTRIBUTES_UDTsize 'set structure size.

    End Sub


    Public Function FileAPI_GetDirectorySize(ByVal dirName As String, ByVal includeSubfolders As Boolean) As Long

        Dim fArray(0 To 0) As String

        If includeSubfolders = True Then
            fArray = IO.Directory.GetFiles(dirName, "*.*", IO.SearchOption.AllDirectories)
        Else
            fArray = IO.Directory.GetFiles(dirName, "*.*", IO.SearchOption.TopDirectoryOnly)
        End If

        Dim totalSize As Long = 0

        Dim i As Integer
        For i = 0 To fArray.Length - 1
            Dim fInfo As IO.FileInfo = New IO.FileInfo(fArray(i))
            totalSize = totalSize + fInfo.Length

        Next i

        Return totalSize

    End Function



    Public Function FileAPI_GetFileSize(ByVal hFile As IntPtr) As Long


        Dim retSize As Long = -1
        Dim retBool As Boolean = GetFileSizeEx(hFile, retSize)
        Return retSize


    End Function

    Public Function FileAPI_CloseFile(ByVal hFile As IntPtr) As Boolean

        Dim retVal As Boolean = False
        Try

            retVal = CloseHandle(hFile)
        Catch ex As Exception

        End Try


        Return retVal

    End Function

    Public Function FileAPI_GetBytes(ByVal hFile As IntPtr, ByVal fileOffset As Long, ByVal readNumBytes As Integer, ByRef byteArray() As Byte) As Boolean


        If fileOffset > -1 Then
            FileAPI_SetFileOffset(hFile, fileOffset)
        End If

        If readNumBytes < 1 Then
            readNumBytes = -1
            readNumBytes = UBound(byteArray) - LBound(byteArray) + 1
        End If

        'unnecessary shit i added while trying to test/fix something.
        'Dim currFileSize As Long = FileAPI_GetFileSize(hFile)
        'Dim currOffset As Long = FileAPI_GetCurrentOffset(hFile)
        'If currOffset + readNumBytes > currFileSize Then
        '    readNumBytes = CInt(currFileSize - currOffset)
        'End If


        Dim numBytesRead As Integer = 0, retVal As Integer = 0
        If readNumBytes > 0 Then
            ReDim byteArray(readNumBytes - 1)
            retVal = ReadFile_NoUDT(hFile, byteArray(0), readNumBytes, numBytesRead, 0)

            If numBytesRead > 0 Then
                ReDim Preserve byteArray(0 To numBytesRead - 1)
            End If
        End If

        FileAPI_GetBytes = (retVal <> 0)

    End Function

    Public Function FileAPI_OpenFile(ByVal theFileName As String, ByVal createNewFile As Boolean) As IntPtr

        Dim secBytes(0 To 0) As Byte
        UDT_InitSECURITY_ATTRIBUTES(secBytes)

        Dim creationFlags As Integer
        If createNewFile = False Then
            creationFlags = OPEN_EXISTING
        Else
            creationFlags = CREATE_ALWAYS
        End If

        theFileName = FileAPI_GetUNCfileName(theFileName)

        Dim hFile As IntPtr

        hFile = CreateFile_NoUDT(theFileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, secBytes(0), creationFlags, 0, 0)

        FileAPI_OpenFile = hFile

    End Function

    Public Function FileAPI_PutBytes(ByVal hFile As IntPtr, ByVal fileOffset As Long, ByVal writeNumBytes As Integer, ByRef byteArray() As Byte) As Boolean


        If fileOffset > -1 Then
            FileAPI_SetFileOffset(hFile, fileOffset)
        End If


        If writeNumBytes < 1 Then
            writeNumBytes = -1
            writeNumBytes = UBound(byteArray) - LBound(byteArray) + 1
        End If

        Dim numBytesWritten, retVal As Integer

        If writeNumBytes > 0 Then
            retVal = WriteFile_NoUDT(hFile, byteArray(0), writeNumBytes, numBytesWritten, 0)
        End If

        FileAPI_PutBytes = (retVal <> 0)

    End Function


    Public Function FileAPI_SetFileOffset(ByVal hFile As IntPtr, ByVal fileOffset As Long) As Boolean

        Dim retInt As Integer
        retInt = SetFilePointerEx(hFile, fileOffset, 0, FILE_BEGIN)
        Return (retInt <> 0)


    End Function

    Public Function FileAPI_IsEOF(ByVal hFile As IntPtr) As Boolean

        Dim fLen As Long, fPos As Long
        fLen = FileAPI_GetFileSize(hFile)
        fPos = FileAPI_GetCurrentOffset(hFile)

        FileAPI_IsEOF = (fPos >= fLen)

    End Function

    Public Function FileAPI_GetCurrentOffset(ByVal hFile As IntPtr) As Long

        Dim retOffset As Long
        Dim tempU64 As UInt64
        SetFilePointerEx(hFile, 0, tempU64, FILE_CURRENT)
        retOffset = CLng(tempU64)
        Return retOffset

    End Function



    Public Function FileAPI_ReadLineFromBinaryFile(ByVal hFile As IntPtr, ByVal inpFoffset As Long, ByVal inpFLen As Long, ByVal EOLstring As String, ByRef retLineStr As String) As Boolean

        retLineStr = ""
        If inpFoffset < 0 Then inpFoffset = FileAPI_GetCurrentOffset(hFile)

        If inpFLen < 1 Then inpFLen = FileAPI_GetFileSize(hFile)

        Dim EOLstringLen As Integer
        EOLstringLen = Len(EOLstring)

        Dim retVal As Boolean
        retVal = True

        Dim tempStr As String
        Dim EOLpos As Integer

        Dim tempOffset As Long
        tempOffset = inpFoffset

        Do
            If tempOffset > inpFLen Then Exit Do

            tempStr = Space(4000)
            FileAPI_GetString(hFile, tempOffset, -1, tempStr)

            EOLpos = InStr(1, tempStr, EOLstring)

            If EOLpos > 1 Then
                retLineStr = retLineStr & Left(tempStr, EOLpos - 1)
                Exit Do
            Else
                If EOLpos = 1 Then
                    retLineStr = retLineStr
                    Exit Do
                Else
                    retLineStr = retLineStr & Left(tempStr, 4000 - EOLstringLen + 1)
                    tempOffset = tempOffset + 4000 - EOLstringLen + 1
                End If
            End If
        Loop

        FileAPI_SetFileOffset(hFile, inpFoffset + Len(retLineStr) + EOLstringLen)

        FileAPI_ReadLineFromBinaryFile = retVal

    End Function

    Public Function FileAPI_GetString(ByVal hFile As IntPtr, ByVal fileOffset As Long, ByVal readNumBytes As Integer, ByRef retString As String) As Boolean

        Dim byteArray() As Byte
        Dim numBytesRead As Integer = 0, retVal As Integer = 0

        If readNumBytes < 1 Then
            readNumBytes = Len(retString)
        End If

        retString = ""

        If fileOffset > -1 Then
            FileAPI_SetFileOffset(hFile, fileOffset)
        End If

        If readNumBytes > 0 Then
            ReDim byteArray(readNumBytes - 1)
            retVal = ReadFile_NoUDT(hFile, byteArray(0), readNumBytes, numBytesRead, 0)
            FileAPI_ByteArrayToString(byteArray, 0, numBytesRead, retString)
        End If


        FileAPI_GetString = (retVal <> 0)

    End Function


    Private Sub FileAPI_ByteArrayToString(ByRef theByteArray() As Byte, ByVal firstByteIdx As Integer, ByVal theByteCount As Integer, ByRef retString As String)


        retString = System.Text.Encoding.Default.GetString(theByteArray, firstByteIdx, theByteCount)

        'retString = Space(theByteCount)
        'Dim i As Integer
        'For i = 0 To theByteCount - 1
        ' Mid(retString, i + 1, 1) = Chr(theByteArray(i))
        'Next i


    End Sub


    Public Function FindFileAPI_FindFirst(ByVal fullPathAndPattern As String, ByRef retFileName As String, retIsDir As Boolean) As IntPtr

        Dim w32fdBytes(0 To 0) As Byte, w32fdAttribs As Integer = 0, w32fdFileName As String = ""
        UDT_InitWIN32_FIND_DATA(w32fdBytes)

        Dim retInt As IntPtr = FindFirstFile_NoUDT(fullPathAndPattern, w32fdBytes(0))
        Dim retAttribs As Integer = 0

        UDT_GetWIN32_FIND_DATA_Info(w32fdBytes, retAttribs, retFileName, 0, 0, 0, 0, 0, 0, 0, 0)

        If (retAttribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            retIsDir = True
        Else
            retIsDir = False
        End If


    End Function


    Public Function FindFileAPI_FindNext(ByVal hFind As IntPtr, ByRef retFileName As String, retIsDir As Boolean) As Boolean

        Dim w32fdBytes(0 To 0) As Byte, w32fdAttribs As Integer = 0, w32fdFileName As String = ""
        UDT_InitWIN32_FIND_DATA(w32fdBytes)

        Dim retInt As Integer = FindNextFile_NoUDT(hFind, w32fdBytes(0))
        Dim retAttribs As Integer = 0

        UDT_GetWIN32_FIND_DATA_Info(w32fdBytes, retAttribs, retFileName, 0, 0, 0, 0, 0, 0, 0, 0)

        If (retAttribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            retIsDir = True
        Else
            retIsDir = False
        End If

        Return CBool(retInt)

    End Function


    Public Function FindFileAPI_CloseHandle(ByVal hFind As IntPtr) As Boolean

        Return CBool(FindClose(hFind))

    End Function


    Public Function FileAPI_GetUNCfileName(ByVal inpFullFileName As String) As String

        inpFullFileName = Trim(inpFullFileName)

        If UCase(Left(inpFullFileName, 8)) = "\\?\UNC\" Then inpFullFileName = "\\" & Mid(inpFullFileName, 9)
        If Left(inpFullFileName, 4) = "\\?\" Then inpFullFileName = Mid(inpFullFileName, 5)

        Dim retStr As String
        If Left(inpFullFileName, 2) = "\\" Then
            retStr = "\\?\" & "UNC\" & Mid(inpFullFileName, 3)
        Else
            retStr = "\\?\" & inpFullFileName
        End If

        FileAPI_GetUNCfileName = retStr

    End Function





    Private Sub UDT_InitWIN32_FIND_DATA(ByRef udtBytes() As Byte)

        ReDim udtBytes(0 To WIN32_FIND_DATA_UDTsize - 1)

    End Sub

    Private Sub UDT_GetWIN32_FIND_DATA_Info(ByRef udtBytes() As Byte, ByRef udtFileAttributes As Integer, ByRef udtFileName As String, ByRef udtFTCreationTimeLOW As Integer, ByRef udtFTCreationTimeHIGH As Integer, ByRef udtFTAccessTimeLOW As Integer, ByRef udtFTAccessTimeHIGH As Integer, ByRef udtFTModifyTimeLOW As Integer, ByRef udtFTModifyTimeHIGH As Integer, ByRef udtFileSizeLOW As Integer, ByRef udtFileSizeHIGH As Integer)

        CopyMemory_FromBYTEtoINT32(udtFileAttributes, udtBytes(0), 4)


        CopyMemory_FromBYTEtoINT32(udtFTCreationTimeLOW, udtBytes(4), 4)
        CopyMemory_FromBYTEtoINT32(udtFTCreationTimeHIGH, udtBytes(8), 4)

        CopyMemory_FromBYTEtoINT32(udtFTAccessTimeLOW, udtBytes(12), 4)
        CopyMemory_FromBYTEtoINT32(udtFTAccessTimeHIGH, udtBytes(16), 4)

        CopyMemory_FromBYTEtoINT32(udtFTModifyTimeLOW, udtBytes(20), 4)
        CopyMemory_FromBYTEtoINT32(udtFTModifyTimeHIGH, udtBytes(24), 4)

        CopyMemory_FromBYTEtoINT32(udtFileSizeHIGH, udtBytes(28), 4)
        CopyMemory_FromBYTEtoINT32(udtFileSizeLOW, udtBytes(32), 4)

        Dim i As Integer
        Dim retStr As String
        retStr = ""
        For i = 44 To (44 + W32FD_MAX_PATH - 1)
            If udtBytes(i) = 0 Then Exit For
            retStr = retStr & Chr(udtBytes(i))
        Next i

        udtFileName = retStr

    End Sub



    Public Function FileAPI_CompareWithFile(ByVal inpFileName1 As String, ByVal inpFileName2 As String) As Boolean

        'returns TRUE if the files have the same content.

        Dim fStream1 As New IO.FileStream(inpFileName1, IO.FileMode.Open, IO.FileAccess.Read)
        Dim fStream2 As New IO.FileStream(inpFileName2, IO.FileMode.Open, IO.FileAccess.Read)
        Dim fBufferSize As Integer = 128000
        Dim fBufferData1(0 To fBufferSize - 1) As Byte
        Dim fBufferData2(0 To fBufferSize - 1) As Byte
        Dim fCurrOffset As Long = 0

        If fStream1.Length <> fStream2.Length Then
            fStream1.Close()
            fStream1.Dispose()
            fStream2.Close()
            fStream2.Dispose()
            FileAPI_CompareWithFile = False
            Exit Function
        End If

        Dim readCount As Integer = 0

        Dim i As Integer
        Dim retBool As Boolean = True

        Do
            readCount = readCount + 1
            If readCount > 1000 Then
                readCount = 0
                System.Windows.Forms.Application.DoEvents()
            End If

            If fStream1.Position >= fStream1.Length - 1 Then
                Exit Do
            End If

            If fStream1.Position + fBufferSize >= fStream1.Length Then
                fBufferSize = CInt(fStream1.Length - (fStream1.Position + 1))
                ReDim fBufferData1(0 To fBufferSize - 1)
                ReDim fBufferData2(0 To fBufferSize - 1)
            End If

            fStream1.Read(fBufferData1, 0, fBufferSize)
            fStream2.Read(fBufferData2, 0, fBufferSize)

            For i = 0 To fBufferSize - 1
                If fBufferData1(i) <> fBufferData2(i) Then
                    retBool = False
                    Exit Do
                End If
            Next

        Loop

        fStream1.Close()
        fStream1.Dispose()
        fStream2.Close()
        fStream2.Dispose()

        FileAPI_CompareWithFile = retBool

    End Function




    Public Function Browse_ShowOpenFolder(ByVal ownerForm As Form, ByVal titleText As String) As String

        Dim tempDlgResult As DialogResult
        Dim tempComDlg As New FolderBrowserDialog

        tempComDlg.ShowNewFolderButton = True
        tempComDlg.Description = titleText



        tempDlgResult = tempComDlg.ShowDialog(ownerForm)

        Dim retStr As String = ""

        If tempDlgResult <> DialogResult.Cancel Then
            retStr = tempComDlg.SelectedPath
        End If

        tempComDlg.Dispose()

        Return retStr


    End Function


    Public Function Browse_ShowOpenFile(ByVal ownerForm As Form, ByVal fileFilter As String, ByVal titleText As String) As String

        Dim tempDlgResult As DialogResult
        Dim tempComDlg As New OpenFileDialog

        tempComDlg.Multiselect = False
        tempComDlg.CheckFileExists = True
        tempComDlg.Title = titleText

        tempComDlg.Filter = fileFilter

        tempDlgResult = tempComDlg.ShowDialog(ownerForm)

        Dim retStr As String = ""

        If tempDlgResult <> DialogResult.Cancel Then
            retStr = tempComDlg.FileName
        End If

        tempComDlg.Dispose()

        Return retStr


    End Function


    Public Function Browse_ShowSaveFile(ByVal ownerForm As Form, ByVal fileFilter As String, ByVal titleText As String, ByVal defaultSaveName As String) As String

        Dim tempDlgResult As DialogResult
        Dim tempComDlg As New SaveFileDialog

        tempComDlg.OverwritePrompt = True
        tempComDlg.Title = titleText
        tempComDlg.FileName = defaultSaveName
        tempComDlg.Filter = fileFilter


        tempDlgResult = tempComDlg.ShowDialog(ownerForm)

        Dim retStr As String = ""

        If tempDlgResult <> DialogResult.Cancel Then
            retStr = tempComDlg.FileName
        End If

        tempComDlg.Dispose()

        Return retStr


    End Function

End Module
