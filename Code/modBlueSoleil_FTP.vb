'modBlueSoleil_FTP - Written by Jesse Yeager.    www.CompulsiveCode.com
'
'This module wraps the Blue Soleil SDK functions for using the File Transfer Profile.
'This is also the only Blue Soleil profile I've programmed for that has required the memory management function Btsdk_FreeMemory.
'
'Here's a quick explanation.  You can change folders, and you can Browse folders.  
'When navigating, \ is the delimiter.  The root is specified as "\"
'When browsing, each item in the directory will be returned in a separate event.  Events are FoundFolder and FoundFile depending on the type.
'I can't see any way to know for sure when the updating is complete.


Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

Module modBlueSoleil_FTP

    Private Const WIN32_FIND_DATA_StruSize As Integer = 318
    Private Const W32FD_MAX_PATH As Integer = 260
    Private Const FILE_ATTRIBUTE_DIRECTORY As Integer = &H10


    Private Const BTSDK_OK As UInt32 = 0
    Private Const BTSDK_TRUE As Byte = 1


    '/* Possible op_type member of Btsdk_FTPBrowseFolder */
    Private Const FTP_OP_REFRESH As Byte = 0
    Private Const FTP_OP_UPDIR As Byte = 1
    Private Const FTP_OP_NEXT As Byte = 2

    Private Const MAX_FTP_FOLDER_LIST As UInt32 = 128
    Private Const MAX_FILENAME As UInt32 = 256
    Private Const BTSDK_PATH_MAXLENGTH As UInt16 = 256     '/* Shall not be larger than FTP_MAX_PATH and OPP_MAX_PATH */


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncFTPuiShowBrowseFile(ByVal ptrFINDDATAbytes As IntPtr)
    Public delegateFTPuiShowBrowseFile As delfuncFTPuiShowBrowseFile = AddressOf BlueSoleil_FTP_Callback_UIshowBrowseFile

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncFTPstatusCallback(ByVal first As Byte, ByVal last As Byte, ByVal ptrFileNameBytes As IntPtr, ByVal fileSize As UInt32, ByVal curSize As UInt32)
    Public delegateFTPstatusCallback As delfuncPBAPstatusCallback = AddressOf BlueSoleil_FTP_Callback_Status

    Private BlueSoleil_FTP_Callback_Status_CurrSize As UInt32 = 0
    Private BlueSoleil_FTP_Callback_Status_TotalSize As UInt32 = 0

    Private BlueSoleil_FTP_CurrDir_Folders(0 To 0) As String
    Private BlueSoleil_FTP_CurrDir_FolderCount As Integer = 0
    Private BlueSoleil_FTP_CurrDir_Files(0 To 0) As String
    Private BlueSoleil_FTP_CurrDir_FileCount As Integer = 0

    Private BlueSoleil_FTP_CurrDir_AbsolutePath As String = ""  'i haven't programmed this yet.  phone problem got in the way.


    Public Event BlueSoleil_Event_FTP_FoundFolder(ByVal folderName As String)
    Public Event BlueSoleil_Event_FTP_FoundFile(ByVal fileName As String, ByVal fileSize As UInt64)



    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_FreeMemory(memBlock As UInt32)
    End Sub
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_FreeMemory(memBlock As IntPtr)
    End Sub

    '<DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    'Private Function Btsdk_FTPRegisterStatusCallback4ThirdParty(ByVal connHandle As UInt32, ByVal functPtr_FTP_STATUS_INFO_CB As IntPtr) As UInt32
    'End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_FTPRegisterStatusCallback4ThirdParty(ByVal connHandle As UInt32, ByVal functPtr_FTP_STATUS_INFO_CB As IntPtr)
    End Sub


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPGetRmtDir(ByVal connHandle As UInt32, ByRef remotePath As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPSetRmtDir(ByVal connHandle As UInt32, ByRef remotePath As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPPutFile(ByVal connHandle As UInt32, ByRef localFile As Byte, ByRef new_file As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPPutDir(ByVal connHandle As UInt32, ByRef localDir As Byte, ByRef new_dir As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPGetFile(ByVal connHandle As UInt32, ByRef remoteFile As Byte, ByRef new_file As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPGetDir(ByVal connHandle As UInt32, ByRef remoteDir As Byte, ByRef new_dir As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPDeleteFile(ByVal connHandle As UInt32, ByRef remoteFile As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPCreateDir(ByVal connHandle As UInt32, ByRef remotePath As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPDeleteDir(ByVal connHandle As UInt32, ByRef remotePath As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPCancelTransfer(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPBackDir(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FTPBrowseFolder(ByVal connHandle As UInt32, ByRef remotePath As Byte, ByVal ptrUIbrowseFile As UInt32, ByVal operationType As Byte) As UInt32
    End Function



    Private Sub BlueSoleil_FTP_Callback_Status(ByVal first As Byte, ByVal last As Byte, ByVal ptrFileNameBytes As IntPtr, ByVal theFileSize As UInt32, ByVal curSize As UInt32)

        'this callback is called by BlueSoleil during FTP transfers.

        Debug.Print("FTP Callback.  first = " & first & "   last = " & last & "   filesize = " & theFileSize & "    cursize = " & curSize)



        If last = 0 Then

            If first <> 0 Then
                'start of transfer.  
                BlueSoleil_FTP_Callback_Status_CurrSize = curSize
                BlueSoleil_FTP_Callback_Status_TotalSize = theFileSize

                Debug.Print("Start FTP xfer.  0 of " & theFileSize & " bytes.")

            Else
                BlueSoleil_FTP_Callback_Status_CurrSize = BlueSoleil_FTP_Callback_Status_CurrSize + curSize
                Debug.Print("Continue FTP xfer.  " & BlueSoleil_FTP_Callback_Status_CurrSize & " of " & theFileSize & " bytes.")
            End If


        Else

            'add final bytes.
            BlueSoleil_FTP_Callback_Status_CurrSize = BlueSoleil_FTP_Callback_Status_CurrSize + curSize
            Debug.Print("Finish FTP xfer.  " & BlueSoleil_FTP_Callback_Status_CurrSize & " of " & BlueSoleil_FTP_Callback_Status_TotalSize & " bytes.")

            'do whatever cuz we're done..

            'reset.
            ' BlueSoleil_FTP_Callback_Status_TotalSize = 0
            ' BlueSoleil_FTP_Callback_Status_CurrSize = 0

        End If

    End Sub


    Public Function BlueSoleil_FTP_RegisterStatusCallback(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        delegateFTPstatusCallback = AddressOf BlueSoleil_FTP_Callback_Status

        Dim retUInt As UInt32 = BTSDK_OK        'i thought Btsdk_FTPRegisterStatusCallback4ThirdParty returned a value, but I guess not.

        Dim functPtr_StatusCallback As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateFTPstatusCallback)
        Btsdk_FTPRegisterStatusCallback4ThirdParty(connHandle, functPtr_StatusCallback)

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_FTP_UnregisterStatusCallback(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        Dim retUInt As UInt32 = BTSDK_OK

        Dim functPtr_StatusCallback As IntPtr = IntPtr.Zero
        Btsdk_FTPRegisterStatusCallback4ThirdParty(connHandle, functPtr_StatusCallback)

        Return (retUInt = BTSDK_OK)

    End Function

    Public Function BlueSoleil_FTP_GetFile(ByVal connHandle As UInt32, ByVal rmtFile As String, ByVal localFile As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(rmtFile & Chr(0))

        Dim lclPathBytes(0 To 0) As Byte
        lclPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(localFile & Chr(0))

        Dim retGetFile As UInt32
        retGetFile = Btsdk_FTPGetFile(connHandle, rmtPathBytes(0), lclPathBytes(0))

        Return (retGetFile = 0)

    End Function




    Public Function BlueSoleil_FTP_PutFile(ByVal connHandle As UInt32, ByVal localFile As String, ByVal rmtFile As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(rmtFile & Chr(0))

        Dim lclPathBytes(0 To 0) As Byte
        lclPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(localFile & Chr(0))

        Dim retGetFile As UInt32
        retGetFile = Btsdk_FTPPutFile(connHandle, lclPathBytes(0), rmtPathBytes(0))

        Return (retGetFile = 0)

    End Function


    Public Function BlueSoleil_FTP_GetFolder(ByVal connHandle As UInt32, ByVal rmtDir As String, ByVal localDir As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(rmtDir & Chr(0))

        Dim lclPathBytes(0 To 0) As Byte
        lclPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(localDir & Chr(0))

        Dim retGetFile As UInt32
        retGetFile = Btsdk_FTPGetDir(connHandle, rmtPathBytes(0), lclPathBytes(0))

        Return (retGetFile = 0)

    End Function

    Public Function BlueSoleil_FTP_PutFolder(ByVal connHandle As UInt32, ByVal localDir As String, ByVal rmtDir As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(rmtDir & Chr(0))

        Dim lclPathBytes(0 To 0) As Byte
        lclPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(localDir & Chr(0))

        Dim retGetFile As UInt32
        retGetFile = Btsdk_FTPPutDir(connHandle, lclPathBytes(0), rmtPathBytes(0))

        Return (retGetFile = 0)

    End Function


    Public Function BlueSoleil_FTP_GetRemotePath(ByVal connHandle As UInt32, ByRef rmtPath As String) As Boolean

        rmtPath = ""

        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To BTSDK_PATH_MAXLENGTH) As Byte

        Dim retGetPath As UInt32
        retGetPath = Btsdk_FTPGetRmtDir(connHandle, rmtPathBytes(0))

        rmtPath = System.Text.Encoding.UTF8.GetString(rmtPathBytes)

        rmtPath = Replace(rmtPath, Chr(0), "")

        Debug.Print("BlueSoleil_FTP_GetRemotePath = " & rmtPath)

        Return (retGetPath = 0)

    End Function

    Public Function BlueSoleil_FTP_SetRemotePath(ByVal connHandle As UInt32, ByVal rmtPath As String, ByRef retFailDueToAccessDenied As Boolean) As Boolean

        retFailDueToAccessDenied = False
        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(rmtPath & Chr(0))

        Dim retSetPath As UInt32
        retSetPath = Btsdk_FTPSetRmtDir(connHandle, rmtPathBytes(0))

        'check retSetPath value for access error.
        If retSetPath = 1732 Then
            retFailDueToAccessDenied = True
        End If

        Return (retSetPath = 0)

    End Function


    Public Function BlueSoleil_FTP_CancelTransfer(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        Dim retUInt As UInt32
        retUInt = Btsdk_FTPCancelTransfer(connHandle)

        Return (retUInt = 0)

    End Function


    Public Function BlueSoleil_FTP_UpOneFolder(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        Dim prevDir As String = ""
        BlueSoleil_FTP_GetRemotePath(connHandle, prevDir)

        Dim retUInt As UInt32
        retUInt = Btsdk_FTPBackDir(connHandle)  'unfortunately, this seems to always return 1.  so we compare remote dir before and after changing dir.

        Dim newDir As String = ""
        BlueSoleil_FTP_GetRemotePath(connHandle, newDir)


        'Return (retUInt = 0)

        Return (newDir <> prevDir)

    End Function

    Public Function BlueSoleil_FTP_CreateDirectory(ByVal connHandle As UInt32, ByVal newDirName As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(newDirName & Chr(0))

        Dim retSetPath As UInt32
        retSetPath = Btsdk_FTPCreateDir(connHandle, rmtPathBytes(0))

        Return (retSetPath = 0)

    End Function

    Public Function BlueSoleil_FTP_DeleteDirectory(ByVal connHandle As UInt32, ByVal delDirName As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(delDirName & Chr(0))

        Dim retSetPath As UInt32
        retSetPath = Btsdk_FTPDeleteDir(connHandle, rmtPathBytes(0))

        Return (retSetPath = 0)

    End Function


    Public Function BlueSoleil_FTP_DeleteFile(ByVal connHandle As UInt32, ByVal delFileName As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        'specify remote FTP path
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(delFileName & Chr(0))

        Dim retSetPath As UInt32
        retSetPath = Btsdk_FTPDeleteFile(connHandle, rmtPathBytes(0))

        Return (retSetPath = 0)

    End Function


    Public Function BlueSoleil_FTP_BrowseFolder(ByVal connHandle As UInt32, ByVal currRemotePathName As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(currRemotePathName & Chr(0))

        delegateFTPuiShowBrowseFile = AddressOf BlueSoleil_FTP_Callback_UIshowBrowseFile
        Dim functPtr_UIshowBrowseFileCallback As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateFTPuiShowBrowseFile)
        Dim funcPtrINT As UInt32 = CUInt(functPtr_UIshowBrowseFileCallback)

        BlueSoleil_FTP_CurrDir_FolderCount = 0
        BlueSoleil_FTP_CurrDir_FileCount = 0

        Dim retUInt As UInt32
        retUInt = Btsdk_FTPBrowseFolder(connHandle, rmtPathBytes(0), funcPtrINT, FTP_OP_REFRESH)

        Debug.Print("BlueSoleil_FTP_BrowseFolder return")

        Return (retUInt = 0)

    End Function




    Private Sub BlueSoleil_Stru_WIN32FINDDATA_GetInfo(ByRef udtBytes() As Byte, ByVal startIdx As Integer, ByRef udtFileAttributes As UInt32, ByRef udtFileName As String, ByRef udtFTCreationTimeLOW As UInt32, ByRef udtFTCreationTimeHIGH As UInt32, ByRef udtFTAccessTimeLOW As UInt32, ByRef udtFTAccessTimeHIGH As UInt32, ByRef udtFTModifyTimeLOW As UInt32, ByRef udtFTModifyTimeHIGH As UInt32, ByRef udtFileSize As UInt64)

        '44 + 260 + 14 = 318


        udtFileAttributes = BitConverter.ToUInt32(udtBytes, startIdx + 0)

        udtFTCreationTimeLOW = BitConverter.ToUInt32(udtBytes, startIdx + 4)
        udtFTCreationTimeHIGH = BitConverter.ToUInt32(udtBytes, startIdx + 8)

        udtFTAccessTimeLOW = BitConverter.ToUInt32(udtBytes, startIdx + 12)
        udtFTAccessTimeHIGH = BitConverter.ToUInt32(udtBytes, startIdx + 16)

        udtFTModifyTimeLOW = BitConverter.ToUInt32(udtBytes, startIdx + 20)
        udtFTModifyTimeHIGH = BitConverter.ToUInt32(udtBytes, startIdx + 24)

        udtFileSize = BitConverter.ToUInt32(udtBytes, startIdx + 28 + 4)    'not sure this is populated correctly.


        Dim i As Integer
        Dim retStr As String
        retStr = ""
        For i = startIdx + 44 To startIdx + (44 + W32FD_MAX_PATH - 1)
            If udtBytes(i) = 0 Then Exit For
            retStr = retStr & Chr(udtBytes(i))
        Next i

        udtFileName = retStr

    End Sub


    Private Sub BlueSoleil_FTP_Callback_UIshowBrowseFile(ByVal ptrFINDDATA As IntPtr)

        Debug.Print("FTP UIshowBrowseFile")

        If ptrFINDDATA = IntPtr.Zero Then
            'does this fire when browsing is complete?  if so, we can add a BrowseComplete event.
            Exit Sub
        End If

        Dim w32fdBytes(0 To WIN32_FIND_DATA_StruSize - 1) As Byte

        Marshal.Copy(ptrFINDDATA, w32fdBytes, 0, w32fdBytes.Length)

        Dim fdSize As UInt64
        Dim fdFileName As String = ""
        Dim fdAttribs As UInt32
        BlueSoleil_Stru_WIN32FINDDATA_GetInfo(w32fdBytes, 0, fdAttribs, fdFileName, 0, 0, 0, 0, 0, 0, fdSize)


        'pretty lame that we receive info for each file / folder separately and in no order.  and we don't know when we have it all.

        If (fdAttribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            Debug.Print("FTP UIshowBrowseFile object is Directory - " & fdFileName)

            ReDim Preserve BlueSoleil_FTP_CurrDir_Folders(0 To BlueSoleil_FTP_CurrDir_FolderCount)
            BlueSoleil_FTP_CurrDir_Folders(BlueSoleil_FTP_CurrDir_FolderCount) = fdFileName
            BlueSoleil_FTP_CurrDir_FolderCount = BlueSoleil_FTP_CurrDir_FolderCount + 1

            RaiseEvent BlueSoleil_Event_FTP_FoundFolder(fdFileName)

        Else
            Debug.Print("FTP UIshowBrowseFile object is File - " & fdFileName)

            ReDim Preserve BlueSoleil_FTP_CurrDir_Files(0 To BlueSoleil_FTP_CurrDir_FileCount)
            BlueSoleil_FTP_CurrDir_Files(BlueSoleil_FTP_CurrDir_FileCount) = fdFileName
            BlueSoleil_FTP_CurrDir_FileCount = BlueSoleil_FTP_CurrDir_FileCount + 1

            RaiseEvent BlueSoleil_Event_FTP_FoundFile(fdFileName, fdSize)

        End If


        Btsdk_FreeMemory(ptrFINDDATA)   'this is what the SDK says to do.


    End Sub

End Module
