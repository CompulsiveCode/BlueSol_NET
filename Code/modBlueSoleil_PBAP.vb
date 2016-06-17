'modBlueSoleil_PBAP - Written by Jesse Yeager.    www.CompulsiveCode.com
'
'This module wraps the Blue Soleil SDK functions for using the Phonebook Access Profile.
'
'This module wraps the Win32 File IO routines so BlueSoleil can access them in a platform-agnostic way.  See the BSFileIO function(s).
'
'My only interest is in the Contacts / Phone Book.   However, this can also be used to retrieve missed calls, incoming calls, outgoing calls, etc.
'
'

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

Module modBlueSoleil_PBAP

    Private Const BTSDK_OK As UInt32 = 0
    Private Const BTSDK_TRUE As Byte = 1



    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPWriteFile(ByVal fHandle As UInt32, ByVal arrayPtr As IntPtr, ByVal arrayLen As UInt32) As UInt32
    Public delegateAPPWriteFile As delfuncAPPWriteFile = AddressOf BSfileIO_APP_WriteFile

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncPBAPstatusCallback(ByVal first As Byte, ByVal last As Byte, ByVal ptrFileNameBytes As IntPtr, ByVal fileSize As UInt32, ByVal curSize As UInt32)
    Public delegatePBAPstatusCallback As delfuncPBAPstatusCallback = AddressOf BlueSoleil_PBAP_Callback_Status



    Private BlueSoleil_PBAP_Callback_Status_CurrSize As UInt32 = 0
    Private BlueSoleil_PBAP_Callback_Status_TotalSize As UInt32 = 0

    Private BSfileIO_APP_CurrentDir As String = IO.Directory.GetCurrentDirectory        'couldn't hurt to keep track of this.


    '/*Phone Book Access Profile*/
    Private Const BTSDK_PBAP_MAX_DELIMITER As Byte = &H02

    '/* Possible values for member 'order' in _BtSdkPBAPParamStru */
    Private Const BTSDK_PBAP_ORDER_INDEXED As Byte = &H00
    Private Const BTSDK_PBAP_ORDER_NAME As Byte = &H01
    Private Const BTSDK_PBAP_ORDER_PHONETIC As Byte = &H02

    '/* Possible flags for member 'mask' in _BtSdkPBAPParamStru */
    Private Const BTSDK_PBAP_PM_ORDER As UInt16 = &H0001
    Private Const BTSDK_PBAP_PM_SCHVALUE As UInt16 = &H0002
    Private Const BTSDK_PBAP_PM_SCHATTR As UInt16 = &H0004
    Private Const BTSDK_PBAP_PM_MAXCOUNT As UInt16 = &H0008
    Private Const BTSDK_PBAP_PM_LISTOFFSET As UInt16 = &H0010
    Private Const BTSDK_PBAP_PM_FILTER As UInt16 = &H0020
    Private Const BTSDK_PBAP_PM_FORMAT As UInt16 = &H0040
    Private Const BTSDK_PBAP_PM_PBSIZE As UInt16 = &H0080
    Private Const BTSDK_PBAP_PM_MISSEDCALLS As UInt16 = &H0100

    '/* Possible values for member 'format' in _BtSdkPBAPParamStru */
    Private Const BTSDK_PBAP_FMT_VCARD21 As Byte = &H00
    Private Const BTSDK_PBAP_FMT_VCARD30 As Byte = &H01

    Private Const BTSDK_PBAP_REPO_LOCAL As Byte = &H01
    Private Const BTSDK_PBAP_REPO_SIM As Byte = &H02

    '/* Filter bit mask supported by PBAP1.0. Possible values for parameter 'flag' of Btsdk_PBAPFilterComposer. */
    Private Const BTSDK_PBAP_FILTER_VERSION As Byte = &H00
    Private Const BTSDK_PBAP_FILTER_FN As Byte = &H01
    Private Const BTSDK_PBAP_FILTER_N As Byte = &H02
    Private Const BTSDK_PBAP_FILTER_PHOTO As Byte = &H03
    Private Const BTSDK_PBAP_FILTER_BDAY As Byte = &H04
    Private Const BTSDK_PBAP_FILTER_ADR As Byte = &H05
    Private Const BTSDK_PBAP_FILTER_LABEL As Byte = &H06
    Private Const BTSDK_PBAP_FILTER_TEL As Byte = &H07
    Private Const BTSDK_PBAP_FILTER_EMAIL As Byte = &H08
    Private Const BTSDK_PBAP_FILTER_MAILER As Byte = &H09
    Private Const BTSDK_PBAP_FILTER_TZ As Byte = &H0A
    Private Const BTSDK_PBAP_FILTER_GEO As Byte = &H0B
    Private Const BTSDK_PBAP_FILTER_TITLE As Byte = &H0C
    Private Const BTSDK_PBAP_FILTER_ROLE As Byte = &H0D
    Private Const BTSDK_PBAP_FILTER_LOGO As Byte = &H0E
    Private Const BTSDK_PBAP_FILTER_AGENT As Byte = &H0F
    Private Const BTSDK_PBAP_FILTER_ORG As Byte = &H10
    Private Const BTSDK_PBAP_FILTER_NOTE As Byte = &H11
    Private Const BTSDK_PBAP_FILTER_REV As Byte = &H12
    Private Const BTSDK_PBAP_FILTER_SOUND As Byte = &H13
    Private Const BTSDK_PBAP_FILTER_URL As Byte = &H14
    Private Const BTSDK_PBAP_FILTER_UID As Byte = &H15
    Private Const BTSDK_PBAP_FILTER_KEY As Byte = &H16
    Private Const BTSDK_PBAP_FILTER_NICKNAME As Byte = &H17
    Private Const BTSDK_PBAP_FILTER_CATEGORIES As Byte = &H18
    Private Const BTSDK_PBAP_FILTER_PROID As Byte = &H19
    Private Const BTSDK_PBAP_FILTER_CLASS As Byte = &H1A
    Private Const BTSDK_PBAP_FILTER_SORT_STRING As Byte = &H1B
    Private Const BTSDK_PBAP_FILTER_X_IRMC_CALL_DATETIME As Byte = &H1C
    Private Const BTSDK_PBAP_FILTER_PROPRIETARY_FILTER As Byte = &H27
    Private Const BTSDK_PBAP_FILTER_INVALID As Byte = &HFF


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PBAPRegisterStatusCallback(ByVal connHandle As UInt32, ByVal functPtr_PBAP_STATUS_INFO_CB As IntPtr) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PBAPRegisterFileIORoutines(ByVal connHandle As UInt32, ByRef PBAPFileIORoutinesStru As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PBAPPullPhoneBook(ByVal connHandle As UInt32, ByRef rmtPath As Byte, ByRef PBtSdkPBAPParamStru As Byte, ByVal file_hdl_towriteto As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PBAPFilterComposer(ByRef rmtPath As Byte, ByVal filterFlagToSet As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PBAPCancelTransfer(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PBAPSetPath(ByVal connHandle As UInt32, ByRef path As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PBAPPullCardList(ByVal connHandle As UInt32, ByRef rmtFolder As Byte, ByRef PBtSdkPBAPParamStru As Byte, ByVal file_hdl_towriteto As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PBAPPullCardEntry(ByVal connHandle As UInt32, ByRef rmtName As Byte, ByRef PBtSdkPBAPParamStru As Byte, ByVal file_hdl_towriteto As UInt32) As UInt32
    End Function



    Public Function BlueSoleil_PBAP_PullPhoneBook(ByVal connHandle As UInt32, ByVal localVCFfileToWrite As String, Optional ByVal pbType As String = "CONTACTS") As Boolean

        If connHandle = 0 Then Return False     'no connection?  nothing to do.  fail out.


        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)

        Dim retUInt As UInt32

        'initialize IO routine pointer(s) for BS to use to write the VCF file.  we only need to provide a pointer to APP_WriteFile (which wraps the Win32 WriteFile API)
        Dim pbapIOroutinesBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapFileIOroutines(pbapIOroutinesBytes, CUInt(functPtr_APPwriteFile))
        retUInt = Btsdk_PBAPRegisterFileIORoutines(connHandle, pbapIOroutinesBytes(0))
        If retUInt <> 0 Then
            'some error?
            Return False
        End If
        BlueSoleil_PBAP_RegisterStatusCallback(connHandle)


        'delete the local vcf file if it already exists.
        Try
            If IO.File.Exists(localVCFfileToWrite) = True Then
                IO.File.Delete(localVCFfileToWrite)
            End If
        Catch ex As Exception
            'meh.  probably need to fail here.
            Return False
        End Try

        'create / open the file (using Win32 CreateFile API), getting a file handle (old school).  
        Dim fHandle As IntPtr = FileAPI_OpenFile(localVCFfileToWrite, True)
        Dim fHandleUint As UInt32 = CUInt(fHandle)

        'specify default remote PBAP path (this is not a folder structure path, it's a heirarchy in the PBAP protocol).
        'note:  For multiple-sim devices, you can also do SIM1/telecom/pb.vcf
        Dim vcfRemotePath As String = "telecom/pb.vcf"
        pbType = UCase(Trim(pbType))
        If pbType = "CONTACTS" Then
            vcfRemotePath = "telecom/pb.vcf"  'Main Phone book
        ElseIf pbType = "INCOMING" Then
            vcfRemotePath = "telecom/ich.vcf" 'Incoming Calls History
        ElseIf pbType = "OUTGOING" Then
            vcfRemotePath = "telecom/och.vcf" 'Outgoing Calls History
        ElseIf pbType = "MISSED" Then
            vcfRemotePath = "telecom/mch.vcf" 'Missed Calls History
        ElseIf pbType = "HISTORY" Then
            vcfRemotePath = "telecom/cch.vcf" 'Combined Calls History
        End If


        Dim vcfRemotePathBytes(0 To 0) As Byte
        vcfRemotePathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfRemotePath & Chr(0))

        'build structure to specify retrieval parameters (if any).  FINALLY GOT THIS WORKING!!
        Dim pbapParamBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapParam(pbapParamBytes, True)

        'tell BS to get the phonebook, using the IO routines to write the VCF file to the file handle.
        retUInt = Btsdk_PBAPPullPhoneBook(connHandle, vcfRemotePathBytes(0), pbapParamBytes(0), fHandleUint)

        'close the file we created.  calls Win32api CloseHandle.  could check size to see if anything was written.
        Dim outSize As Long = FileAPI_GetFileSize(fHandle)

        FileAPI_CloseFile(fHandle)

        If outSize < 10 Then    'return false for empty file.
            retUInt = 1
        End If

        'done.  Now go parse that VCard file!
        Return (retUInt = 0)

    End Function


    Public Function BlueSoleil_PBAP_PullPhoneBook_ByPath(ByVal connHandle As UInt32, ByVal localVCFfileToWrite As String, ByVal vcfRemotePath As String) As Boolean

        If connHandle = 0 Then Return False     'no connection?  nothing to do.  fail out.


        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)

        Dim retUInt As UInt32

        'initialize IO routine pointer(s) for BS to use to write the VCF file.  we only need to provide a pointer to APP_WriteFile (which wraps the Win32 WriteFile API)
        Dim pbapIOroutinesBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapFileIOroutines(pbapIOroutinesBytes, CUInt(functPtr_APPwriteFile))
        retUInt = Btsdk_PBAPRegisterFileIORoutines(connHandle, pbapIOroutinesBytes(0))
        If retUInt <> 0 Then
            'some error?
            Return False
        End If
        BlueSoleil_PBAP_RegisterStatusCallback(connHandle)


        'delete the local vcf file if it already exists.
        Try
            If IO.File.Exists(localVCFfileToWrite) = True Then
                IO.File.Delete(localVCFfileToWrite)
            End If
        Catch ex As Exception
            'meh.  probably need to fail here.
            Return False
        End Try

        'create / open the file (using Win32 CreateFile API), getting a file handle (old school).  
        Dim fHandle As IntPtr = FileAPI_OpenFile(localVCFfileToWrite, True)
        Dim fHandleUint As UInt32 = CUInt(fHandle)



        Dim vcfRemotePathBytes(0 To 0) As Byte
        vcfRemotePathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfRemotePath & Chr(0))

        'build structure to specify retrieval parameters (if any).  FINALLY GOT THIS WORKING!!
        Dim pbapParamBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapParam(pbapParamBytes, True)

        'tell BS to get the phonebook, using the IO routines to write the VCF file to the file handle.
        retUInt = Btsdk_PBAPPullPhoneBook(connHandle, vcfRemotePathBytes(0), pbapParamBytes(0), fHandleUint)

        'close the file we created.  calls Win32api CloseHandle.  could check size to see if anything was written.
        Dim outSize As Long = FileAPI_GetFileSize(fHandle)

        FileAPI_CloseFile(fHandle)

        If outSize < 10 Then    'return false for empty file.
            retUInt = 1
        End If

        'done.  Now go parse that VCard file!
        Return (retUInt = 0)

    End Function

    Public Function BlueSoleil_PBAP_PullCardList(ByVal connHandle As UInt32, ByVal localXMLfileToWrite As String) As Boolean

        'you dont need this function unless you want to pull individual cards for some reason.


        If connHandle = 0 Then Return False     'no connection?  nothing to do.  fail out.


        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)

        Dim retUInt As UInt32

        'initialize IO routine pointer(s) for BS to use to write the VCF file.  we only need to provide a pointer to APP_WriteFile (which wraps the Win32 WriteFile API)
        Dim pbapIOroutinesBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapFileIOroutines(pbapIOroutinesBytes, CUInt(functPtr_APPwriteFile))
        retUInt = Btsdk_PBAPRegisterFileIORoutines(connHandle, pbapIOroutinesBytes(0))
        If retUInt <> 0 Then
            'some error?
            Return False
        End If
        BlueSoleil_PBAP_RegisterStatusCallback(connHandle)


        'delete the local XML file if it already exists.
        Try
            If IO.File.Exists(localXMLfileToWrite) = True Then
                IO.File.Delete(localXMLfileToWrite)
            End If
        Catch ex As Exception
            'meh.  probably need to fail here.
            Return False
        End Try

        'create / open the file (using Win32 CreateFile API), getting a file handle (old school).  
        Dim fHandle As IntPtr = FileAPI_OpenFile(localXMLfileToWrite, True)
        Dim fHandleUint As UInt32 = CUInt(fHandle)

        'specify default remote PBAP path (this is not a folder structure path, it's a heirarchy in the PBAP protocol).
        Dim vcfRemotePath As String = "telecom/pb.vcf"
        Dim vcfRemotePathBytes(0 To 0) As Byte
        vcfRemotePathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfRemotePath & Chr(0))

        'build structure to specify retrieval parameters (if any).
        Dim pbapParamBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapParam(pbapParamBytes, True)

        'tell BS to get the phonebook, using the IO routines to write the VCF file to the file handle.
        retUInt = Btsdk_PBAPPullCardList(connHandle, vcfRemotePathBytes(0), pbapParamBytes(0), fHandleUint)

        'close the file we created.  calls Win32api CloseHandle
        FileAPI_CloseFile(fHandle)

        'done.  Now go parse that VCardList XML file!
        Return (retUInt = 0)


    End Function


    Public Function BlueSoleil_PBAP_PullCardList_ByPath(ByVal connHandle As UInt32, ByVal localXMLfileToWrite As String, ByVal vcfRemotePath As String) As Boolean

        'you dont need this function unless you want to pull individual cards for some reason.


        If connHandle = 0 Then Return False     'no connection?  nothing to do.  fail out.


        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)

        Dim retUInt As UInt32

        'initialize IO routine pointer(s) for BS to use to write the VCF file.  we only need to provide a pointer to APP_WriteFile (which wraps the Win32 WriteFile API)
        Dim pbapIOroutinesBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapFileIOroutines(pbapIOroutinesBytes, CUInt(functPtr_APPwriteFile))
        retUInt = Btsdk_PBAPRegisterFileIORoutines(connHandle, pbapIOroutinesBytes(0))
        If retUInt <> 0 Then
            'some error?
            Return False
        End If
        BlueSoleil_PBAP_RegisterStatusCallback(connHandle)


        'delete the local XML file if it already exists.
        Try
            If IO.File.Exists(localXMLfileToWrite) = True Then
                IO.File.Delete(localXMLfileToWrite)
            End If
        Catch ex As Exception
            'meh.  probably need to fail here.
            Return False
        End Try

        'create / open the file (using Win32 CreateFile API), getting a file handle (old school).  
        Dim fHandle As IntPtr = FileAPI_OpenFile(localXMLfileToWrite, True)
        Dim fHandleUint As UInt32 = CUInt(fHandle)

        'specify default remote PBAP path (this is not a folder structure path, it's a heirarchy in the PBAP protocol).
        Dim vcfRemotePathBytes(0 To 0) As Byte
        vcfRemotePathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfRemotePath & Chr(0))

        'build structure to specify retrieval parameters (if any).
        Dim pbapParamBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapParam(pbapParamBytes, True)

        'tell BS to get the phonebook, using the IO routines to write the VCF file to the file handle.
        retUInt = Btsdk_PBAPPullCardList(connHandle, vcfRemotePathBytes(0), pbapParamBytes(0), fHandleUint)

        'close the file we created.  calls Win32api CloseHandle
        FileAPI_CloseFile(fHandle)

        'done.  Now go parse that VCardList XML file!
        Return (retUInt = 0)


    End Function

    Private Sub BlueSoleil_PBAP_Callback_Status(ByVal first As Byte, ByVal last As Byte, ByVal ptrFileNameBytes As IntPtr, ByVal theFileSize As UInt32, ByVal curSize As UInt32)

        'this callback is called by BlueSoleil during PBAP transfers.

        Debug.Print("PBAP Callback.  first = " & first & "   last = " & last & "   filesize = " & theFileSize & "    cursize = " & curSize)



        If last = 0 Then

            If first <> 0 Then
                'start of transfer.  
                BlueSoleil_PBAP_Callback_Status_CurrSize = curSize
                BlueSoleil_PBAP_Callback_Status_TotalSize = theFileSize

                Debug.Print("Start PBAP xfer.  0 of " & theFileSize & " bytes.")

            Else
                BlueSoleil_PBAP_Callback_Status_CurrSize = BlueSoleil_PBAP_Callback_Status_CurrSize + curSize
                Debug.Print("Continue PBAP xfer.  " & BlueSoleil_PBAP_Callback_Status_CurrSize & " of " & theFileSize & " bytes.")
            End If


        Else

            'add final bytes.
            BlueSoleil_PBAP_Callback_Status_CurrSize = BlueSoleil_PBAP_Callback_Status_CurrSize + curSize
            Debug.Print("Finish PBAP xfer.  " & BlueSoleil_PBAP_Callback_Status_CurrSize & " of " & BlueSoleil_PBAP_Callback_Status_TotalSize & " bytes.")

            'do whatever cuz we're done..

            'RaiseEvent BlueSoleil_Event_PBAP_TransferComplete()
            ' Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_PBAP_TransferComplete())
            ' t.Start()

            'reset.
            ' BlueSoleil_PBAP_Callback_Status_TotalSize = 0
            ' BlueSoleil_PBAP_Callback_Status_CurrSize = 0

        End If

    End Sub



    Public Function BlueSoleil_PBAP_PullCard(ByVal connHandle As UInt32, ByVal remoteCardHandle As String, ByVal localVCFfileToWrite As String) As Boolean

        'you dont need this function unless you want to pull individual cards for some reason.

        If connHandle = 0 Then Return False     'no connection?  nothing to do.  fail out.

        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)

        Dim retUInt As UInt32

        'initialize IO routine pointer(s) for BS to use to write the VCF file.  we only need to provide a pointer to APP_WriteFile (which wraps the Win32 WriteFile API)
        Dim pbapIOroutinesBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapFileIOroutines(pbapIOroutinesBytes, CUInt(functPtr_APPwriteFile))
        retUInt = Btsdk_PBAPRegisterFileIORoutines(connHandle, pbapIOroutinesBytes(0))
        If retUInt <> 0 Then
            'some error?
            Return False
        End If
        BlueSoleil_PBAP_RegisterStatusCallback(connHandle)

        'delete the local vcf file if it already exists.
        Try
            If IO.File.Exists(localVCFfileToWrite) = True Then
                IO.File.Delete(localVCFfileToWrite)
            End If
        Catch ex As Exception
            'meh.  probably need to fail here.
            Return False
        End Try

        'create / open the file (using Win32 CreateFile API), getting a file handle (old school).  
        Dim fHandle As IntPtr = FileAPI_OpenFile(localVCFfileToWrite, True)
        Dim fHandleUint As UInt32 = CUInt(fHandle)

        'specify default remote PBAP path (this is not a folder structure path, it's a heirarchy in the PBAP protocol).
        Dim vcfRemoteFile As String = remoteCardHandle
        Dim vcfRemoteFileBytes(0 To 0) As Byte
        vcfRemoteFileBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfRemoteFile & Chr(0))

        'build structure to specify retrieval parameters (if any).
        Dim pbapParamBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapParam(pbapParamBytes, True)


        Dim rmtPathTorF As Boolean = False
        rmtPathTorF = BlueSoleil_PBAP_SetRemotePath(connHandle, "..")
        rmtPathTorF = BlueSoleil_PBAP_SetRemotePath(connHandle, "..")
        rmtPathTorF = BlueSoleil_PBAP_SetRemotePath(connHandle, "telecom")
        rmtPathTorF = BlueSoleil_PBAP_SetRemotePath(connHandle, "pb")

        'tell BS to get the phonebook, using the IO routines to write the VCF file to the file handle.
        retUInt = Btsdk_PBAPPullCardEntry(connHandle, vcfRemoteFileBytes(0), pbapParamBytes(0), fHandleUint)

        'close the file we created.  calls Win32api CloseHandle
        FileAPI_CloseFile(fHandle)

        'done.  Now go parse that VCard file!
        Return (retUInt = 0)

    End Function



    Public Function BlueSoleil_PBAP_PullCard_ByPath(ByVal connHandle As UInt32, ByVal remoteCardHandle As String, ByVal localVCFfileToWrite As String, ByVal vcfRemotePath As String) As Boolean

        'you dont need this function unless you want to pull individual cards for some reason.

        vcfRemotePath = Replace(vcfRemotePath, ".VCF", "", 1, -1, CompareMethod.Text)

        If connHandle = 0 Then Return False     'no connection?  nothing to do.  fail out.

        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)

        Dim retUInt As UInt32

        'initialize IO routine pointer(s) for BS to use to write the VCF file.  we only need to provide a pointer to APP_WriteFile (which wraps the Win32 WriteFile API)
        Dim pbapIOroutinesBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapFileIOroutines(pbapIOroutinesBytes, CUInt(functPtr_APPwriteFile))
        retUInt = Btsdk_PBAPRegisterFileIORoutines(connHandle, pbapIOroutinesBytes(0))
        If retUInt <> 0 Then
            'some error?
            Return False
        End If
        BlueSoleil_PBAP_RegisterStatusCallback(connHandle)

        'delete the local vcf file if it already exists.
        Try
            If IO.File.Exists(localVCFfileToWrite) = True Then
                IO.File.Delete(localVCFfileToWrite)
            End If
        Catch ex As Exception
            'meh.  probably need to fail here.
            Return False
        End Try

        'create / open the file (using Win32 CreateFile API), getting a file handle (old school).  
        Dim fHandle As IntPtr = FileAPI_OpenFile(localVCFfileToWrite, True)
        Dim fHandleUint As UInt32 = CUInt(fHandle)

        'specify default remote PBAP path (this is not a folder structure path, it's a heirarchy in the PBAP protocol).
        Dim vcfRemoteFile As String = remoteCardHandle
        Dim vcfRemoteFileBytes(0 To 0) As Byte
        vcfRemoteFileBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfRemoteFile & Chr(0))

        'build structure to specify retrieval parameters (if any).
        Dim pbapParamBytes(0 To 0) As Byte
        BlueSoleil_PBAP_InitStruBytes_pbapParam(pbapParamBytes, True)


        Dim pathItems(0 To 0) As String
        pathItems = Split(vcfRemotePath, "/")

        Dim rmtPathTorF As Boolean = False
        rmtPathTorF = BlueSoleil_PBAP_SetRemotePath(connHandle, "..")
        rmtPathTorF = BlueSoleil_PBAP_SetRemotePath(connHandle, "..")
        rmtPathTorF = BlueSoleil_PBAP_SetRemotePath(connHandle, "..")

        Dim i As Integer
        For i = 0 To pathItems.Length - 1
            If pathItems(i) <> "" Then
                rmtPathTorF = BlueSoleil_PBAP_SetRemotePath(connHandle, pathItems(i))
            End If
        Next i


        'tell BS to get the phonebook, using the IO routines to write the VCF file to the file handle.
        retUInt = Btsdk_PBAPPullCardEntry(connHandle, vcfRemoteFileBytes(0), pbapParamBytes(0), fHandleUint)

        'close the file we created.  calls Win32api CloseHandle
        FileAPI_CloseFile(fHandle)

        'done.  Now go parse that VCard file!
        Return (retUInt = 0)

    End Function

    Private Function BlueSoleil_PBAP_SetRemotePath(ByVal connHandle As UInt32, ByVal rmtPath As String) As Boolean

        'you dont need this function unless you want to pull individual cards for some reason.


        'specify default remote PBAP path (this is not a folder structure path, it's a heirarchy in the PBAP protocol).
        Dim rmtPathBytes(0 To 0) As Byte
        rmtPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(rmtPath & Chr(0))

        Dim retSetPath As UInt32
        retSetPath = Btsdk_PBAPSetPath(connHandle, rmtPathBytes(0))

        Return (retSetPath = 0)

    End Function

    Private Sub BlueSoleil_PBAP_InitStruBytes_pbapFileIOroutines(ByRef inpByteArray() As Byte, ByVal functPtr_APPWriteFile As UInt32)

        Dim sizeOfStru As Integer = 7 * 4

        ReDim inpByteArray(0 To sizeOfStru - 1)

        Dim tempBytes(0 To 3) As Byte
        tempBytes = BitConverter.GetBytes(functPtr_APPWriteFile)

        Array.Copy(tempBytes, 0, inpByteArray, 8, 4)

    End Sub

    Private Sub BlueSoleil_PBAP_InitStruBytes_pbapParam(ByRef inpByteArray() As Byte, ByVal getVcard30 As Boolean)

        Dim sizeOfStru As Integer = 2 + 8 + 2 + 2 + 1 + 1 + 4 + 1 + 1 + 2     ' 2 + 8 + 2 + 2 + 1 + 1 + 4 + 1 + 1 + 2

        sizeOfStru = sizeOfStru * 2 'not sure why i put this here.  get rid.

        ReDim inpByteArray(0 To sizeOfStru - 1)


        'set mask and filter.
        If getVcard30 = True Then

            ' i thought this value should go in byte (1) since it's a 16bit value and im just writing the lower byte, 
            ' but for whatever reason, putting it in byte (0) makes it work.  :-|
            inpByteArray(0) = CByte(BTSDK_PBAP_PM_FILTER Or BTSDK_PBAP_PM_FORMAT)   'Or BTSDK_PBAP_PM_ORDER)     'removed since it is not supported by all phones.

            ' byte (14) is where the ORDER flag is located within the structure.
            inpByteArray(14) = 0        'BTSDK_PBAP_ORDER_NAME      'removed since it is not supported by all phones.


            ' byte (15) is where the FORMAT flag is located within the structure.
            inpByteArray(15) = BTSDK_PBAP_FMT_VCARD30

            ' we could have manually populated this 64bit bitfield, but Blue Soleil provides a 'composer' function.  
            ' We pass byte (2) because it's the first byte of the FILTER within the structure.

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_VERSION)    'reqd
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_FN)         'reqd
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_N)          'reqd
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_PHOTO)      'yay!
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_BDAY)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_ADR)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_LABEL)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_TEL)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_EMAIL)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_MAILER)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_TZ)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_GEO)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_TITLE)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_ROLE)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_LOGO)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_AGENT)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_ORG)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_NOTE)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_REV)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_SOUND)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_URL)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_UID)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_KEY)
            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_NICKNAME)

            Btsdk_PBAPFilterComposer(inpByteArray(2), BTSDK_PBAP_FILTER_X_IRMC_CALL_DATETIME)

        End If


    End Sub




    Private Function BSfileIO_APP_WriteFile(ByVal fHandle As UInt32, ByVal arrayPtr As IntPtr, ByVal arrayLen As UInt32) As UInt32

        Debug.Print("BSfileIO_WriteFile Len = " & arrayLen & "  Ptr = " & arrayPtr.ToInt64 & "  handle = " & fHandle)

        If fHandle = 0 Then
            Return 0
        End If

        If arrayLen = 0 Then        'this is important.  because we return number of "elements" written, and we wrote one element of zero-length.  whatever.
            Return 1
        End If

        Dim tempBytesArray(0 To 0) As Byte
        ReDim tempBytesArray(0 To CInt(arrayLen - 1))

        System.Runtime.InteropServices.Marshal.Copy(arrayPtr, tempBytesArray, 0, CInt(arrayLen))

        Dim TorF As Boolean = FileAPI_PutBytes(CType(fHandle, IntPtr), -1, tempBytesArray.Length, tempBytesArray)

        If TorF = True Then
            Return arrayLen
        Else
            Return 0
        End If

    End Function


    Private Function BlueSoleil_PBAP_RegisterStatusCallback(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        delegatePBAPstatusCallback = AddressOf BlueSoleil_PBAP_Callback_Status

        Dim retUInt As UInt32
        Dim functPtr_StatusCallback As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegatePBAPstatusCallback)
        retUInt = Btsdk_PBAPRegisterStatusCallback(connHandle, functPtr_StatusCallback)

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_PBAP_XML_GetCardListInfo(ByVal fnMsgList As String, ByRef retCardHandles() As String, ByRef retCardNames() As String) As Integer

        'returns the number of messages.

        'read XML file.

        Dim retCardCount As Integer = 0

        ReDim retCardHandles(0 To 0)
        ReDim retCardNames(0 To 0)

        Dim inpXMLreader As New Xml.XmlTextReader(fnMsgList)

        Dim lineNodeName As String = ""
        Dim lineNodeText As String = ""

        Dim inDURATION As Boolean = False
        Dim inDISTANCE As Boolean = False
        Dim inSTEP As Boolean = False

        Dim outItemCount As Integer = 0

        Dim tempStr As String = ""

        Dim TorF As Boolean = True
        Do

            TorF = False
            Try
                TorF = inpXMLreader.Read
            Catch ex As Exception
                If InStr(1, ex.Message, "find file", CompareMethod.Text) <> 0 Then
                    'searching for some stylesheet or some nonsense.
                    TorF = True
                Else
                    If outItemCount = 0 Then outItemCount = -1
                    Exit Do
                End If
            End Try
            If TorF = False Then Exit Do



            If inpXMLreader.NodeType = System.Xml.XmlNodeType.Element Then
                lineNodeName = inpXMLreader.Name
                lineNodeName = UCase(lineNodeName)

                Select Case lineNodeName


                    Case "CARD"

                        ReDim Preserve retCardHandles(0 To retCardCount)
                        ReDim Preserve retCardNames(0 To retCardCount)

                        retCardNames(retCardCount) = inpXMLreader.GetAttribute("name")
                        retCardHandles(retCardCount) = inpXMLreader.GetAttribute("handle")

                        retCardCount = retCardCount + 1

                End Select


            End If

        Loop


        inpXMLreader.Close()

        Return retCardCount

    End Function



End Module
