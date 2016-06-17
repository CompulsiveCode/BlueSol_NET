'modBlueSoleil_OPP - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'This module wraps the Blue Soleil SDK functions for using the Object Push (Obex) Profile.

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices
Module modBlueSoleil_OPP

    Private Const BTSDK_OK As UInt32 = 0
    Private Const BTSDK_TRUE As Byte = 1



    Private BlueSoleil_OPP_Callback_Status_CurrSize As UInt32 = 0
    Private BlueSoleil_OPP_Callback_Status_TotalSize As UInt32 = 0


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncOPPstatusCallback(ByVal first As Byte, ByVal last As Byte, ByVal ptrFileNameBytes As IntPtr, ByVal fileSize As UInt32, ByVal curSize As UInt32)
    Public delegateOPPstatusCallback As delfuncOPPstatusCallback = AddressOf BlueSoleil_OPP_Callback_Status

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_OPPRegisterStatusCallback4ThirdParty(ByVal connHandle As UInt32, ByVal functPtr_OPP_STATUS_INFO_CB As IntPtr)
    End Sub

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_OPPCancelTransfer(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_OPPPullObj(ByVal connHandle As UInt32, ByRef localFN As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_OPPPushObj(ByVal connHandle As UInt32, ByRef localFN As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_OPPExchangeObj(ByVal connHandle As UInt32, ByRef pushFN As Byte, ByRef pullFN As Byte, ByRef pushResult As UInt32, ByRef pullResult As UInt32) As UInt32
    End Function







    Private Sub BlueSoleil_OPP_Callback_Status(ByVal first As Byte, ByVal last As Byte, ByVal ptrFileNameBytes As IntPtr, ByVal theFileSize As UInt32, ByVal curSize As UInt32)

        'this callback is called by BlueSoleil during OPP transfers.

        Debug.Print("OPP Callback.  first = " & first & "   last = " & last & "   filesize = " & theFileSize & "    cursize = " & curSize)



        If last = 0 Then

            If first <> 0 Then
                'start of transfer.  
                BlueSoleil_OPP_Callback_Status_CurrSize = curSize
                BlueSoleil_OPP_Callback_Status_TotalSize = theFileSize

                Debug.Print("Start OPP xfer.  0 of " & theFileSize & " bytes.")

            Else
                BlueSoleil_OPP_Callback_Status_CurrSize = BlueSoleil_OPP_Callback_Status_CurrSize + curSize
                Debug.Print("Continue OPP xfer.  " & BlueSoleil_OPP_Callback_Status_CurrSize & " of " & theFileSize & " bytes.")
            End If


        Else

            'add final bytes.
            BlueSoleil_OPP_Callback_Status_CurrSize = BlueSoleil_OPP_Callback_Status_CurrSize + curSize
            Debug.Print("Finish OPP xfer.  " & BlueSoleil_OPP_Callback_Status_CurrSize & " of " & BlueSoleil_OPP_Callback_Status_TotalSize & " bytes.")

            'do whatever cuz we're done..

            'reset.
            ' BlueSoleil_OPP_Callback_Status_TotalSize = 0
            ' BlueSoleil_OPP_Callback_Status_CurrSize = 0

        End If

    End Sub


    Public Function BlueSoleil_OPP_PullVCard(ByVal connHandle As UInt32, ByVal folderToSaveTo_RemoteVCF As String, ByRef retFailDueToAccessDenied As Boolean) As Boolean

        'file is always saved with filename remote.vcf
        retFailDueToAccessDenied = False

        If connHandle = 0 Then
            Return False
        End If

        Dim pullPathBytes(0 To 0) As Byte
        pullPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(folderToSaveTo_RemoteVCF & Chr(0))

        Dim retUInt As UInt32 = Btsdk_OPPPullObj(connHandle, pullPathBytes(0))

        'check retSetPath value for access error.
        If retUInt = 1732 Then
            retFailDueToAccessDenied = True
        End If

        Return (retUInt = BTSDK_OK)

    End Function

    Public Function BlueSoleil_OPP_PushVCard(ByVal connHandle As UInt32, ByVal vcfFileNameToPush As String) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        Dim pushPathBytes(0 To 0) As Byte
        pushPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfFileNameToPush & Chr(0))

        Dim retUInt As UInt32 = Btsdk_OPPPushObj(connHandle, pushPathBytes(0))

        Return (retUInt = BTSDK_OK)

    End Function

    Public Function BlueSoleil_OPP_ExchangeVCards(ByVal connHandle As UInt32, ByVal vcfFileNameToPush As String, ByVal vcfPulledPathToWriteRemoteVCF As String, ByRef retPushBool As Boolean, ByRef retPullBool As Boolean) As Boolean

        retPushBool = False
        retPullBool = False

        If connHandle = 0 Then
            Return False
        End If

        Dim pushPathBytes(0 To 0) As Byte
        pushPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfFileNameToPush & Chr(0))

        Dim pullPathBytes(0 To 0) As Byte
        pullPathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(vcfPulledPathToWriteRemoteVCF & Chr(0))

        Dim retPush As UInt32 = 0, retPull As UInt32 = 0
        Dim retUInt As UInt32 = Btsdk_OPPExchangeObj(connHandle, pushPathBytes(0), pullPathBytes(0), retPush, retPull)

        retPushBool = (retPush = BTSDK_OK)
        retPullBool = (retPull = BTSDK_OK)

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_OPP_RegisterStatusCallback(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        delegateOPPstatusCallback = AddressOf BlueSoleil_OPP_Callback_Status

        Dim retUInt As UInt32 = BTSDK_OK        'i thought Btsdk_OPPRegisterStatusCallback4ThirdParty returned a value, but I guess not.

        Dim functPtr_StatusCallback As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateOPPstatusCallback)
        Btsdk_OPPRegisterStatusCallback4ThirdParty(connHandle, functPtr_StatusCallback)

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_OPP_UnregisterStatusCallback(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then
            Return False
        End If

        Dim retUInt As UInt32 = BTSDK_OK

        Dim functPtr_StatusCallback As IntPtr = IntPtr.Zero
        Btsdk_OPPRegisterStatusCallback4ThirdParty(connHandle, functPtr_StatusCallback)

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_OPP_CancelTransfer(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False


        Dim retUInt As UInt32
        retUInt = Btsdk_OPPCancelTransfer(connHandle)

        Return (retUInt = 0)

    End Function


End Module
