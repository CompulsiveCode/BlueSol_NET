'modBlueSoleil_A2DP - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices
Module modBlueSoleil_A2DP


    '/*A2DP*/
    Private Const BTSDK_A2DP_AUDIOCARD_NAME_LEN As UInt32 = &H80

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RegisterA2DPSNKService(ByVal len As UInt16, ByRef ptrStrAudioCard As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_UnregisterA2DPSNKService() As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RegisterA2DPSRCService() As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_UnregisterA2DPSRCService() As UInt32
    End Function





    Public Function BlueSoleil_A2DP_RegisterSNKservice(ByVal audioCardName As String) As UInt32

        audioCardName = Trim(audioCardName)

        Dim cardnameLen As UInt16 = CUShort(Len(audioCardName) + 1)      '+1 to cover the null I guess.

        Dim cardnameBytes(0 To 0) As Byte
        cardnameBytes = System.Text.Encoding.UTF8.GetBytes(audioCardName & Chr(0))

        Dim retUInt As UInt32 = Btsdk_RegisterA2DPSNKService(cardnameLen, cardnameBytes(0))

        Return retUInt

    End Function


    Public Sub BlueSoleil_A2DP_UnregisterSNKservice()

        Dim retUInt As UInt32 = Btsdk_UnregisterA2DPSNKService()

    End Sub



    Public Function BlueSoleil_A2DP_RegisterSRCservice() As UInt32

        Dim retUInt As UInt32 = Btsdk_RegisterA2DPSRCService()

        Return retUInt

    End Function


    Public Sub BlueSoleil_A2DP_UnregisterSRCservice()

        Dim retUInt As UInt32 = Btsdk_UnregisterA2DPSRCService()

    End Sub

End Module
