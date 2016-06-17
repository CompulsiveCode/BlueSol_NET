'modBlueSoleil_SPP - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices
Module modBlueSoleil_SPP

    '/* Parameters for Btsdk_PlugOutVComm and Btsdk_PlugInVComm */
    Private Const COMM_SET_USAGETYPE As UInt32 = &H1
    Private Const COMM_SET_RECORD As UInt32 = &H10



    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetAvailableExtSPPCOMPort(ByVal isLocal As Boolean) As Byte
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetClientPort(ByVal connHandle As UInt32) As UInt16
    End Function


    'I think the following API's are for creating COM ports, etc.
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_DeinitCommObj(ByVal com_idx As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_InitCommObj(ByVal com_idx As Byte, ByVal svcClass As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetASerialNum() As UInt32
    End Function



    Public Function BlueSoleil_SPP_GetCOMMportNum(ByVal connHandle As UInt32) As Integer

        If connHandle = 0 Then Return 0

        Dim retInt As Integer = 0

        retInt = Btsdk_GetClientPort(connHandle)

        Return retInt

    End Function


    Public Function BlueSoleil_SPP_GetAvailableExtPort(ByVal isLocal As Boolean) As Integer

        'not sure what this does, but I figured I would write a wrapper for it since it's simple.

        Dim retInt As Integer = 0

        retInt = Btsdk_GetAvailableExtSPPCOMPort(isLocal)

        Return retInt

    End Function

End Module
