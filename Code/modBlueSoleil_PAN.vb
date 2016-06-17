'modBlueSoleil_PAN - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices
Module modBlueSoleil_PAN


    Private BlueSoleil_PAN_IPaddress As String = ""


    Public Event BlueSoleil_Event_PAN_IPchanged(ByVal ipAddress As String)
    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncPANevent(ByVal inpEvent As UInt16, ByVal inpLen As UInt16, ByVal ptrParam As IntPtr)
    Public delegatePANevent As delfuncPANevent = AddressOf BlueSoleil_PAN_Callback_PANevent


    'use this if you want to register a callback to receive info, such as IP address.  That's the only info it provides.
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_PAN_RegIndCbk4ThirdParty(ByVal functPtr As UInt32)
    End Sub


    Private Sub BlueSoleil_PAN_Callback_PANevent(ByVal inpEvent As UInt16, ByVal inpLen As UInt16, ByVal ptrParam As IntPtr)

        Debug.Print("BlueSoleil_PAN_Callback_PANevent  length = " & inpLen)

        If inpLen < 4 Then Exit Sub

        'get the IP address.  'expecting LEN to be 4.
        Dim paramBytes(0 To inpLen - 1) As Byte
        Marshal.Copy(ptrParam, paramBytes, 0, paramBytes.Length)

        Dim IPstr As String = ""

        Dim i As Integer
        For i = 0 To paramBytes.Length - 1
            If i < paramBytes.Length - 1 Then
                IPstr = IPstr & CStr(paramBytes(i)) & "."
            Else
                IPstr = IPstr & CStr(paramBytes(i))
            End If
        Next i

        Debug.Print("BlueSoleil_PAN_Callback_PANevent  ip = " & IPstr)

        BlueSoleil_PAN_IPaddress = IPstr

        'RaiseEvent BlueSoleil_Event_PAN_IPchanged(IPstr)
        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_PAN_IPchanged(IPstr))
        t.Start()


    End Sub

    Public Sub BlueSoleil_PAN_RegisterCallbackForIPaddress()

        'this isn't required.  but if enabled, the variable BlueSoleil_PAN_IPaddress will get populated.

        Dim funcPtr As IntPtr = Marshal.GetFunctionPointerForDelegate(delegatePANevent)
        Dim funcPtrINT As UInt32 = CUInt(funcPtr)

        Btsdk_PAN_RegIndCbk4ThirdParty(funcPtrINT)

    End Sub


    Public Sub BlueSoleil_PAN_UnregisterCallbackForIPaddress()

        Dim funcPtr As IntPtr = IntPtr.Zero
        Dim funcPtrINT As UInt32 = CUInt(funcPtr)

        Btsdk_PAN_RegIndCbk4ThirdParty(funcPtrINT)

    End Sub

End Module
