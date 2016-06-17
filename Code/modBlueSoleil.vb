'modBlueSoleil - Written by Jesse Yeager.   www.CompulsiveCode.com
'


Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

Module modBlueSoleil



    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub btCallback_StatusInfo(ByVal msgType As UInt32, ByVal pulData As UInt32, ByVal funcParam As UInt32, ByVal funcArg As IntPtr)
    Public delegateBTstatusInfo As btCallback_StatusInfo = AddressOf BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo
    ' Public Event BlueSoleil_Event_ReceiveBluetoothStatusInfo(ByVal msgType As UInt32, ByVal pulData As UInt32, ByVal funcParam As UInt32)

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub btCallback_ConnEvent(ByVal connHandle As UInt32, ByVal evtType As UInt16, ByVal ptrArgs As IntPtr)
    Public delegateBTconnEvent As btCallback_ConnEvent = AddressOf BlueSoleil_Status_Callback_ConnEvent
    ' Public Event BlueSoleil_Event_ReceiveBluetoothStatusInfo(ByVal msgType As UInt32, ByVal pulData As UInt32, ByVal funcParam As UInt32)

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub btCallback_DeviceFound(ByVal dvcHandle As UInt32)
    Public delegateBTdeviceFound As btCallback_DeviceFound = AddressOf BlueSoleil_Status_Callback_DeviceFound


    Private BlueSoleil_Inquiry_IsComplete As Boolean = False
    Private BlueSoleil_Inquiry_DvcHandles(0 To 0) As UInt32
    Private BlueSoleil_Inquiry_DvcCount As Integer = 0


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub btCallback_InquiryCompleteInd()
    Public delegateBTinquiryCompleteInd As btCallback_InquiryCompleteInd = AddressOf BlueSoleil_Status_Callback_InquiryComplete
    ' Public Event BlueSoleil_Event_ReceiveBluetoothStatusInfo(ByVal msgType As UInt32, ByVal pulData As UInt32, ByVal funcParam As UInt32)

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub btCallback_InquiryResultInd(ByVal dvcHandle As UInt32)
    Public delegateBTinquiryResultInd As btCallback_InquiryResultInd = AddressOf BlueSoleil_Status_Callback_InquiryResult



    Public Event BlueSoleil_Status_TurnOn()
    Public Event BlueSoleil_Status_TurnOff()
    Public Event BlueSoleil_Status_Plugged()
    Public Event BlueSoleil_Status_Unplugged()
    Public Event BlueSoleil_Status_DevicePaired()
    Public Event BlueSoleil_Status_DeviceUnpaired()
    Public Event BlueSoleil_Status_DeviceDeleted()

    Public Event BlueSoleil_Status_ServiceConnectedInbound(ByVal dvcHandle As UInt32, ByVal propSvcHandle As UInt32, ByVal propSvcClass As UInt16)
    Public Event BlueSoleil_Status_ServiceConnectedOutbound(ByVal dvcHandle As UInt32, ByVal propSvcHandle As UInt32, ByVal propSvcClass As UInt16)
    Public Event BlueSoleil_Status_ServiceDisconnectedInbound(ByVal dvcHandle As UInt32, ByVal propSvcHandle As UInt32, ByVal propSvcClass As UInt16)
    Public Event BlueSoleil_Status_ServiceDisconnectedOutbound(ByVal dvcHandle As UInt32, ByVal propSvcHandle As UInt32, ByVal propSvcClass As UInt16)

    Public Event BlueSoleil_Status_DeviceFound(ByVal dvcHandle As UInt32)




    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_IsSDKInitialized() As Byte        'returns true/false
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_IsServerConnected() As Byte ' UInt32       'returns true/false
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RegisterCallback4ThirdParty(ByRef struCallback As Byte) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RegisterGetStatusInfoCB4ThirdParty(ByVal ptrFunc_ReceiveBluetoothStatusInfo As UInt32) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_SetStatusInfoFlag(ByVal msgTypes As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MallocMemory(ByVal numBytesToAllocate As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_FreeMemory(ByVal ptrMemBlock As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_IsBluetoothReady() As Byte            'returns true/false
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_IsBluetoothHardwareExisted() As Byte 'returns true/false
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_Init()
    End Sub

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_Done()
    End Sub

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StartBluetooth() As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StopBluetooth() As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetDiscoveryMode(ByRef retBTdiscMode As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_SetDiscoveryMode(ByVal BTdiscMode As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetLocalDeviceAddress(ByRef array6bytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetLocalName(ByRef array256bytes As Byte, ByRef arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_SetLocalDeviceClass(ByVal dvcClass As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetLocalDeviceClass(ByRef dvcClass As UInt32) As UInt32  'not sure about dvcClass...
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetLocalLMPInfo(ByRef struLocalLMPInfo As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_SetFixedPincode(ByRef arrayPINcodeBytes As Byte, ByVal arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetFixedPincode(ByRef arrayPINcodeBytes As Byte, ByRef arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_VendorCommand(ByVal evFlag As UInt32, ByRef struVendorCmd As Byte, ByRef struEventParam As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_EnumAVDriver() As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_DeEnumAVDriver()
    End Sub

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_ActivateEx(ByRef arraySerialNoBytes As Byte, ByVal arrayLen As Int16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StartDeviceDiscovery(ByVal dvcClass As UInt32, ByVal maxDevices As UInt16, ByVal maxDurations As UInt16) As UInt32  'default durations is 10.
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StopDeviceDiscovery() As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_UpdateRemoteDeviceName(ByVal dvcHandle As UInt32, ByRef arrayNameBytes As Byte, ByVal arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_CancelUpdateRemoteDeviceName(ByVal dvcHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_IsDevicePaired(ByVal dvcHandle As UInt32, ByRef retIsPaired As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PairDevice(ByVal dvcHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_UnPairDevice(ByVal dvcHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RegisterCallbackEx(ByRef struCallBack As Byte, ByVal pairingPriority As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_PinCodeReply(ByVal dvcHandle As UInt32, ByRef arrayPINBytes As Byte, ByVal arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AuthorizationResponse(ByVal svcHandle As UInt32, ByVal dvcHandle As UInt32, ByVal authResponse As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_IsDeviceConnected(ByVal dvcHandle As UInt32) As Byte
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteDeviceRole(ByVal dvcHandle As UInt32, ByRef retDvcRole As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteLMPInfo(ByVal dvcHandle As UInt32, ByRef struRemoteLMPInfo As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteRSSI(ByVal dvcHandle As UInt32, ByRef retRSSIval As SByte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteLinkQuality(ByVal dvcHandle As UInt32, ByRef linkQual As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetSupervisionTimeout(ByVal dvcHandle As UInt32, ByRef retTimeoutSlots As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_SetSupervisionTimeout(ByVal dvcHandle As UInt32, ByVal timeoutSlots As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_ChangeConnectionPacketType(ByVal dvcHandle As UInt32, ByVal pktTypeFlags As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteDeviceHandle(ByRef arrayDvcAdrsBytes As Byte) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AddRemoteDevice(ByRef arrayDvcAdrsBytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_DeleteRemoteDeviceByHandle(ByVal dvcHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_DeleteUnpairedDevicesByClass(ByVal dvcClass As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetStoredDevicesByClass(ByVal dvcClass As UInt32, ByRef arrayDvcHandles As UInt32, ByVal arrayMaxNumEntries As UInt32) As Integer
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl, EntryPoint:="Btsdk_GetStoredDevicesByClass")>
    Private Function Btsdk_GetStoredDevicesByClass_ByVal(ByVal dvcClass As UInt32, ByVal nullInt1 As UInt32, ByVal nullInt2 As UInt32) As Integer
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetPairedDevices(ByRef arrayDvcHandles As UInt32, ByVal arrayMaxNumEntries As UInt32) As Integer
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl, EntryPoint:="Btsdk_GetPairedDevices")>
    Private Function Btsdk_GetPairedDevices_ByValArray(ByVal arrayDvcHandles As UInt32, ByVal arrayMaxNumEntries As UInt32) As Integer
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetInquiredDevices(ByRef arrayDvcHandles As UInt32, ByVal arrayMaxNumEntries As UInt32) As Integer
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl, EntryPoint:="Btsdk_GetInquiredDevices")>
    Private Function Btsdk_GetInquiredDevices_ByValArray(ByVal arrayDvcHandles As UInt32, ByVal arrayMaxNumEntries As UInt32) As Integer
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StartEnumRemoteDevice(ByVal flag As UInt32, ByVal dvcClass As UInt32) As UInt32         'returns enum handle.
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_EnumRemoteDevice(ByVal enumHandle As UInt32, ByRef struRemoteDvcProp As Byte) As UInt32     'returns device handle.
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_EndEnumRemoteDevice(ByVal enumHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteDeviceAddress(ByVal dvcHandle As UInt32, ByRef array6bytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteDeviceName(ByVal dvcHandle As UInt32, ByRef arrayBytes As Byte, ByRef arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteDeviceClass(ByVal dvcHandle As UInt32, ByRef retDvcClass As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteDeviceProperty(ByVal dvcHandle As UInt32, ByRef struRemoteDeviceProp As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RemoteDeviceFlowStatistic(ByVal dvcHandle As UInt32, ByRef rxNumBytes As UInt32, ByRef txNumBytes As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_BrowseRemoteServicesEx(ByVal dvcHandle As UInt32, ByRef struARRAY_SDPSearchPatternBytes As Byte, ByRef arrayNumEntries As UInt32, ByRef retArraySvcHandles As UInt32, ByRef arrayMaxNumEntries As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_BrowseRemoteServices(ByVal dvcHandle As UInt32, ByRef retArraySvcHandles As UInt32, ByRef arrayMaxNumEntries As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl, EntryPoint:="Btsdk_BrowseRemoteServices")>
    Private Function Btsdk_BrowseRemoteServices_ByValArray(ByVal dvcHandle As UInt32, ByVal nullInt1 As UInt32, ByRef arrayMaxNumEntries As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RefreshRemoteServiceAttributes(ByVal dvcHandle As UInt32, ByRef struRemoteServiceAttrBytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteServicesEx(ByVal dvcHandle As UInt32, ByRef struARRAY_SDPSearchPatternBytes As Byte, ByRef arrayNumEntries As UInt32, ByRef retArraySvcHandles As UInt32, ByRef arrayMaxNumEntries As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteServices(ByVal dvcHandle As UInt32, ByRef retArraySvcHandles As UInt32, ByRef arrayMaxNumEntries As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl, EntryPoint:="Btsdk_GetRemoteServices")>
    Private Function Btsdk_GetRemoteServices_ByValArray(ByVal dvcHandle As UInt32, ByVal retArraySvcHandles As UInt32, ByRef arrayMaxNumEntries As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteServiceAttributes(ByVal svcHandle As UInt32, ByRef struRemoteServiceAttrBytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StartEnumRemoteService() As UInt32      'returns enum handle
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_EnumRemoteService(ByVal enumHandle As UInt32, ByRef struRemoteServiceAttrBytes As Byte) As UInt32  'returns service handle
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_EndEnumRemoteService(ByVal enumHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_SetRemoteServiceParam(ByVal svcHandle As UInt32, ByVal appParm As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetRemoteServiceParam(ByVal svcHandle As UInt32, ByRef retAppParm As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Connect(ByVal svcHandle As UInt32, ByVal lParm As UInt32, ByRef retConnHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_ConnectEx(ByVal dvcHandle As UInt32, ByVal svcClass As UInt32, ByVal lParm As UInt32, ByRef retConnHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetConnectionProperty(ByVal connHandle As UInt32, ByRef struConnProp As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Disconnect(ByVal connHandle As UInt32) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StartServer(ByVal svcHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StopServer(ByVal svcHandle As UInt32) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_StartEnumLocalServer() As UInt32      'returns enum handle
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_EnumLocalServer(ByVal enumHandle As UInt32, ByRef struLocalServiceAttrBytes As Byte) As UInt32  'returns service handle
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_EndEnumLocalServer(ByVal enumHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetServerAttributes(ByVal svcHandle As UInt32, ByRef struLocalServerAttrBytes As Byte) As UInt32
    End Function



    Private Const BTSDK_TRUE As UInt32 = 1
    Private Const BTSDK_FALSE As UInt32 = 0



    '/* Max size value used in service attribute structures */
    Private Const BTSDK_SERVICENAME_MAXLENGTH As UInt16 = 80
    Private Const BTSDK_MAX_SUPPORT_FORMAT As UInt16 = 6       '/* OPP format number */
    Private Const BTSDK_PATH_MAXLENGTH As UInt16 = 256     '/* Shall not be larger than FTP_MAX_PATH and OPP_MAX_PATH */
    Private Const BTSDK_CARDNAME_MAXLENGTH As UInt16 = 256     '/* Shall not be larger than OPP_MAX_NAME */
    Private Const BTSDK_PACKETTYPE_MAXNUM As UInt16 = 10      '/* PAN supported network packet type */

    '/* Max size value used in device attribute structures */
    Private Const BTSDK_DEVNAME_LEN As UInt16 = 64      '/* Shall not be larger than MAX_NAME_LEN */
    Private Const BTSDK_SHORTCUT_NAME_LEN As UInt16 = 100
    Private Const BTSDK_BDADDR_LEN As UInt16 = 6
    Private Const BTSDK_LINKKEY_LEN As UInt16 = 16
    Private Const BTSDK_PINCODE_LEN As UInt16 = 16


    '/* Invalid handle value for all handle type */
    Private Const BTSDK_INVALID_HANDLE As UInt32 = &H0

    '/* Error Code List */
    Private Const BTSDK_OK As UInt32 = &H0

    '/* SDP error */
    Private Const BTSDK_ER_SDP_INDEX As UInt32 = &HC0
    Private Const BTSDK_ER_SERVER_IS_ACTIVE As UInt32 = (BTSDK_ER_SDP_INDEX + &H0)
    Private Const BTSDK_ER_NO_SERVICE As UInt32 = (BTSDK_ER_SDP_INDEX + &H1)
    Private Const BTSDK_ER_SERVICE_RECORD_NOT_EXIST As UInt32 = (BTSDK_ER_SDP_INDEX + &H2)

    '/* General Error */
    Private Const BTSDK_ER_GENERAL_INDEX As UInt32 = &H300
    Private Const BTSDK_ER_HANDLE_NOT_EXIST As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H1)
    Private Const BTSDK_ER_OPERATION_FAILURE As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H2)
    Private Const BTSDK_ER_SDK_UNINIT As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H3)
    Private Const BTSDK_ER_INVALID_PARAMETER As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H4)
    Private Const BTSDK_ER_NULL_POINTER As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H5)
    Private Const BTSDK_ER_NO_MEMORY As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H6)
    Private Const BTSDK_ER_BUFFER_NOT_ENOUGH As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H7)
    Private Const BTSDK_ER_FUNCTION_NOTSUPPORT As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H8)
    Private Const BTSDK_ER_NO_FIXED_PIN_CODE As UInt32 = (BTSDK_ER_GENERAL_INDEX + &H9)
    Private Const BTSDK_ER_CONNECTION_EXIST As UInt32 = (BTSDK_ER_GENERAL_INDEX + &HA)
    Private Const BTSDK_ER_OPERATION_CONFLICT As UInt32 = (BTSDK_ER_GENERAL_INDEX + &HB)
    Private Const BTSDK_ER_NO_MORE_CONNECTION_ALLOWED As UInt32 = (BTSDK_ER_GENERAL_INDEX + &HC)
    Private Const BTSDK_ER_ITEM_EXIST As UInt32 = (BTSDK_ER_GENERAL_INDEX + &HD)
    Private Const BTSDK_ER_ITEM_INUSE As UInt32 = (BTSDK_ER_GENERAL_INDEX + &HE)
    Private Const BTSDK_ER_DEVICE_UNPAIRED As UInt32 = (BTSDK_ER_GENERAL_INDEX + &HF)

    '/* HCI Error */
    Private Const BTSDK_ER_HCI_INDEX As UInt32 = &H400
    Private Const BTSDK_ER_UNKNOWN_HCI_COMMAND As UInt32 = (BTSDK_ER_HCI_INDEX + &H1)
    Private Const BTSDK_ER_NO_CONNECTION As UInt32 = (BTSDK_ER_HCI_INDEX + &H2)
    Private Const BTSDK_ER_HARDWARE_FAILURE As UInt32 = (BTSDK_ER_HCI_INDEX + &H3)
    Private Const BTSDK_ER_PAGE_TIMEOUT As UInt32 = (BTSDK_ER_HCI_INDEX + &H4)
    Private Const BTSDK_ER_AUTHENTICATION_FAILURE As UInt32 = (BTSDK_ER_HCI_INDEX + &H5)
    Private Const BTSDK_ER_KEY_MISSING As UInt32 = (BTSDK_ER_HCI_INDEX + &H6)
    Private Const BTSDK_ER_MEMORY_FULL As UInt32 = (BTSDK_ER_HCI_INDEX + &H7)
    Private Const BTSDK_ER_CONNECTION_TIMEOUT As UInt32 = (BTSDK_ER_HCI_INDEX + &H8)
    Private Const BTSDK_ER_MAX_NUMBER_OF_CONNECTIONS As UInt32 = (BTSDK_ER_HCI_INDEX + &H9)
    Private Const BTSDK_ER_MAX_NUMBER_OF_SCO_CONNECTIONS As UInt32 = (BTSDK_ER_HCI_INDEX + &HA)
    Private Const BTSDK_ER_ACL_CONNECTION_ALREADY_EXISTS As UInt32 = (BTSDK_ER_HCI_INDEX + &HB)
    Private Const BTSDK_ER_COMMAND_DISALLOWED As UInt32 = (BTSDK_ER_HCI_INDEX + &HC)
    Private Const BTSDK_ER_HOST_REJECTED_LIMITED_RESOURCES As UInt32 = (BTSDK_ER_HCI_INDEX + &HD)
    Private Const BTSDK_ER_HOST_REJECTED_SECURITY_REASONS As UInt32 = (BTSDK_ER_HCI_INDEX + &HE)
    Private Const BTSDK_ER_HOST_REJECTED_PERSONAL_DEVICE As UInt32 = (BTSDK_ER_HCI_INDEX + &HF)
    Private Const BTSDK_ER_HOST_TIMEOUT As UInt32 = (BTSDK_ER_HCI_INDEX + &H10)
    Private Const BTSDK_ER_UNSUPPORTED_FEATURE As UInt32 = (BTSDK_ER_HCI_INDEX + &H11)
    Private Const BTSDK_ER_INVALID_HCI_COMMAND_PARAMETERS As UInt32 = (BTSDK_ER_HCI_INDEX + &H12)
    Private Const BTSDK_ER_PEER_DISCONNECTION_USER_END As UInt32 = (BTSDK_ER_HCI_INDEX + &H13)
    Private Const BTSDK_ER_PEER_DISCONNECTION_LOW_RESOURCES As UInt32 = (BTSDK_ER_HCI_INDEX + &H14)
    Private Const BTSDK_ER_PEER_DISCONNECTION_TO_POWER_OFF As UInt32 = (BTSDK_ER_HCI_INDEX + &H15)
    Private Const BTSDK_ER_LOCAL_DISCONNECTION As UInt32 = (BTSDK_ER_HCI_INDEX + &H16)
    Private Const BTSDK_ER_REPEATED_ATTEMPTS As UInt32 = (BTSDK_ER_HCI_INDEX + &H17)
    Private Const BTSDK_ER_PAIRING_NOT_ALLOWED As UInt32 = (BTSDK_ER_HCI_INDEX + &H18)
    Private Const BTSDK_ER_UNKNOWN_LMP_PDU As UInt32 = (BTSDK_ER_HCI_INDEX + &H19)
    Private Const BTSDK_ER_UNSUPPORTED_REMOTE_FEATURE As UInt32 = (BTSDK_ER_HCI_INDEX + &H1A)
    Private Const BTSDK_ER_SCO_OFFSET_REJECTED As UInt32 = (BTSDK_ER_HCI_INDEX + &H1B)
    Private Const BTSDK_ER_SCO_INTERVAL_REJECTED As UInt32 = (BTSDK_ER_HCI_INDEX + &H1C)
    Private Const BTSDK_ER_SCO_AIR_MODE_REJECTED As UInt32 = (BTSDK_ER_HCI_INDEX + &H1D)
    Private Const BTSDK_ER_INVALID_LMP_PARAMETERS As UInt32 = (BTSDK_ER_HCI_INDEX + &H1E)
    Private Const BTSDK_ER_UNSPECIFIED_ERROR As UInt32 = (BTSDK_ER_HCI_INDEX + &H1F)
    Private Const BTSDK_ER_UNSUPPORTED_LMP_PARAMETER_VALUE As UInt32 = (BTSDK_ER_HCI_INDEX + &H20)
    Private Const BTSDK_ER_ROLE_CHANGE_NOT_ALLOWED As UInt32 = (BTSDK_ER_HCI_INDEX + &H21)
    Private Const BTSDK_ER_LMP_RESPONSE_TIMEOUT As UInt32 = (BTSDK_ER_HCI_INDEX + &H22)
    Private Const BTSDK_ER_LMP_ERROR_TRANSACTION_COLLISION As UInt32 = (BTSDK_ER_HCI_INDEX + &H23)
    Private Const BTSDK_ER_LMP_PDU_NOT_ALLOWED As UInt32 = (BTSDK_ER_HCI_INDEX + &H24)
    Private Const BTSDK_ER_ENCRYPTION_MODE_NOT_ACCEPTABLE As UInt32 = (BTSDK_ER_HCI_INDEX + &H25)
    Private Const BTSDK_ER_UNIT_KEY_USED As UInt32 = (BTSDK_ER_HCI_INDEX + &H26)
    Private Const BTSDK_ER_QOS_IS_NOT_SUPPORTED As UInt32 = (BTSDK_ER_HCI_INDEX + &H27)
    Private Const BTSDK_ER_INSTANT_PASSED As UInt32 = (BTSDK_ER_HCI_INDEX + &H28)
    Private Const BTSDK_ER_PAIRING_WITH_UNIT_KEY_NOT_SUPPORTED As UInt32 = (BTSDK_ER_HCI_INDEX + &H29)
    Private Const BTSDK_ER_DIFFERENT_TRANSACTION_COLLISION As UInt32 = (BTSDK_ER_HCI_INDEX + &H2A)
    Private Const BTSDK_ER_QOS_UNACCEPTABLE_PARAMETER As UInt32 = (BTSDK_ER_HCI_INDEX + &H2C)
    Private Const BTSDK_ER_QOS_REJECTED As UInt32 = (BTSDK_ER_HCI_INDEX + &H2D)
    Private Const BTSDK_ER_CHANNEL_CLASS_NOT_SUPPORTED As UInt32 = (BTSDK_ER_HCI_INDEX + &H2E)
    Private Const BTSDK_ER_INSUFFICIENT_SECURITY As UInt32 = (BTSDK_ER_HCI_INDEX + &H2F)
    Private Const BTSDK_ER_PARAMETER_OUT_OF_RANGE As UInt32 = (BTSDK_ER_HCI_INDEX + &H30)
    Private Const BTSDK_ER_ROLE_SWITCH_PENDING As UInt32 = (BTSDK_ER_HCI_INDEX + &H32)
    Private Const BTSDK_ER_RESERVED_SLOT_VIOLATION As UInt32 = (BTSDK_ER_HCI_INDEX + &H34)
    Private Const BTSDK_ER_ROLE_SWITCH_FAILED As UInt32 = (BTSDK_ER_HCI_INDEX + &H35)

    '/* OBEX error */
    Private Const BTSDK_ER_OBEX_INDEX As UInt32 = &H600
    Private Const BTSDK_ER_CONTINUE As UInt32 = (BTSDK_ER_OBEX_INDEX + &H90)
    Private Const BTSDK_ER_SUCCESS As UInt32 = (BTSDK_ER_OBEX_INDEX + &HA0)
    Private Const BTSDK_ER_CREATED As UInt32 = (BTSDK_ER_OBEX_INDEX + &HA1)
    Private Const BTSDK_ER_ACCEPTED As UInt32 = (BTSDK_ER_OBEX_INDEX + &HA2)
    Private Const BTSDK_ER_NON_AUTH_INFO As UInt32 = (BTSDK_ER_OBEX_INDEX + &HA3)
    Private Const BTSDK_ER_NO_CONTENT As UInt32 = (BTSDK_ER_OBEX_INDEX + &HA4)
    Private Const BTSDK_ER_RESET_CONTENT As UInt32 = (BTSDK_ER_OBEX_INDEX + &HA5)
    Private Const BTSDK_ER_PARTIAL_CONTENT As UInt32 = (BTSDK_ER_OBEX_INDEX + &HA6)
    Private Const BTSDK_ER_MULT_CHOICES As UInt32 = (BTSDK_ER_OBEX_INDEX + &HB0)
    Private Const BTSDK_ER_MOVE_PERM As UInt32 = (BTSDK_ER_OBEX_INDEX + &HB1)
    Private Const BTSDK_ER_MOVE_TEMP As UInt32 = (BTSDK_ER_OBEX_INDEX + &HB2)
    Private Const BTSDK_ER_SEE_OTHER As UInt32 = (BTSDK_ER_OBEX_INDEX + &HB3)
    Private Const BTSDK_ER_NOT_MODIFIED As UInt32 = (BTSDK_ER_OBEX_INDEX + &HB4)
    Private Const BTSDK_ER_USE_PROXY As UInt32 = (BTSDK_ER_OBEX_INDEX + &HB5)
    Private Const BTSDK_ER_BAD_REQUEST As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC0)
    Private Const BTSDK_ER_UNAUTHORIZED As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC1)
    Private Const BTSDK_ER_PAY_REQ As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC2)
    Private Const BTSDK_ER_FORBIDDEN As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC3)
    Private Const BTSDK_ER_NOTFOUND As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC4)
    Private Const BTSDK_ER_METHOD_NOT_ALLOWED As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC5)
    Private Const BTSDK_ER_NOT_ACCEPTABLE As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC6)
    Private Const BTSDK_ER_PROXY_AUTH_REQ As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC7)
    Private Const BTSDK_ER_REQUEST_TIMEOUT As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC8)
    Private Const BTSDK_ER_CONFLICT As UInt32 = (BTSDK_ER_OBEX_INDEX + &HC9)
    Private Const BTSDK_ER_GONE As UInt32 = (BTSDK_ER_OBEX_INDEX + &HCA)
    Private Const BTSDK_ER_LEN_REQ As UInt32 = (BTSDK_ER_OBEX_INDEX + &HCB)
    Private Const BTSDK_ER_PREC_FAIL As UInt32 = (BTSDK_ER_OBEX_INDEX + &HCC)
    Private Const BTSDK_ER_REQ_ENTITY_TOO_LARGE As UInt32 = (BTSDK_ER_OBEX_INDEX + &HCD)
    Private Const BTSDK_ER_URL_TOO_LARGE As UInt32 = (BTSDK_ER_OBEX_INDEX + &HCE)
    Private Const BTSDK_ER_UNSUPPORTED_MEDIA_TYPE As UInt32 = (BTSDK_ER_OBEX_INDEX + &HCF)
    Private Const BTSDK_ER_SVR_ERR As UInt32 = (BTSDK_ER_OBEX_INDEX + &HD0)
    Private Const BTSDK_ER_NOTIMPLEMENTED As UInt32 = (BTSDK_ER_OBEX_INDEX + &HD1)
    Private Const BTSDK_ER_BAD_GATEWAY As UInt32 = (BTSDK_ER_OBEX_INDEX + &HD2)
    Private Const BTSDK_ER_SERVICE_UNAVAILABLE As UInt32 = (BTSDK_ER_OBEX_INDEX + &HD3)
    Private Const BTSDK_ER_GATEWAY_TIMEOUT As UInt32 = (BTSDK_ER_OBEX_INDEX + &HD4)
    Private Const BTSDK_ER_HTTP_NOTSUPPORT As UInt32 = (BTSDK_ER_OBEX_INDEX + &HD5)
    Private Const BTSDK_ER_DATABASE_FULL As UInt32 = (BTSDK_ER_OBEX_INDEX + &HE0)
    Private Const BTSDK_ER_DATABASE_LOCK As UInt32 = (BTSDK_ER_OBEX_INDEX + &HE1)

    '/* Class of Device */
    '/*major service classes*/
    Private Const BTSDK_SRVCLS_LDM As UInt32 = &H2000     '/* Limited Discoverable Mode */
    Private Const BTSDK_SRVCLS_POSITION As UInt32 = &H10000
    Private Const BTSDK_SRVCLS_NETWORK As UInt32 = &H20000              '
    Private Const BTSDK_SRVCLS_RENDER As UInt32 = &H40000
    Private Const BTSDK_SRVCLS_CAPTURE As UInt32 = &H80000              '
    Private Const BTSDK_SRVCLS_OBJECT As UInt32 = &H100000              '
    Private Const BTSDK_SRVCLS_AUDIO As UInt32 = &H200000
    Private Const BTSDK_SRVCLS_TELEPHONE As UInt32 = &H400000           '
    Private Const BTSDK_SRVCLS_INFOR As UInt32 = &H800000
    'Private Const BTSDK_SRVCLS_MASK(a) As UInt32 = (((BTUINT32)(a) >> 13) And &H7FF)   '!!!

    '/*major device classes*/			                                    
    Private Const BTSDK_DEVCLS_MISC As UInt32 = &H0
    Private Const BTSDK_DEVCLS_COMPUTER As UInt32 = &H100
    Private Const BTSDK_DEVCLS_PHONE As UInt32 = &H200
    Private Const BTSDK_DEVCLS_LAP As UInt32 = &H300
    Private Const BTSDK_DEVCLS_AUDIO As UInt32 = &H400
    Private Const BTSDK_DEVCLS_PERIPHERAL As UInt32 = &H500
    Private Const BTSDK_DEVCLS_IMAGE As UInt32 = &H600
    Private Const BTSDK_DEVCLS_WEARABLE As UInt32 = &H700
    Private Const BTSDK_DEVCLS_UNCLASSIFIED As UInt32 = &H1F00
    'Private Const BTSDK_DEVCLS_MASK(a) As UInt32 = (((BTUINT32)(a) >> 8) And &H1F) '!!!
    'Private Const BTSDK_MINDEVCLS_MASK(a) As UInt32 = (((BTUINT32)(a) >> 2) And &H3F) '!!!

    '/*the minor device class field - computer major class */
    Private Const BTSDK_COMPCLS_UNCLASSIFIED As UInt32 = (BTSDK_DEVCLS_COMPUTER Or &H0)
    Private Const BTSDK_COMPCLS_DESKTOP As UInt32 = (BTSDK_DEVCLS_COMPUTER Or &H4)
    Private Const BTSDK_COMPCLS_SERVER As UInt32 = (BTSDK_DEVCLS_COMPUTER Or &H8)
    Private Const BTSDK_COMPCLS_LAPTOP As UInt32 = (BTSDK_DEVCLS_COMPUTER Or &HC)
    Private Const BTSDK_COMPCLS_HANDHELD As UInt32 = (BTSDK_DEVCLS_COMPUTER Or &H10)
    Private Const BTSDK_COMPCLS_PALMSIZED As UInt32 = (BTSDK_DEVCLS_COMPUTER Or &H14)
    Private Const BTSDK_COMPCLS_WEARABLE As UInt32 = (BTSDK_DEVCLS_COMPUTER Or &H18)

    '/*the minor device class field - phone major class*/
    Private Const BTSDK_PHONECLS_UNCLASSIFIED As UInt32 = (BTSDK_DEVCLS_PHONE Or &H0)
    Private Const BTSDK_PHONECLS_CELLULAR As UInt32 = (BTSDK_DEVCLS_PHONE Or &H4)
    Private Const BTSDK_PHONECLS_CORDLESS As UInt32 = (BTSDK_DEVCLS_PHONE Or &H8)
    Private Const BTSDK_PHONECLS_SMARTPHONE As UInt32 = (BTSDK_DEVCLS_PHONE Or &HC)
    Private Const BTSDK_PHONECLS_WIREDMODEM As UInt32 = (BTSDK_DEVCLS_PHONE Or &H10)
    Private Const BTSDK_PHONECLS_COMMONISDNACCESS As UInt32 = (BTSDK_DEVCLS_PHONE Or &H14)
    Private Const BTSDK_PHONECLS_SIMCARDREADER As UInt32 = (BTSDK_DEVCLS_PHONE Or &H18)

    '/*the minor device class field - LAN/Network access point major class*/
    Private Const BTSDK_LAP_FULLY As UInt32 = (BTSDK_DEVCLS_LAP Or &H0)
    Private Const BTSDK_LAP_17 As UInt32 = (BTSDK_DEVCLS_LAP Or &H20)
    Private Const BTSDK_LAP_33 As UInt32 = (BTSDK_DEVCLS_LAP Or &H40)
    Private Const BTSDK_LAP_50 As UInt32 = (BTSDK_DEVCLS_LAP Or &H60)
    Private Const BTSDK_LAP_67 As UInt32 = (BTSDK_DEVCLS_LAP Or &H80)
    Private Const BTSDK_LAP_83 As UInt32 = (BTSDK_DEVCLS_LAP Or &HA0)
    Private Const BTSDK_LAP_99 As UInt32 = (BTSDK_DEVCLS_LAP Or &HC0)
    Private Const BTSDK_LAP_NOSRV As UInt32 = (BTSDK_DEVCLS_LAP Or &HE0)

    '/*the minor device class field - Audio/Video major class*/
    Private Const BTSDK_AV_UNCLASSIFIED As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H0)
    Private Const BTSDK_AV_HEADSET As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H4)
    Private Const BTSDK_AV_HANDSFREE As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H8)
    Private Const BTSDK_AV_HEADANDHAND As UInt32 = (BTSDK_DEVCLS_AUDIO Or &HC)
    Private Const BTSDK_AV_MICROPHONE As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H10)
    Private Const BTSDK_AV_LOUDSPEAKER As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H14)
    Private Const BTSDK_AV_HEADPHONES As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H18)
    Private Const BTSDK_AV_PORTABLEAUDIO As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H1C)
    Private Const BTSDK_AV_CARAUDIO As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H20)
    Private Const BTSDK_AV_SETTOPBOX As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H24)
    Private Const BTSDK_AV_HIFIAUDIO As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H28)
    Private Const BTSDK_AV_VCR As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H2C)
    Private Const BTSDK_AV_VIDEOCAMERA As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H30)
    Private Const BTSDK_AV_CAMCORDER As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H34)
    Private Const BTSDK_AV_VIDEOMONITOR As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H38)
    Private Const BTSDK_AV_VIDEODISPANDLOUDSPK As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H3C)
    Private Const BTSDK_AV_VIDEOCONFERENCE As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H40)
    Private Const BTSDK_AV_GAMEORTOY As UInt32 = (BTSDK_DEVCLS_AUDIO Or &H48)

    '/*the minor device class field - peripheral major class*/
    Private Const BTSDK_PERIPHERAL_UNCLASSIFIED As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &H0)
    Private Const BTSDK_PERIPHERAL_JOYSTICK As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &H4)
    Private Const BTSDK_PERIPHERAL_GAMEPAD As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &H8)
    Private Const BTSDK_PERIPHERAL_REMCONTROL As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &HC)
    Private Const BTSDK_PERIPHERAL_SENSE As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &H10)
    Private Const BTSDK_PERIPHERAL_TABLET As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &H14)
    Private Const BTSDK_PERIPHERAL_SIMCARDREADER As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &H18)
    Private Const BTSDK_PERIPHERAL_KEYBOARD As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &H40)
    Private Const BTSDK_PERIPHERAL_POINT As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &H80)
    Private Const BTSDK_PERIPHERAL_KEYORPOINT As UInt32 = (BTSDK_DEVCLS_PERIPHERAL Or &HC0)

    '/*the minor device class field - imaging major class*/
    Private Const BTSDK_IMAGE_DISPLAY As UInt32 = (BTSDK_DEVCLS_IMAGE Or &H10)
    Private Const BTSDK_IMAGE_CAMERA As UInt32 = (BTSDK_DEVCLS_IMAGE Or &H20)
    Private Const BTSDK_IMAGE_SCANNER As UInt32 = (BTSDK_DEVCLS_IMAGE Or &H40)
    Private Const BTSDK_IMAGE_PRINTER As UInt32 = (BTSDK_DEVCLS_IMAGE Or &H80)

    '/*the minor device class field - wearable major class*/
    Private Const BTSDK_WERABLE_WATCH As UInt32 = (BTSDK_DEVCLS_WEARABLE Or &H4)
    Private Const BTSDK_WERABLE_PAGER As UInt32 = (BTSDK_DEVCLS_WEARABLE Or &H8)
    Private Const BTSDK_WERABLE_JACKET As UInt32 = (BTSDK_DEVCLS_WEARABLE Or &HC)
    Private Const BTSDK_WERABLE_HELMET As UInt32 = (BTSDK_DEVCLS_WEARABLE Or &H10)
    Private Const BTSDK_WERABLE_GLASSES As UInt32 = (BTSDK_DEVCLS_WEARABLE Or &H14)

    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!yay
    '/* Class of Service */
    Private Const BTSDK_CLS_SERIAL_PORT As UInt32 = &H1101
    Private Const BTSDK_CLS_LAN_ACCESS As UInt32 = &H1102
    Private Const BTSDK_CLS_DIALUP_NET As UInt32 = &H1103
    Private Const BTSDK_CLS_IRMC_SYNC As UInt32 = &H1104
    Private Const BTSDK_CLS_OBEX_OBJ_PUSH As UInt32 = &H1105
    Private Const BTSDK_CLS_OBEX_FILE_TRANS As UInt32 = &H1106
    Private Const BTSDK_CLS_IRMC_SYNC_CMD As UInt32 = &H1107
    Private Const BTSDK_CLS_HEADSET As UInt32 = &H1108
    Private Const BTSDK_CLS_CORDLESS_TELE As UInt32 = &H1109
    Private Const BTSDK_CLS_AUDIO_SOURCE As UInt32 = &H110A     '
    Private Const BTSDK_CLS_AUDIO_SINK As UInt32 = &H110B
    Private Const BTSDK_CLS_AVRCP_TG As UInt32 = &H110C
    Private Const BTSDK_CLS_ADV_AUDIO_DISTRIB As UInt32 = &H110D
    Private Const BTSDK_CLS_AVRCP_CT As UInt32 = &H110E
    Private Const BTSDK_CLS_VIDEO_CONFERENCE As UInt32 = &H110F
    Private Const BTSDK_CLS_INTERCOM As UInt32 = &H1110
    Private Const BTSDK_CLS_FAX As UInt32 = &H1111
    Private Const BTSDK_CLS_HEADSET_AG As UInt32 = &H1112
    Private Const BTSDK_CLS_WAP As UInt32 = &H1113
    Private Const BTSDK_CLS_WAP_CLIENT As UInt32 = &H1114
    Private Const BTSDK_CLS_PAN_PANU As UInt32 = &H1115
    Private Const BTSDK_CLS_PAN_NAP As UInt32 = &H1116
    Private Const BTSDK_CLS_PAN_GN As UInt32 = &H1117
    Private Const BTSDK_CLS_DIRECT_PRINT As UInt32 = &H1118
    Private Const BTSDK_CLS_REF_PRINT As UInt32 = &H1119
    Private Const BTSDK_CLS_IMAGING As UInt32 = &H111A
    Private Const BTSDK_CLS_IMAG_RESPONDER As UInt32 = &H111B
    Private Const BTSDK_CLS_IMAG_AUTO_ARCH As UInt32 = &H111C
    Private Const BTSDK_CLS_IMAG_REF_OBJ As UInt32 = &H111D
    Private Const BTSDK_CLS_HANDSFREE As UInt32 = &H111E
    Private Const BTSDK_CLS_HANDSFREE_AG As UInt32 = &H111F
    Private Const BTSDK_CLS_DPS_REF_OBJ As UInt32 = &H1120
    Private Const BTSDK_CLS_REFLECTED_UI As UInt32 = &H1121
    Private Const BTSDK_CLS_BASIC_PRINT As UInt32 = &H1122
    Private Const BTSDK_CLS_PRINT_STATUS As UInt32 = &H1123
    Private Const BTSDK_CLS_HID As UInt32 = &H1124
    Private Const BTSDK_CLS_HCRP As UInt32 = &H1125
    Private Const BTSDK_CLS_HCR_PRINT As UInt32 = &H1126
    Private Const BTSDK_CLS_HCR_SCAN As UInt32 = &H1127
    Private Const BTSDK_CLS_SIM_ACCESS As UInt32 = &H112D
    Private Const BTSDK_CLS_PBAP_PCE As UInt32 = &H112E
    Private Const BTSDK_CLS_PBAP_PSE As UInt32 = &H112F
    Private Const BTSDK_CLS_PHONEBOOK_ACCESS As UInt32 = &H1130
    Private Const BTSDK_CLS_PNP_INFO As UInt32 = &H1200

    Private Const BTSDK_CLS_OBEX_MESSAGEACCESSSERVER As UInt32 = &H1132
    Private Const BTSDK_CLS_OBEX_MESSAGENOTIFICATIONSERVER As UInt32 = &H1133
    Private Const BTSDK_CLS_OBEX_MESSAGEACCESSPROFILE As UInt32 = &H1134


    '/* Type of Connection Event */
    Private Const BTSDK_APP_EV_CONN_IND As UInt32 = &H1
    Private Const BTSDK_APP_EV_DISC_IND As UInt32 = &H2
    Private Const BTSDK_APP_EV_CONN_CFM As UInt32 = &H7
    Private Const BTSDK_APP_EV_DISC_CFM As UInt32 = &H8

    '/* Definitions for Compatibility */
    Private Const BTSDK_APP_EV_CONN As UInt32 = &H1
    Private Const BTSDK_APP_EV_DISC As UInt32 = &H2

    '//Call back user priority
    Private Const BTSDK_CLIENTCBK_PRIORITY_HIGH As UInt32 = 3
    Private Const BTSDK_CLIENTCBK_PRIORITY_MEDIUM As UInt32 = 2

    '//Whether user handle pin code and authorization callback
    Private Const BTSDK_CLIENTCBK_HANDLED As UInt32 = 1
    Private Const BTSDK_CLIENTCBK_NOTHANDLED As UInt32 = 0

    '/* Authorization Result */
    Private Const BTSDK_AUTHORIZATION_GRANT As UInt32 = &H1
    Private Const BTSDK_AUTHORIZATION_DENY As UInt32 = &H2

    Private Const BTSDK_APP_EV_BASE As UInt32 = &H100
    '/* OPP specific event */
    Private Const BTSDK_APP_EV_OPP_BASE As UInt32 = &H200
    Private Const BTSDK_APP_EV_OPP_PULL As UInt32 = (BTSDK_APP_EV_OPP_BASE + 2)
    Private Const BTSDK_APP_EV_OPP_PUSH As UInt32 = (BTSDK_APP_EV_OPP_BASE + 3)
    Private Const BTSDK_APP_EV_OPP_PUSH_CARD As UInt32 = (BTSDK_APP_EV_OPP_BASE + 4)
    Private Const BTSDK_APP_EV_OPP_EXCHG As UInt32 = (BTSDK_APP_EV_OPP_BASE + 5)

    '/* FTP specific event */
    Private Const BTSDK_APP_EV_FTP_BASE As UInt32 = &H300
    Private Const BTSDK_APP_EV_FTP_PUT As UInt32 = (BTSDK_APP_EV_FTP_BASE + 0)
    Private Const BTSDK_APP_EV_FTP_GET As UInt32 = (BTSDK_APP_EV_FTP_BASE + 1)
    Private Const BTSDK_APP_EV_FTP_DEL_FILE As UInt32 = (BTSDK_APP_EV_FTP_BASE + 3)
    Private Const BTSDK_APP_EV_FTP_DEL_FOLDER As UInt32 = (BTSDK_APP_EV_FTP_BASE + 4)



    '//start bluetooth error extend
    'Private Const BTSDK_ER_FAIL_INITIALIZE_BTSDK As UInt32 = (BTSDK_ER_APPEXTEND_INDEX + &H6)


    'Private Const BTSDK_NTSERVICE_STATUS_FLAG As UInt32 = ?
    Public Const BTSDK_BLUETOOTH_STATUS_FLAG As UInt32 = &H2 '//status change about Bluetooth
    Public Const BTSDK_REFRESH_STATUS_FLAG As UInt32 = &H8

    Public Const BTSDK_UNPAIR_DEVICE As UInt32 = &H3
    Public Const BTSDK_PAIR_DEVICE As UInt32 = &H6
    Public Const BTSDK_DEL_DEVICE As UInt32 = &H15
    'seen 5's and 9's sometimes, and a lot of BTSDK_DEL_DEVICE when deleting.


    '//status change about Bluetooth
    Public Const BTSDK_BTSTATUS_TURNON As UInt32 = &H1
    Public Const BTSDK_BTSTATUS_TURNOFF As UInt32 = &H2
    Public Const BTSDK_BTSTATUS_HWPLUGGED As UInt32 = &H3
    Public Const BTSDK_BTSTATUS_HWPULLED As UInt32 = &H4


    '/* Possible roles for member 'role' in _BtSdkConnectionPropertyStru */
    Private Const BTSDK_CONNROLE_INITIATOR As UInt32 = &H2
    Private Const BTSDK_CONNROLE_ACCEPTOR As UInt32 = &H1


    '/* Type of Callback Indication */
    Private Const BTSDK_INQUIRY_RESULT_IND As UInt32 = &H4
    Private Const BTSDK_INQUIRY_COMPLETE_IND As UInt32 = &H5
    Private Const BTSDK_CONNECTION_EVENT_IND As UInt32 = &H9
    Private Const BTSDK_PIN_CODE_IND As UInt32 = &H0
    Private Const BTSDK_AUTHORIZATION_IND As UInt32 = &H6
    Private Const BTSDK_LINK_KEY_NOTIF_IND As UInt32 = &H2
    Private Const BTSDK_AUTHENTICATION_FAIL_IND As UInt32 = &H3

    '/*BT2.1 Supported indication*/
    Private Const BTSDK_IOCAP_REQ_IND As UInt32 = &H0C
    Private Const BTSDK_USR_CFM_REQ_IND As UInt32 = &H0D
    Private Const BTSDK_PASSKEY_REQ_IND As UInt32 = &H0E
    Private Const BTSDK_REM_OOBDATA_REQ_IND As UInt32 = &H0F
    Private Const BTSDK_PASSKEY_NOTIF_IND As UInt32 = &H10
    Private Const BTSDK_SIMPLE_PAIR_COMPLETE_IND As UInt32 = &H11
    Private Const BTSDK_OBEX_AUTHEN_REQ_IND As UInt32 = &H12
    Private Const BTSDK_VENDOR_EVENT_IND As UInt32 = &H0B
    Private Const BTSDK_CONNECTION_COMPLETE_IND As UInt32 = &H08
    Private Const BTSDK_DISCONNECTION_COMPLETE_IND As UInt32 = &H17
    Private Const BTSDK_DEVICE_FOUND_IND As UInt32 = &H19




    '/* Discovery Mode for Btsdk_SetDiscoveryMode() and Btsdk_GetDiscoveryMode() */
    Private Const BTSDK_GENERAL_DISCOVERABLE As UInt16 = &H1
    Private Const BTSDK_LIMITED_DISCOVERABLE As UInt16 = &H2
    Private Const BTSDK_DISCOVERABLE As UInt16 = BTSDK_GENERAL_DISCOVERABLE
    Private Const BTSDK_CONNECTABLE As UInt16 = &H4
    Private Const BTSDK_PAIRABLE As UInt16 = &H8
    Private Const BTSDK_DISCOVERY_DEFAULT_MODE As UInt16 = (BTSDK_DISCOVERABLE Or BTSDK_CONNECTABLE Or BTSDK_PAIRABLE)

    '/*for win32 only*/
    '/* PAN Event */
    Private Const BTSDK_PAN_EV_BASE As UInt32 = &H100
    Private Const BTSDK_PAN_EV_IP_CHANGE As UInt32 = BTSDK_PAN_EV_BASE + 1


    Private Const DEVICE_CLASS_MASK As UInt32 = &H1FFC

    '/* Default role of local device when creating a new ACL connection. */
    Private Const BTSDK_MASTER_ROLE As UInt32 = &H0
    Private Const BTSDK_SLAVE_ROLE As UInt32 = &H1

    '/* Possible values for "flag" parameter of Btsdk_StartEnumRemoteDevice. */
    Private Const BTSDK_ERD_FLAG_NOLIMIT As UInt32 = &H0
    Private Const BTSDK_ERD_FLAG_PAIRED As UInt32 = &H1
    Private Const BTSDK_ERD_FLAG_CONNECTED As UInt32 = &H2
    Private Const BTSDK_ERD_FLAG_INQUIRED As UInt32 = &H4
    Private Const BTSDK_ERD_FLAG_TRUSTED As UInt32 = &H20
    Private Const BTSDK_ERD_FLAG_DEVCLASS As UInt32 = &H10000

    '/* Possible values for "mask" member of BtSdkRemoteDevicePropertyStru structure. */
    Private Const BTSDK_RDPM_HANDLE As UInt32 = &H1
    Private Const BTSDK_RDPM_ADDRESS As UInt32 = &H2
    Private Const BTSDK_RDPM_NAME As UInt32 = &H4
    Private Const BTSDK_RDPM_CLASS As UInt32 = &H8
    Private Const BTSDK_RDPM_LMPINFO As UInt32 = &H10
    Private Const BTSDK_RDPM_LINKKEY As UInt32 = &H20

    '/* Possible ACL connection packet type */
    Private Const BTSDK_ACL_PKT_2DH1 As UInt32 = &H2      '/* Only supported by V2.0EDR */
    Private Const BTSDK_ACL_PKT_3DH1 As UInt32 = &H4      '/* Only supported by V2.0EDR */
    Private Const BTSDK_ACL_PKT_DM1 As UInt32 = &H8
    Private Const BTSDK_ACL_PKT_DH1 As UInt32 = &H10
    Private Const BTSDK_ACL_PKT_2DH3 As UInt32 = &H100      '/* Only supported by V2.0EDR */
    Private Const BTSDK_ACL_PKT_3DH3 As UInt32 = &H200      '/* Only supported by V2.0EDR */
    Private Const BTSDK_ACL_PKT_DM3 As UInt32 = &H400
    Private Const BTSDK_ACL_PKT_DH3 As UInt32 = &H800
    Private Const BTSDK_ACL_PKT_2DH5 As UInt32 = &H1000      '/* Only supported by V2.0EDR */
    Private Const BTSDK_ACL_PKT_3DH5 As UInt32 = &H2000      '/* Only supported by V2.0EDR */
    Private Const BTSDK_ACL_PKT_DM5 As UInt32 = &H4000
    Private Const BTSDK_ACL_PKT_DH5 As UInt32 = &H8000

    '/* Possible flags for member 'mask' in _BtSdkSDPSearchPatternStru */
    Private Const BTSDK_SSPM_UUID16 As UInt32 = &H1
    Private Const BTSDK_SSPM_UUID32 As UInt32 = &H2
    Private Const BTSDK_SSPM_UUID128 As UInt32 = &H4

    '/* Possible flags for member 'mask' in _BtSdkRemoteServiceAttrStru */
    Private Const BTSDK_RSAM_SERVICENAME As UInt32 = &H1
    Private Const BTSDK_RSAM_EXTATTRIBUTES As UInt32 = &H2

    Private Const BTSDK_MAX_SEARCH_PATTERNS As UInt32 = 12

    '/* Possible parameters for Btsdk_VDIInstallDev */
    Private Const HARDWAREID_MDMDUN As String = "{F12D3CF8-B11D-457e-8641-BE2AF2D6D204}\\MDMBTGEN336"
    Private Const HARDWAREID_MDMFAX As String = "{F12D3CF8-B11D-457e-8641-BE2AF2D6D204}\\MDMBTFAX"









    Private Function Btsdk_Func_IS_SAME_TYPE_DEVICE_CLASS(ByVal a As UInt32, ByVal b As UInt32) As Boolean

        Dim retBool As Boolean = False

        Dim tempA As UInt32 = (a And DEVICE_CLASS_MASK)
        Dim tempB As UInt32 = (b And DEVICE_CLASS_MASK)

        retBool = (tempA = tempB)

        Return retBool

    End Function

    Private Function BlueSoleil_UInt32toBool(ByVal inpUInt32 As UInt32) As Boolean

        If inpUInt32 = BTSDK_FALSE Then
            Return False
        Else
            Return True
        End If

    End Function



    Public Function BlueSoleil_Status_RegisterCallbacks() As Boolean

        Dim retBool As Boolean = False

        Dim retUInt32 As UInt32

        Dim funcPtr As IntPtr
        funcPtr = Marshal.GetFunctionPointerForDelegate(delegateBTstatusInfo)

        Dim funcPtrINT As UInt32 = CUInt(funcPtr)
        If funcPtr = IntPtr.Zero Then
            retUInt32 = Btsdk_RegisterGetStatusInfoCB4ThirdParty(funcPtrINT)
        Else
            retUInt32 = Btsdk_RegisterGetStatusInfoCB4ThirdParty(funcPtrINT)
        End If

        Dim statusFlags As UInt16 = BTSDK_BLUETOOTH_STATUS_FLAG Or BTSDK_REFRESH_STATUS_FLAG
        Dim setStatusBool As Boolean = BlueSoleil_SetStatusInfoFlag(statusFlags)


        'register connection event callback
        Dim funcPtr2 As IntPtr
        funcPtr2 = Marshal.GetFunctionPointerForDelegate(delegateBTconnEvent)
        Dim funcPtr2INT As UInt32 = CUInt(funcPtr2)
        Dim struCbkBytes(0 To 0) As Byte
        BlueSoleil_Stru_CallBack_Init(struCbkBytes, BTSDK_CONNECTION_EVENT_IND, funcPtr2INT)
        retUInt32 = Btsdk_RegisterCallback4ThirdParty(struCbkBytes(0))


        'register devicefound event callback.
        Dim funcPtr3 As IntPtr
        funcPtr3 = Marshal.GetFunctionPointerForDelegate(delegateBTdeviceFound)
        Dim funcPtr3INT As UInt32 = CUInt(funcPtr3)
        BlueSoleil_Stru_CallBack_Init(struCbkBytes, BTSDK_DEVICE_FOUND_IND, funcPtr3INT)
        retUInt32 = Btsdk_RegisterCallback4ThirdParty(struCbkBytes(0))


        If retUInt32 = BTSDK_OK Then
            retBool = setStatusBool
        Else
            retBool = False
        End If

        Return retBool

    End Function



    Public Function BlueSoleil_IsSDKinitialized() As Boolean

        Dim retUInt32 As UInt32
        Dim retBool As Boolean

        retUInt32 = Btsdk_IsSDKInitialized()
        retBool = BlueSoleil_UInt32toBool(retUInt32)

        Return retBool

    End Function



    Public Function BlueSoleil_IsBluetoothHardwarePresent() As Boolean

        Dim retUInt32 As UInt32
        Dim retBool As Boolean

        retUInt32 = Btsdk_IsBluetoothHardwareExisted()
        retBool = BlueSoleil_UInt32toBool(retUInt32)

        Return retBool

    End Function


    Public Function BlueSoleil_IsBluetoothReady() As Boolean

        Return True

        Dim retUInt32 As UInt32
        Dim retBool As Boolean

        retUInt32 = Btsdk_IsBluetoothReady()
        retBool = BlueSoleil_UInt32toBool(retUInt32)

        Return retBool

    End Function

    Public Function BlueSoleil_IsServerConnected() As Boolean

        '   Return True

        Dim retUInt32 As UInt32
        Dim retBool As Boolean

        retUInt32 = Btsdk_IsServerConnected()
        retBool = BlueSoleil_UInt32toBool(retUInt32)

        If retBool = False Then
            If retUInt32 <> BTSDK_FALSE Then
                retBool = True
            End If
        End If

        Return retBool

    End Function

    Public Sub BlueSoleil_Init()

        Btsdk_Init()

    End Sub


    Public Function BlueSoleil_StartBlueTooth() As Boolean

        Dim retUInt32 As UInt32
        Dim retBool As Boolean

        retUInt32 = Btsdk_StartBluetooth()

        If retUInt32 = BTSDK_OK Then
            retBool = True
        Else
            retBool = False
        End If

        Application.DoEvents()

        Return retBool

    End Function



    Public Function BlueSoleil_StopBlueTooth() As Boolean

        Dim retUInt32 As UInt32
        Dim retBool As Boolean

        retUInt32 = Btsdk_StopBluetooth()

        If retUInt32 = BTSDK_OK Then
            retBool = True
        Else
            retBool = False
        End If

        Return retBool

    End Function


    Public Function BlueSoleil_GetSDKDLLfilename() As String


        Dim tempStr As String = System.Environment.GetFolderPath(Environment.SpecialFolder.System)
        If Right(tempStr, 1) <> "\" Then tempStr = tempStr & "\"
        tempStr = tempStr & "bsSDK.dll"

        Return tempStr

    End Function

    Public Function BlueSoleil_IsInstalled() As Boolean

        Dim retBool As Boolean = False

        Dim tempStr As String = BlueSoleil_GetSDKDLLfilename()
        retBool = IO.File.Exists(tempStr)

        Return retBool

    End Function

    Public Sub BlueSoleil_Done()

        If BlueSoleil_IsServerConnected() = False Then
            Return
        End If


        Btsdk_Done()

    End Sub

    Private Sub BlueSoleil_Stru_ConnectionProperty_GetInfo(ByVal ptrStru_ConnProps As IntPtr, ByRef retRole As Byte, ByRef retResult As UInt32, ByRef retDvcHandle As UInt32, ByRef retSvcHandle As UInt32, ByRef retSvcClass As UInt16, ByRef retDurationSecs As UInt32, ByRef retBytesReceived As UInt32, ByRef retBytesSent As UInt32)

        'expecting 26 bytes.  

        Dim struLen As Integer = 26
        Dim evtData(0 To struLen - 1) As Byte

        'and now copy the whole thing from the pointer to our array.
        Marshal.Copy(ptrStru_ConnProps, evtData, 0, evtData.Length)



        'parse structure.  
        Dim currByteIdx As Integer = 0


        currByteIdx = currByteIdx + 4

        retDvcHandle = BitConverter.ToUInt32(evtData, currByteIdx)
        currByteIdx = currByteIdx + 4

        retSvcHandle = BitConverter.ToUInt32(evtData, currByteIdx)
        currByteIdx = currByteIdx + 4

        retSvcClass = BitConverter.ToUInt16(evtData, currByteIdx)
        currByteIdx = currByteIdx + 2

        retDurationSecs = BitConverter.ToUInt32(evtData, currByteIdx)
        currByteIdx = currByteIdx + 4

        retBytesReceived = BitConverter.ToUInt32(evtData, currByteIdx)
        currByteIdx = currByteIdx + 4

        retBytesSent = BitConverter.ToUInt32(evtData, currByteIdx)



    End Sub



    Private Sub BlueSoleil_Status_Callback_DeviceFound(ByVal dvcHandle As UInt32)

        Debug.Print("BlueSoleil_Status_Callback_DeviceFound - dvcHandle = " & dvcHandle)

        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_DeviceFound(dvcHandle))
        t.Start()

    End Sub

    Private Sub BlueSoleil_Status_Callback_InquiryResult(ByVal dvcHandle As UInt32)

        Debug.Print("BlueSoleil_Status_Callback_InquiryResult  dvcHandle = " & dvcHandle)

        'add dvc handle to array of found devices.
        ReDim Preserve BlueSoleil_Inquiry_DvcHandles(0 To BlueSoleil_Inquiry_DvcCount)
        BlueSoleil_Inquiry_DvcHandles(BlueSoleil_Inquiry_DvcCount) = dvcHandle
        BlueSoleil_Inquiry_DvcCount = BlueSoleil_Inquiry_DvcCount + 1

    End Sub

    Private Sub BlueSoleil_Status_Callback_InquiryComplete()

        Debug.Print("BlueSoleil_Status_Callback_InquiryComplete")

        BlueSoleil_Inquiry_IsComplete = True

    End Sub

    Private Sub BlueSoleil_Status_Callback_ConnEvent(ByVal connHandle As UInt32, ByVal evtType As UInt16, ByVal ptrArgs As IntPtr)

        Debug.Print("BlueSoleil_Status_Callback_ConnEvent EvtType = " & evtType)

        'ptrArgs is a ptr to a BtSdkConnectionPropertyStru

        Dim propRole As Byte
        Dim propResult As UInt32
        Dim propDvcHandle As UInt32
        Dim propSvcHandle As UInt32
        Dim propSvcClass As UInt16
        Dim propDuration As UInt32
        Dim propBytesReceived As UInt32
        Dim propBytesSent As UInt32
        If ptrArgs <> IntPtr.Zero Then
            BlueSoleil_Stru_ConnectionProperty_GetInfo(ptrArgs, propRole, propResult, propDvcHandle, propSvcHandle, propSvcClass, propDuration, propBytesReceived, propBytesSent)
        End If

        Select Case evtType

            Case BTSDK_APP_EV_CONN_IND
                Debug.Print("BlueSoleil_Status_Callback_ConnEvent EvtType = BTSDK_APP_EV_CONN_IND")
                'remote device connected to local svc
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_ServiceConnectedInbound(propDvcHandle, propSvcHandle, propSvcClass))
                t.Start()

            Case BTSDK_APP_EV_DISC_IND
                Debug.Print("BlueSoleil_Status_Callback_ConnEvent EvtType = BTSDK_APP_EV_DISC_IND")
                'remote device disconnected from local svc... OR device went out of range?
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_ServiceDisconnectedInbound(propDvcHandle, propSvcHandle, propSvcClass))
                t.Start()

            Case BTSDK_APP_EV_CONN_CFM
                Debug.Print("BlueSoleil_Status_Callback_ConnEvent EvtType = BTSDK_APP_EV_CONN_CFM")
                'local device connected to remote svc
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_ServiceConnectedOutbound(propDvcHandle, propSvcHandle, propSvcClass))
                t.Start()

            Case BTSDK_APP_EV_DISC_CFM
                Debug.Print("BlueSoleil_Status_Callback_ConnEvent EvtType = BTSDK_APP_EV_DISC_CFM")
                'local device disconnected from remote svc
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_ServiceDisconnectedOutbound(propDvcHandle, propSvcHandle, propSvcClass))
                t.Start()

            Case Else
                Debug.Print("BlueSoleil_Status_Callback_ConnEvent EvtType = UNKNOWN")

        End Select

    End Sub

    Private Sub BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo(ByVal msgType As UInt32, ByVal pulData As UInt32, ByVal funcParam As UInt32, ByVal funcArg As IntPtr)

        Debug.Print("BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo msgType = " & msgType & " pulData = " & pulData & "  funcParam = " & funcParam)


        If msgType = BTSDK_BLUETOOTH_STATUS_FLAG Then

            Select Case pulData

                Case BTSDK_BTSTATUS_TURNON

                    Debug.Print("BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo BTSDK_BTSTATUS_TURNON")
                    'RaiseEvent BlueSoleil_Status_TurnOn()
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_TurnOn())
                    t.Start()

                Case BTSDK_BTSTATUS_TURNOFF
                    Debug.Print("BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo BTSDK_BTSTATUS_TURNOFF")
                    'RaiseEvent BlueSoleil_Status_TurnOff()
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_TurnOff())
                    t.Start()

                Case BTSDK_BTSTATUS_HWPLUGGED
                    Debug.Print("BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo BTSDK_BTSTATUS_HWPLUGGED")
                    'RaiseEvent BlueSoleil_Status_Plugged()
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_Plugged())
                    t.Start()

                Case BTSDK_BTSTATUS_HWPULLED
                    Debug.Print("BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo BTSDK_BTSTATUS_HWPULLED")
                    'RaiseEvent BlueSoleil_Status_Unplugged()
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_Unplugged())
                    t.Start()

            End Select



        ElseIf msgType = BTSDK_REFRESH_STATUS_FLAG Then

            Select Case pulData

                Case BTSDK_INQUIRY_RESULT_IND
                    ReDim Preserve BlueSoleil_Inquiry_DvcHandles(0 To BlueSoleil_Inquiry_DvcCount)

                    BlueSoleil_Inquiry_DvcHandles(BlueSoleil_Inquiry_DvcCount) = funcParam
                    BlueSoleil_Inquiry_DvcCount = BlueSoleil_Inquiry_DvcCount + 1

                Case BTSDK_INQUIRY_COMPLETE_IND
                    BlueSoleil_Inquiry_IsComplete = True


                Case BTSDK_CONNECTION_EVENT_IND




                Case BTSDK_PAIR_DEVICE
                    Debug.Print("BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo BTSDK_PAIR_DEVICE")
                    'RaiseEvent BlueSoleil_Status_DevicePaired()
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_DevicePaired())
                    t.Start()

                Case BTSDK_UNPAIR_DEVICE
                    Debug.Print("BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo BTSDK_UNPAIR_DEVICE")
                    'RaiseEvent BlueSoleil_Status_DeviceUnpaired()
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_DeviceUnpaired())
                    t.Start()

                Case BTSDK_DEL_DEVICE
                    Debug.Print("BlueSoleil_Status_Callback_ReceiveBluetoothStatusInfo BTSDK_DEL_DEVICE")
                    'RaiseEvent BlueSoleil_Status_DeviceDeleted()
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Status_DeviceDeleted())
                    t.Start()

            End Select


        Else    'BTSDK_NTSERVICE_STATUS_FLAG

            Select Case pulData




            End Select

            'MsgBox(msgType)
        End If


    End Sub


    Public Function BlueSoleil_GetAllDevices_NamesAndHandles(ByRef retDvcNames() As String, ByRef retDvcHandles() As UInt32, ByRef retDvcCount As Integer, Optional ByVal maxWait_Seconds As UInt16 = 45, Optional ByVal maxDevicesFound As UInt16 = 0, Optional ByVal dvcClass_ToFind As UInt32 = 0) As Boolean

        Dim retBool As Boolean = True

        ReDim retDvcNames(0 To 0)
        ReDim retDvcHandles(0 To 0)
        retDvcCount = 0

        BlueSoleil_GetAllDevices(retDvcHandles, retDvcCount, maxWait_Seconds, maxDevicesFound, dvcClass_ToFind)

        If retDvcCount = 0 Then
            Return retBool
        End If

        Dim i As Integer
        ReDim retDvcNames(0 To retDvcCount - 1)
        For i = 0 To retDvcCount - 1
            retDvcNames(i) = BlueSoleil_GetRemoteDeviceName(retDvcHandles(i))
        Next i

        Return retBool

    End Function

    Public Function BlueSoleil_GetAllDevices(ByRef retDvcHandles() As UInt32, ByRef retDvcCount As Integer, Optional ByVal maxWait_Seconds As UInt16 = 45, Optional ByVal maxDevicesFound As UInt16 = 0, Optional ByVal dvcClass_ToFind As UInt32 = 0) As Boolean

        ReDim retDvcHandles(0 To 0)
        retDvcCount = 0
        Dim retBool As Boolean = False

        Dim retUInt32 As UInt32
        Dim struCbkBytes(0 To 0) As Byte

        'register inquiry-complete event callback.
        Dim funcPtr4 As IntPtr
        funcPtr4 = Marshal.GetFunctionPointerForDelegate(delegateBTinquiryCompleteInd)
        Dim funcPtr4INT As UInt32 = CUInt(funcPtr4)
        BlueSoleil_Stru_CallBack_Init(struCbkBytes, BTSDK_INQUIRY_COMPLETE_IND, funcPtr4INT)
        retUInt32 = Btsdk_RegisterCallback4ThirdParty(struCbkBytes(0))

        'register inquiry-result event callback.
        Dim funcPtr5 As IntPtr
        funcPtr5 = Marshal.GetFunctionPointerForDelegate(delegateBTinquiryResultInd)
        Dim funcPtr5INT As UInt32 = CUInt(funcPtr5)
        BlueSoleil_Stru_CallBack_Init(struCbkBytes, BTSDK_INQUIRY_RESULT_IND, funcPtr5INT)
        'retUInt32 = Btsdk_RegisterCallback4ThirdParty(struCbkBytes(0))
        retUInt32 = Btsdk_RegisterCallbackEx(struCbkBytes(0), BTSDK_CLIENTCBK_PRIORITY_MEDIUM)

        BlueSoleil_Inquiry_DvcCount = 0
        BlueSoleil_Inquiry_IsComplete = False

        Dim startTicks As Long = DateTime.UtcNow.Ticks \ TimeSpan.TicksPerMillisecond
        retUInt32 = Btsdk_StartDeviceDiscovery(dvcClass_ToFind, maxDevicesFound, maxWait_Seconds)

        Dim nowTicks As Long = 0
        Do
            Threading.Thread.Sleep(400)
            Application.DoEvents()

            If BlueSoleil_Inquiry_IsComplete = True Then Exit Do

            nowTicks = DateTime.UtcNow.Ticks \ TimeSpan.TicksPerMillisecond

            If (nowTicks - startTicks) / 1000 > maxWait_Seconds + 5 Then
                Exit Do
            End If

        Loop

        If BlueSoleil_Inquiry_IsComplete = False Then
            'uh-oh!

        Else
            retBool = True
        End If

        If BlueSoleil_Inquiry_DvcCount > 0 Then
            ReDim retDvcHandles(0 To BlueSoleil_Inquiry_DvcCount - 1)
            Array.Copy(BlueSoleil_Inquiry_DvcHandles, retDvcHandles, BlueSoleil_Inquiry_DvcCount)
            retDvcCount = BlueSoleil_Inquiry_DvcCount
        End If

        'unregister callbacks...
        ReDim struCbkBytes(0 To 0)
        BlueSoleil_Stru_CallBack_Init(struCbkBytes, BTSDK_INQUIRY_COMPLETE_IND, 0)
        retUInt32 = Btsdk_RegisterCallback4ThirdParty(struCbkBytes(0))

        ReDim struCbkBytes(0 To 0)
        BlueSoleil_Stru_CallBack_Init(struCbkBytes, BTSDK_INQUIRY_RESULT_IND, 0)
        retUInt32 = Btsdk_RegisterCallback4ThirdParty(struCbkBytes(0))


        Return retBool

    End Function

    Public Function BlueSoleil_GetPairedDevices_NamesAndHandles(ByRef retDvcNames() As String, ByRef retDvcHandles() As UInt32, ByRef retDvcCount As Integer) As Boolean

        Dim retBool As Boolean = True

        ReDim retDvcNames(0 To 0)
        ReDim retDvcHandles(0 To 0)
        retDvcCount = 0

        'retBool = BlueSoleil_GetPairedDevices(retDvcHandles, retDvcCount)

        'we changed this to BlueSoleil_GetStoredDevicesByClass so we can call it whether BlueTooth is currently enabled or not.
        retBool = BlueSoleil_GetStoredDevicesByClass(0, retDvcHandles, retDvcCount)
        Dim i As Integer
        If retDvcCount = 0 Then

            retBool = BlueSoleil_GetStoredDevicesByClass(0, retDvcHandles, retDvcCount)

            If retDvcCount = 0 Then
                retBool = BlueSoleil_GetPairedDevices(retDvcHandles, retDvcCount)

            End If

        End If

        If retDvcCount = 0 Then
            Return retBool
        End If

        ReDim retDvcNames(0 To retDvcCount - 1)
        For i = 0 To retDvcCount - 1
            retDvcNames(i) = BlueSoleil_GetRemoteDeviceName(retDvcHandles(i))
        Next i

        Return retBool

    End Function



    Public Function BlueSoleil_GetInquiredDevices_NamesAndHandles(ByRef retDvcNames() As String, ByRef retDvcHandles() As UInt32, ByRef retDvcCount As Integer) As Boolean

        Dim retBool As Boolean = True

        ReDim retDvcNames(0 To 0)
        ReDim retDvcHandles(0 To 0)
        retDvcCount = 0

        retBool = BlueSoleil_GetInquiredDevices(retDvcHandles, retDvcCount)

        If retDvcCount = 0 Then
            Return retBool
        End If

        ReDim retDvcNames(0 To retDvcCount - 1)

        Dim i As Integer
        For i = 0 To retDvcCount - 1
            retDvcNames(i) = BlueSoleil_GetRemoteDeviceName(retDvcHandles(i))
        Next i

        Return retBool

    End Function


    Public Function BlueSoleil_GetStoredDevicesByClass(ByVal bsClassCode As UInt32, ByRef dvcHandles() As UInt32, ByRef dvcCount As Integer) As Boolean


        ReDim dvcHandles(0 To 0)
        dvcCount = 0

        'first, get total handle count.
        Dim tempArray(0 To 0) As UInt32


        ' bsClassCode = DEVICE_CLASS_MASK Or BTSDK_DEVCLS_PHONE Or BTSDK_SRVCLS_TELEPHONE Or BTSDK_SRVCLS_AUDIO Or BTSDK_SRVCLS_OBJECT

        Dim retCount As Integer = 0 ' Btsdk_GetStoredDevicesByClass_ByVal(bsClassCode, 0, 0)

        retCount = Btsdk_GetStoredDevicesByClass(bsClassCode, tempArray(0), CUInt(retCount))

        If retCount = 0 Then
            Return True
            Exit Function
        End If


        dvcCount = retCount
        ReDim tempArray(0 To retCount - 1)

        retCount = Btsdk_GetStoredDevicesByClass(bsClassCode, tempArray(0), CUInt(retCount))

        dvcHandles = tempArray

        Return True

    End Function



    Public Function BlueSoleil_GetPairedDevices(ByRef dvcHandles() As UInt32, ByRef dvcCount As Integer) As Boolean

        ReDim dvcHandles(0 To 0)
        dvcCount = 0

        'first, get total handle count.
        Dim tempArray(0 To 0) As UInt32

        Dim retCount As Integer = Btsdk_GetPairedDevices_ByValArray(tempArray(0), 0)

        If retCount = 0 Then
            Return True
            Exit Function
        End If

        dvcCount = retCount
        ReDim tempArray(0 To retCount - 1)

        retCount = Btsdk_GetPairedDevices(tempArray(0), CUInt(retCount))

        dvcHandles = tempArray

        Return True

    End Function


    Public Function BlueSoleil_GetInquiredDevices(ByRef dvcHandles() As UInt32, ByRef dvcCount As Integer) As Boolean

        ReDim dvcHandles(0 To 0)
        dvcCount = 0

        'first, get total handle count.
        Dim tempArray(0 To 0) As UInt32

        Dim retCount As Integer = Btsdk_GetInquiredDevices_ByValArray(tempArray(0), 0)

        If retCount = 0 Then
            Return True
            Exit Function
        End If

        dvcCount = retCount
        ReDim tempArray(0 To retCount - 1)

        retCount = Btsdk_GetInquiredDevices(tempArray(0), CUInt(retCount))

        dvcHandles = tempArray

        Return True

    End Function


    Public Function BlueSoleil_IsDeviceConnected(ByVal dvcHandle As UInt32) As Boolean

        Dim retUInt32 As UInt32
        Dim retBool As Boolean

        retUInt32 = Btsdk_IsDeviceConnected(dvcHandle)
        retBool = BlueSoleil_UInt32toBool(retUInt32)

        Return retBool

    End Function


    Public Function BlueSoleil_GetRemoteDeviceName(ByVal dvcHandle As UInt32) As String

        Dim retUInt32 As UInt32
        Dim retStr As String = ""

        Dim byteArray(0 To BTSDK_DEVNAME_LEN - 1) As Byte

        Dim retCount As UInt16 = BTSDK_DEVNAME_LEN

        ReDim byteArray(0 To retCount - 1)
        retUInt32 = Btsdk_GetRemoteDeviceName(dvcHandle, byteArray(0), retCount)

        If retCount < 1 Then
            ReDim byteArray(0 To BTSDK_DEVNAME_LEN - 1)
            retCount = BTSDK_DEVNAME_LEN
            retUInt32 = Btsdk_UpdateRemoteDeviceName(dvcHandle, byteArray(0), retCount)

        End If

        If retCount < 1 Then
            Return ""
            Exit Function
        End If

        ReDim Preserve byteArray(0 To retCount - 1)

        If retUInt32 = BTSDK_OK Then
            retStr = System.Text.Encoding.UTF8.GetString(byteArray)
            retStr = Replace(retStr, Chr(0), "")
        Else
            retStr = ""
        End If



        Return retStr

    End Function


    Public Function BlueSoleil_GetRemoteDeviceName_Refresh(ByVal dvcHandle As UInt32) As String

        Dim retUInt32 As UInt32
        Dim retStr As String = ""

        Dim byteArray(0 To BTSDK_DEVNAME_LEN - 1) As Byte

        Dim retCount As UInt16 = BTSDK_DEVNAME_LEN


        ReDim byteArray(0 To BTSDK_DEVNAME_LEN - 1)
        retCount = BTSDK_DEVNAME_LEN
        retUInt32 = Btsdk_UpdateRemoteDeviceName(dvcHandle, byteArray(0), retCount)

        If retCount < 1 Then
            Return ""
            Exit Function
        End If

        ReDim Preserve byteArray(0 To retCount - 1)

        If retUInt32 = BTSDK_OK Then
            retStr = System.Text.Encoding.UTF8.GetString(byteArray)
            retStr = Replace(retStr, Chr(0), "")
        Else
            retStr = ""
        End If



        Return retStr

    End Function


    Public Function BlueSoleil_GetRemoteDeviceClass(ByVal dvcHandle As UInt32, ByRef retDvcClass As UInt32) As Boolean

        Dim retUInt32 As UInt32

        retUInt32 = Btsdk_GetRemoteDeviceClass(dvcHandle, retDvcClass)


        'Dim tempClsCompare As UInt32 = BTSDK_DEVCLS_MASK(BTSDK_DEVCLS_PHONE And DEVICE_CLASS_MASK)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_GetRemoteDeviceServiceHandles(ByVal dvcHandle As UInt32, ByRef svcHandleArray() As UInt32, ByRef svcHandleCount As Integer) As Boolean

        'device must be CONNECTED in order to get Services.

        ReDim svcHandleArray(0 To 0)
        svcHandleCount = 0

        Dim tempArray(0 To 0) As UInt32

        Dim retCount As UInt32

        Dim retUInt32 As UInt32 = Btsdk_GetRemoteServices_ByValArray(dvcHandle, tempArray(0), retCount)

        If retCount = 0 Then
            Return True
            Exit Function
        End If

        svcHandleCount = CInt(retCount)
        ReDim tempArray(0 To svcHandleCount - 1)

        retUInt32 = Btsdk_GetRemoteServices(dvcHandle, tempArray(0), retCount)

        svcHandleArray = tempArray

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If


    End Function

    Public Function BlueSoleil_GetRemoteDeviceServiceHandles_Refresh(ByVal dvcHandle As UInt32, ByRef svcClassArray() As UInt32, ByRef svcClassCount As Integer) As Boolean


        ReDim svcClassArray(0 To 0)
        svcClassCount = 0

        If dvcHandle = 0 Then
            Return False
        End If


        Dim tempArray(0 To 0) As UInt32

        Dim retCount As UInt32

        Dim retUInt32 As UInt32 = Btsdk_BrowseRemoteServices_ByValArray(dvcHandle, tempArray(0), retCount)

        If retCount = 0 Then
            Return True
            Exit Function
        End If

        svcClassCount = CInt(retCount)
        ReDim tempArray(0 To svcClassCount - 1)

        retUInt32 = Btsdk_BrowseRemoteServices(dvcHandle, tempArray(0), CUInt(retCount))

        'retUInt32 = Btsdk_GetRemoteServices(dvcHandle, tempArray(0), CUInt(retCount))


        svcClassArray = tempArray

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub BlueSoleil_Stru_RemoteServiceAttribs_GetInfo(ByRef bArray() As Byte, ByRef retSvcClass As UInt16, ByRef retSvcName As String)

        retSvcName = ""

        Dim tempBytes(0 To BTSDK_SERVICENAME_MAXLENGTH - 1) As Byte
        Array.Copy(bArray, 8, tempBytes, 0, BTSDK_SERVICENAME_MAXLENGTH)
        retSvcName = System.Text.Encoding.UTF8.GetString(tempBytes)
        retSvcName = Replace(retSvcName, Chr(0), "")

        retSvcClass = BitConverter.ToUInt16(bArray, 2)

    End Sub



    Private Sub BlueSoleil_Stru_CallBack_Init(ByRef bArray() As Byte, ByVal msgTypeToHook As UInt16, ByVal cbFuncPointer As UInt32)

        'msgTypes:  BTSDK_CONNECTION_EVENT_IND, BTSDK_AUTHENTICATION_FAIL_IND, etc.

        Dim struSize As Integer = 2 + 4
        ReDim bArray(0 To struSize - 1)

        Dim tempBytes(0 To 1) As Byte

        tempBytes = BitConverter.GetBytes(msgTypeToHook)
        Array.Copy(tempBytes, 0, bArray, 0, 2)

        ReDim tempBytes(0 To 3)
        tempBytes = BitConverter.GetBytes(cbFuncPointer)
        Array.Copy(tempBytes, 0, bArray, 2, 4)


    End Sub



    Private Sub BlueSoleil_Stru_VendorCmd_Init(ByRef bArray() As Byte, ByVal ocfOpcodeCommandField As UInt16, ByRef vendorCommands() As Byte, ByVal vendorCommandsLen As Byte)


        Dim struSize As Integer = 2 + vendorCommandsLen + 1
        ReDim bArray(0 To struSize - 1)

        Dim tempBytes(0 To 1) As Byte

        tempBytes = BitConverter.GetBytes(ocfOpcodeCommandField)
        Array.Copy(tempBytes, 0, bArray, 0, 2)

        bArray(2) = vendorCommandsLen

        Array.Copy(vendorCommands, 0, bArray, 3, vendorCommandsLen)


    End Sub


    Private Sub BlueSoleil_Stru_LocalServiceAttribs_Init(ByRef bArray() As Byte)

        Dim struSize As Integer = 2 + 2 + BTSDK_SERVICENAME_MAXLENGTH + 2 + 2 '+ ? + 4

        struSize = 2 + 2 + BTSDK_SERVICENAME_MAXLENGTH + 2 + 2 + 100  'ensure large enough.

        ReDim bArray(0 To struSize - 1)

        Dim tempMask As UInt32 = BTSDK_RSAM_SERVICENAME

        Dim temp4bytes(0 To 3) As Byte


        'copy the MASK value into the structure, as position 0.
        temp4bytes = BitConverter.GetBytes(tempMask)
        Array.Copy(temp4bytes, 0, bArray, 0, 4)



    End Sub

    Private Sub BlueSoleil_Stru_LocalServiceAttribs_GetInfo(ByRef bArray() As Byte, ByRef retSvcClass As UInt16, ByRef retSvcName As String)

        retSvcName = ""

        Dim tempBytes(0 To BTSDK_SERVICENAME_MAXLENGTH - 1) As Byte
        Array.Copy(bArray, 4, tempBytes, 0, BTSDK_SERVICENAME_MAXLENGTH)
        retSvcName = System.Text.Encoding.UTF8.GetString(tempBytes)
        retSvcName = Replace(retSvcName, Chr(0), "")

        retSvcClass = BitConverter.ToUInt16(bArray, 2)

    End Sub

    Private Sub BlueSoleil_Stru_RemoteServiceAttribs_Init(ByRef bArray() As Byte)

        Dim struSize As Integer = 4 + 2 + 4 + BTSDK_SERVICENAME_MAXLENGTH + 4 + 2


        ReDim bArray(0 To struSize - 1)

        Dim tempMask As UInt32 = BTSDK_RSAM_SERVICENAME

        Dim temp4bytes(0 To 3) As Byte


        'copy the MASK value into the structure, as position 0.
        temp4bytes = BitConverter.GetBytes(tempMask)
        Array.Copy(temp4bytes, 0, bArray, 0, 4)


    End Sub




    Private Sub BlueSoleil_Stru_ConnectionProperty_Init(ByRef bArray() As Byte)

        Dim struSize As Integer = 4 + 4 + 4 + 2 + 4 + 4 + 4
        'or is it  4 + 4 + 4 + 4 + 2 + 4 + 4 + 4

        ReDim bArray(0 To struSize - 1)




    End Sub


    Public Function BlueSoleil_GetConnectionProperties(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False 

        Dim struConnProp(0 To 0) As Byte
        BlueSoleil_Stru_ConnectionProperty_Init(struConnProp)
        Dim retUInt As UInt32 = Btsdk_GetConnectionProperty(connHandle, struConnProp(0))

        Return (retUInt = BTSDK_OK)

    End Function



    Public Function BlueSoleil_GetRemoteServiceAttributes(ByVal svcHandle As UInt32, ByRef retSvcName As String, ByRef retSvcClass As UInt16) As Boolean


        Dim struRemoteServiceAttrs(0 To 0) As Byte
        BlueSoleil_Stru_RemoteServiceAttribs_Init(struRemoteServiceAttrs)

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_GetRemoteServiceAttributes(svcHandle, struRemoteServiceAttrs(0))

        BlueSoleil_Stru_RemoteServiceAttribs_GetInfo(struRemoteServiceAttrs, retSvcClass, retSvcName)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_ConnectService_BySvcHandle(ByVal dvcHandle As UInt32, ByVal svcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_Connect(svcHandle, 0, retConnHandle)
        'retUInt32 = Btsdk_ConnectEx(dvcHandle, svcHandle, 0, retConnHandle)

        If retUInt32 = 778 Then
            'already connected?   i dont F'ing know.
            BlueSoleil_StopService(svcHandle)
            Btsdk_Disconnect(svcHandle)
            BlueSoleil_StartService(svcHandle)
            retUInt32 = Btsdk_Connect(svcHandle, 0, retConnHandle)
        End If



        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_ConnectService_ByDvcHandle_PBAP(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_PBAP_PSE, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandle_MAP(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_OBEX_MESSAGEACCESSSERVER, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandle_PAN(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_PAN_NAP, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandle_SPP(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_SERIAL_PORT, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandle_OPP(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_OBEX_OBJ_PUSH, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandle_FTP(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_OBEX_FILE_TRANS, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandle_HFP(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_HANDSFREE_AG, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandle_AVRCP(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_AVRCP_TG, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandle_A2DP(ByVal dvcHandle As UInt32, ByRef retConnHandle As UInt32) As Boolean
        Return BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandle, BTSDK_CLS_AUDIO_SOURCE, retConnHandle)
    End Function

    Public Function BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(ByVal dvcHandle As UInt32, ByVal svcClass As UInt16, ByRef retConnHandle As UInt32) As Boolean

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_ConnectEx(dvcHandle, svcClass, 0, retConnHandle)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function BlueSoleil_DisconnectServiceConn(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then
            Return True
        End If

        If BlueSoleil_IsServerConnected() = False Then
            Return False
        End If

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_Disconnect(connHandle)


        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_StartService(ByVal svcHandle As UInt32) As Boolean


        If BlueSoleil_IsServerConnected() = False Then
            Return False
        End If

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_StartServer(svcHandle)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function BlueSoleil_StopService(ByVal svcHandle As UInt32) As Boolean

        If svcHandle = 0 Then Return True

        If BlueSoleil_IsServerConnected() = False Then
            Return False
        End If

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_StopServer(svcHandle)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function







    Public Sub BlueSoleil_Status_UnregisterCallbacks()

        Dim retUInt32 As UInt32 = 0

        BlueSoleil_SetStatusInfoFlag(0)

        retUInt32 = Btsdk_RegisterGetStatusInfoCB4ThirdParty(0)

        Dim struCbkBytes(0 To 0) As Byte
        BlueSoleil_Stru_CallBack_Init(struCbkBytes, BTSDK_CONNECTION_EVENT_IND, 0)
        retUInt32 = Btsdk_RegisterCallback4ThirdParty(struCbkBytes(0))

        ReDim struCbkBytes(0 To 0)
        BlueSoleil_Stru_CallBack_Init(struCbkBytes, BTSDK_DEVICE_FOUND_IND, 0)
        retUInt32 = Btsdk_RegisterCallback4ThirdParty(struCbkBytes(0))



    End Sub


    Public Function BlueSoleil_SetStatusInfoFlag(ByVal msgTypes As UInt16) As Boolean


        Dim retUInt32 As UInt32 = Btsdk_SetStatusInfoFlag(msgTypes)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_GetLocalDeviceServices(ByRef svcClassArray() As UInt16, ByRef svcHandleArray() As UInt32, ByRef svcNameArray() As String, ByRef svcCount As Integer) As Boolean

        Dim hEnum As UInt32 = Btsdk_StartEnumLocalServer()

        Dim struLocalServerAttribs(0 To 0) As Byte

        Dim tempHandle As UInt32, tempClass As UInt16, tempName As String = ""

        svcCount = 0

        Do
            BlueSoleil_Stru_LocalServiceAttribs_Init(struLocalServerAttribs)

            tempHandle = Btsdk_EnumLocalServer(hEnum, struLocalServerAttribs(0))

            If tempHandle = BTSDK_INVALID_HANDLE Then Exit Do

            BlueSoleil_Stru_LocalServiceAttribs_GetInfo(struLocalServerAttribs, tempClass, tempName)

            ReDim Preserve svcClassArray(0 To svcCount)
            ReDim Preserve svcHandleArray(0 To svcCount)
            ReDim Preserve svcNameArray(0 To svcCount)
            svcClassArray(svcCount) = tempClass
            svcHandleArray(svcCount) = tempHandle
            svcNameArray(svcCount) = tempName

            svcCount = svcCount + 1
        Loop

        Return (svcCount > 0)

    End Function





    Public Function BlueSoleil_GetLocalDeviceClass(ByRef retDvcClass_AndSvcClass As UInt32) As Boolean

        Dim retUInt32 As UInt32 = Btsdk_GetLocalDeviceClass(retDvcClass_AndSvcClass)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_SvcClass_IsHandsFreeAG(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            ' Case BTSDK_CLS_HANDSFREE
            '     retBool = True

            Case BTSDK_CLS_HANDSFREE_AG
                retBool = True

        End Select

        Return retBool

    End Function


    Public Function BlueSoleil_SvcClass_IsHandsFree(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_HEADSET_AG
                retBool = True

        End Select

        Return retBool

    End Function




    Public Function BlueSoleil_SvcClass_IsA2DP(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_AUDIO_SOURCE ', BTSDK_CLS_ADV_AUDIO_DISTRIB 
                retBool = True

        End Select

        Return retBool

    End Function


    Public Function BlueSoleil_SvcClass_IsSPP(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_SERIAL_PORT
                retBool = True

        End Select

        Return retBool

    End Function


    Public Function BlueSoleil_SvcClass_IsOPP(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_OBEX_OBJ_PUSH
                retBool = True

        End Select

        Return retBool

    End Function


    Public Function BlueSoleil_SvcClass_IsPBAP(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_PBAP_PSE
                retBool = True

        End Select

        Return retBool

    End Function


    Public Function BlueSoleil_SvcClass_IsPAN(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_PAN_NAP   'BTSDK_CLS_PAN_NAP ', BTSDK_CLS_PAN_GN ', BTSDK_CLS_PAN_PANU
                retBool = True

        End Select

        Return retBool

    End Function


    Public Function BlueSoleil_SvcClass_IsMAP(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_OBEX_MESSAGEACCESSSERVER ', BTSDK_CLS_OBEX_OBJ_PUSH   ', BTSDK_CLS_OBEX_FILE_TRANS 
                retBool = True

        End Select

        Return retBool

    End Function





    Public Function BlueSoleil_SvcClass_IsFTP(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_OBEX_FILE_TRANS
                retBool = True

        End Select

        Return retBool

    End Function


    Public Function BlueSoleil_SvcClass_IsAVRCP(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_AVRCP_TG
                retBool = True

        End Select

        Return retBool

    End Function


    Public Function BlueSoleil_SvcClass_IsHeadsetAG(ByVal svcClass As UInt16) As Boolean

        Dim retBool As Boolean = False

        Select Case svcClass

            Case BTSDK_CLS_HEADSET_AG
                retBool = True

                'Case BTSDK_CLS_HEADSET
                '    retBool = True

        End Select

        Return retBool

    End Function






    Public Function BlueSoleil_SetLocalDeviceServiceClass(ByVal isDesktop As Boolean, ByVal isNetWork As Boolean, ByVal isObject As Boolean) As Boolean

        Dim dvcClass As UInt32 = 0
        If isDesktop = True Then dvcClass = dvcClass Or BTSDK_COMPCLS_DESKTOP
        If isNetWork = True Then dvcClass = dvcClass Or BTSDK_SRVCLS_NETWORK
        If isObject = True Then dvcClass = dvcClass Or BTSDK_SRVCLS_OBJECT

        Dim isAudio As Boolean = True
        If isAudio = True Then dvcClass = dvcClass Or BTSDK_DEVCLS_AUDIO



        Dim retUInt As UInt32

        'retUInt = Btsdk_SetLocalDeviceClass(dvcClass)

        retUInt = Btsdk_SetLocalDeviceClass(&H240404)

        Return (retUInt = BTSDK_OK)


    End Function


    Public Function BlueSoleil_ConnectService_ByName(ByVal profileName As String, ByVal profileDvcName As String, ByRef retDvcHandle As UInt32, ByRef retConnHandle As UInt32, ByRef retSvcHandle As UInt32) As Boolean

        If BlueSoleil_IsInstalled() = False Then Return False
        If BlueSoleil_IsBluetoothReady() = False Then Return False


        retDvcHandle = 0
        retConnHandle = 0
        retSvcHandle = 0


        Dim retBool As Boolean = False

        Dim dvcHandles(0 To 0) As UInt32, dvcNames(0 To 0) As String, dvcCount As Integer = 0

        'do this a second time after starting BlueTooth... just for the hell of it.
        BlueSoleil_GetPairedDevices_NamesAndHandles(dvcNames, dvcHandles, dvcCount)

        Dim TorF As Boolean = False
        If dvcCount = 0 Then
            Return False
        End If


        'time to connect to service.
        Dim svcHandles(0 To 0) As UInt32, svcCount As Integer
        Dim svcName As String = "", svcClass As UInt16
        Dim svcConnHandle As UInt32 = 0, connTorF As Boolean = False

        Dim matchDvcHandle As UInt32 = 0

        Dim i As Integer, j As Integer
        'find default device.
        For i = 0 To dvcCount - 1
            If dvcNames(i) = profileDvcName Then

                matchDvcHandle = dvcHandles(i)

                'try direct connect
                Select Case profileName
                    Case "PBAP"
                        connTorF = BlueSoleil_ConnectService_ByDvcHandle_PBAP(matchDvcHandle, retConnHandle)
                        retBool = connTorF And (retConnHandle <> 0)


                    Case "MAP"
                        connTorF = BlueSoleil_ConnectService_ByDvcHandle_MAP(matchDvcHandle, retConnHandle)
                        retBool = connTorF And (retConnHandle <> 0)


                    Case "PAN"
                        connTorF = BlueSoleil_ConnectService_ByDvcHandle_PAN(matchDvcHandle, retConnHandle)
                        retBool = connTorF And (retConnHandle <> 0)

                    Case "HFP"
                        connTorF = BlueSoleil_ConnectService_ByDvcHandle_HFP(matchDvcHandle, retConnHandle)
                        retBool = connTorF And (retConnHandle <> 0)

                    Case "OPP"
                        connTorF = BlueSoleil_ConnectService_ByDvcHandle_OPP(matchDvcHandle, retConnHandle)
                        retBool = connTorF And (retConnHandle <> 0)

                    'removed because we might want service handle.
                    ' Case "SPP"
                    '     connTorF = BlueSoleil_ConnectService_ByDvcHandle_SPP(matchDvcHandle, retConnHandle)
                    '     retBool = connTorF And (retConnHandle <> 0)

                    Case "AVRCP"
                        connTorF = BlueSoleil_ConnectService_ByDvcHandle_AVRCP(matchDvcHandle, retConnHandle)
                        retBool = connTorF And (retConnHandle <> 0)
                        retBool = retBool


                    Case "A2DP"
                        connTorF = BlueSoleil_ConnectService_ByDvcHandle_A2DP(matchDvcHandle, retConnHandle)
                        retBool = connTorF And (retConnHandle <> 0)
                        retBool = retBool

                    Case "FTP"
                        connTorF = BlueSoleil_ConnectService_ByDvcHandle_FTP(matchDvcHandle, retConnHandle)
                        retBool = connTorF And (retConnHandle <> 0)
                        retBool = retBool

                End Select
                If connTorF = True Then
                    retDvcHandle = matchDvcHandle
                End If


                If retBool = False Then
                    'find service.
                    BlueSoleil_GetRemoteDeviceServiceHandles(dvcHandles(i), svcHandles, svcCount)

                    If svcCount = 0 Then
                        BlueSoleil_GetRemoteDeviceServiceHandles_Refresh(dvcHandles(i), svcHandles, svcCount)
                        BlueSoleil_GetRemoteDeviceServiceHandles(dvcHandles(i), svcHandles, svcCount)
                    End If


                    For j = 0 To svcCount - 1
                        BlueSoleil_GetRemoteServiceAttributes(svcHandles(j), svcName, svcClass)

                        Select Case profileName
                            Case "PAN"
                                If BlueSoleil_SvcClass_IsPAN(svcClass) = True Then
                                    'connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    connTorF = BlueSoleil_ConnectService_ByDvcHandleAndSvcClass(dvcHandles(i), svcClass, svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If


                            Case "MAP"
                                If BlueSoleil_SvcClass_IsMAP(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If


                            Case "PBAP"
                                If BlueSoleil_SvcClass_IsPBAP(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If


                            Case "SPP"
                                If BlueSoleil_SvcClass_IsSPP(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If

                            Case "OPP"
                                If BlueSoleil_SvcClass_IsOPP(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If

                            Case "AVRCP"
                                If BlueSoleil_SvcClass_IsAVRCP(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    ' If connTorF = True Then ' And svcConnHandle <> 0 Then
                                    retConnHandle = svcConnHandle
                                    retDvcHandle = dvcHandles(i)
                                    retSvcHandle = svcHandles(j)
                                    connTorF = (retDvcHandle <> 0) ' And retSvcHandle <> 0)
                                    retBool = True
                                    Exit For
                                    ' End If
                                End If


                            Case "HFP"
                                If BlueSoleil_SvcClass_IsHandsFreeAG(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If



                            Case "HSP"
                                If BlueSoleil_SvcClass_IsHeadsetAG(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If

                            Case "HFU"
                                If BlueSoleil_SvcClass_IsHandsFree(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If


                            Case "A2DP"     '??
                                If BlueSoleil_SvcClass_IsA2DP(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If



                            Case "FTP"
                                If BlueSoleil_SvcClass_IsFTP(svcClass) = True Then
                                    connTorF = BlueSoleil_ConnectService_BySvcHandle(dvcHandles(i), svcHandles(j), svcConnHandle)
                                    If connTorF = True And svcConnHandle <> 0 Then
                                        retConnHandle = svcConnHandle
                                        retDvcHandle = dvcHandles(i)
                                        retSvcHandle = svcHandles(j)
                                        retBool = True
                                        Exit For
                                    End If
                                End If


                            Case Else
                                'unknown profile?
                                '
                                MsgBox("Unknown Bluetooth profile:  " & profileName)


                        End Select


                    Next j

                End If

            End If

            If retBool = True Then Exit For
        Next i






        Return retBool


    End Function



    Public Function BlueSoleil_GetRemoteRSSI_Decibles(ByVal dvcHandle As UInt32) As Double

        Dim tempSByte As SByte
        Dim retUInt As UInt32 = Btsdk_GetRemoteRSSI(dvcHandle, tempSByte)

        Return tempSByte

    End Function



    Public Function BlueSoleil_GetRemoteLinkQualityPct(ByVal dvcHandle As UInt32) As Double

        Dim retPct As Double = -1

        Dim retUInt As UInt32 = 0
        Dim retLink As UInt16 = 0
        retUInt = Btsdk_GetRemoteLinkQuality(dvcHandle, retLink)

        If retUInt = BTSDK_OK Then
            If retLink > &HFF Then
                retLink = retLink
            Else
                retPct = 100 * retLink / &HFF
            End If
        Else
            retPct = -1
        End If

        Return retPct

    End Function


    Public Function BlueSoleil_GetRemoteLinkDataStatistics(ByVal dvcHandle As UInt32, ByRef rcvdBytes As UInt32, ByRef sentBytes As UInt32) As Double

        Dim retPct As Double = -1

        Dim retUInt As UInt32 = 0
        Dim retRcv As UInt32, retSent As UInt32
        retUInt = Btsdk_RemoteDeviceFlowStatistic(dvcHandle, retRcv, retSent)

        If retUInt = BTSDK_OK Then
            rcvdBytes = retRcv
            sentBytes = retSent
        End If


        Return retPct

    End Function

    Public Function BlueSoleil_DevicePairing_PairDevice(ByVal dvcHandle As UInt32) As Boolean

        Dim retUInt As UInt32
        retUInt = Btsdk_PairDevice(dvcHandle)

        Return retUInt = BTSDK_OK

    End Function

    Public Function BlueSoleil_DevicePairing_UnpairDevice(ByVal dvcHandle As UInt32) As Boolean

        Dim retUInt As UInt32
        retUInt = Btsdk_UnPairDevice(dvcHandle)

        Return retUInt = BTSDK_OK

    End Function

    Public Function BlueSoleil_DevicePairing_DeleteDevice(ByVal dvcHandle As UInt32) As Boolean

        Dim retUInt As UInt32
        retUInt = Btsdk_DeleteRemoteDeviceByHandle(dvcHandle)

        Return retUInt = BTSDK_OK

    End Function

    Public Function BlueSoleil_DevicePairing_IsDevicePaired(ByVal dvcHandle As UInt32) As Boolean

        Dim retisPaired As Byte

        Dim retUInt As UInt32
        retUInt = Btsdk_IsDevicePaired(dvcHandle, retisPaired)

        If retUInt <> BTSDK_OK Then
            'failed.
            Return False
        End If

        Return (retisPaired <> BTSDK_FALSE)

    End Function

    Public Function BlueSoleil_GetRemoteDeviceAddress(ByVal dvcHandle As UInt32) As String

        Dim array6bytes(0 To 5) As Byte

        Dim retUInt As UInt32
        retUInt = Btsdk_GetRemoteDeviceAddress(dvcHandle, array6bytes(0))

        Dim retStr As String = ""

        Dim i As Integer
        For i = array6bytes.Length - 1 To 0 Step -1
            retStr = retStr & Strings.Right("00" & Hex(array6bytes(i)), 2)
            If i <> 0 Then
                retStr = retStr & ":"
            End If
        Next i

        Return retStr

    End Function


    Public Function BlueSoleil_GetTickCount() As Long

        'this is a simple GetTickCount function.  Added to the Blue Soleil library for simplicity.

        Return DateTime.UtcNow.Ticks \ TimeSpan.TicksPerMillisecond

    End Function

    Public Sub BlueSoleil_TimeOut(ByVal dblSeconds As Double)

        Dim ticksStartTime As Long
        Dim ticksEndTime As Long
        Dim currTime As Long

        ticksStartTime = BlueSoleil_GetTickCount()
        ticksEndTime = ticksStartTime + CLng(dblSeconds * 1000)

        Do
            currTime = BlueSoleil_GetTickCount()
            If currTime >= ticksEndTime OrElse currTime < ticksStartTime Then Exit Do

            Threading.Thread.Sleep(25)

            currTime = BlueSoleil_GetTickCount()
            If currTime >= ticksEndTime OrElse currTime < ticksStartTime Then Exit Do

            System.Windows.Forms.Application.DoEvents()
            Windows.Forms.Application.DoEvents()
        Loop

    End Sub

    Private Sub BlueSoleil_Stru_RemoteDeviceProp_Init(ByRef struBytes() As Byte)

        Dim LMPinfo_StruSize As Integer = 8 + 2 + 2 + 1

        Dim struSize As Integer = 4 + 4 + BTSDK_BDADDR_LEN + BTSDK_DEVNAME_LEN + 4 + LMPinfo_StruSize + BTSDK_LINKKEY_LEN

        ReDim struBytes(0 To struSize - 1)

    End Sub

    Private Sub BlueSoleil_Stru_RemoteDeviceProp_GetInfo(ByRef bArray() As Byte, ByRef retDvcHandle As UInt32, ByRef retDvcAddr As String, ByRef retDvcName As String, ByRef retDvcClass As UInt32, ByRef lmpFeature As UInt64, ByRef lmpManuCode As UInt16, ByRef lmpSubVersion As UInt16, ByRef lmpVersion As Byte)

        Dim currByteIdx As Integer = 0

        currByteIdx = currByteIdx + 4

        retDvcHandle = BitConverter.ToUInt32(bArray, currByteIdx)
        currByteIdx = currByteIdx + 4

        Dim i As Integer
        For i = currByteIdx + 5 To currByteIdx Step -1
            retDvcAddr = retDvcAddr & Strings.Right("0" & Hex(bArray(currByteIdx)), 2)
            If i <> currByteIdx Then retDvcAddr = retDvcAddr & ":"
        Next i
        currByteIdx = currByteIdx + BTSDK_BDADDR_LEN

        retDvcName = System.Text.Encoding.UTF8.GetString(bArray, currByteIdx, BTSDK_DEVNAME_LEN)
        currByteIdx = currByteIdx + BTSDK_DEVNAME_LEN

        retDvcClass = BitConverter.ToUInt32(bArray, currByteIdx)
        currByteIdx = currByteIdx + 4

        'parse LMPinfo part...
        lmpFeature = BitConverter.ToUInt64(bArray, currByteIdx)
        currByteIdx = currByteIdx + 8

        lmpManuCode = BitConverter.ToUInt16(bArray, currByteIdx)
        currByteIdx = currByteIdx + 2

        lmpSubVersion = BitConverter.ToUInt16(bArray, currByteIdx)
        currByteIdx = currByteIdx + 2

        lmpVersion = bArray(currByteIdx)
        currByteIdx = currByteIdx + 1

        'link key?


    End Sub

    Public Function BlueSoleil_GetRemoteDeviceProperties(ByVal dvcHandle As UInt32, ByRef retDvcAddr As String, ByRef retDvcName As String, ByRef retDvcClass As UInt32, ByRef remoteDataAvailable As Boolean) As Boolean

        'struRemoteDeviceProp


        retDvcAddr = ""
        retDvcName = ""
        retDvcClass = 0
        remoteDataAvailable = False

        If dvcHandle = 0 Then Return False


        Dim struRemoteDevicePropBytes(0 To 0) As Byte
        BlueSoleil_Stru_RemoteDeviceProp_Init(struRemoteDevicePropBytes)

        Dim retUInt As UInt32 = Btsdk_GetRemoteDeviceProperty(dvcHandle, struRemoteDevicePropBytes(0))

        Dim retDvcHandle As UInt32
        Dim lmpFeature As UInt64, lmpManuCode As UInt16, lmpSubVersion As UInt16, lmpVersion As Byte

        BlueSoleil_Stru_RemoteDeviceProp_GetInfo(struRemoteDevicePropBytes, retDvcHandle, retDvcAddr, retDvcName, retDvcClass, lmpFeature, lmpManuCode, lmpSubVersion, lmpVersion)

        If lmpFeature <> 0 OrElse lmpManuCode <> 0 OrElse lmpSubVersion <> 0 OrElse lmpVersion <> 0 Then
            remoteDataAvailable = True
        End If

        Return (retUInt = BTSDK_OK)

    End Function



    Public Function BlueSoleil_SetDiscoveryMode(ByVal makeDiscoverable As Boolean, ByVal makePairable As Boolean, ByVal makeConnectable As Boolean) As Boolean

        Dim discFlags As UInt16 = 0

        If makeDiscoverable = True Then discFlags = discFlags Or BTSDK_DISCOVERABLE
        If makePairable = True Then discFlags = discFlags Or BTSDK_PAIRABLE
        If makeConnectable = True Then discFlags = discFlags Or BTSDK_CONNECTABLE

        Dim retUInt As UInt32 = Btsdk_SetDiscoveryMode(discFlags)

        Return (retUInt = BTSDK_OK)

    End Function

End Module

