'modBlueSoleil_HID - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'I am hoping to wrap these HID functions for completeness.  SDK documentation seems incomplete.
'

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices
Module modBlueSoleil_HID

    Private Const BTSDK_SERVICENAME_MAXLENGTH As Integer = 80

    Private Const BTSDK_TRUE As UInt32 = 1
    Private Const BTSDK_FALSE As UInt32 = 0

    Private Const BTSDK_DEVNAME_LEN As Integer = 64       '/* Shall Not be larger than MAX_NAME_LEN */
    Private Const BTSDK_SHORTCUT_NAME_LEN As Integer = 100
    Private Const BTSDK_BDADDR_LEN As Integer = 6
    Private Const BTSDK_LINKKEY_LEN As Integer = 16

    Private Const HidP_Input As Integer = 0
    Private Const HidP_Output As Integer = 1
    Private Const HidP_Feature As Integer = 2


    '/*HID APIS*/
    'BTINT32 Btsdk_CreateShortCutEx(PBtSdkShortCutPropertyStru shc_prop);
    'BTINT32 Btsdk_GetShortCutProperty(PBtSdkShortCutPropertyStru pshc_prop);
    'BTINT32 Btsdk_SetShortCutProperty(PBtSdkShortCutPropertyStru pshc_prop);


    'BTINT32 Btsdk_RecoverRemoteDeviceLinkKey(BTDEVHDL dev_hdl, BTUINT8* link_key);
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RecoverRemoteDeviceLinkKey(ByRef dvcHandle As UInt32, ByRef retLinkKey16Bytes As Byte) As UInt32
    End Function

    'BTUINT32 Btsdk_GetShortCutByDeviceHandle(BTDEVHDL dev_hdl, BTUINT16 service_class, BTSHCHDL *pshc_hdl, BTUINT32 max_shc_num);
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_GetShortCutByDeviceHandle(ByRef dvcHandle As UInt32, ByVal svcClass As UInt16, ByRef shortcutHandle As UInt32, ByVal maxShortcutNum As UInt32) As UInt32
    End Function

    'BTINT32 Btsdk_DeleteShortCut(BTSHCHDL shc_hdl);
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_DeleteShortCut(ByVal shortcutHandle As UInt32) As UInt32
    End Function

    'BTINT32 Btsdk_ConnectShortCut(BTSHCHDL shc_hdl);
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_ConnectShortCut(ByVal shortcutHandle As UInt32) As UInt32
    End Function

    'BTINT32 Btsdk_DisconnectShortCut(BTSHCHDL shc_hdl);
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_DisconnectShortCut(ByVal shortcutHandle As UInt32) As UInt32
    End Function



    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Hid_ClntGetPlugFlag(ByRef dvcAdrsBytes As Byte) As Byte        'returns BTSDK_TRUE or BTSDK_FALSE
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Hid_ClntGetReport(ByVal connHandle As UInt32, ByVal rptType As Byte, ByVal rptID As Byte, ByVal rptSize As UInt16, ByRef rptDataBytes As Byte) As UInt16
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Hid_ClntSetReport(ByVal connHandle As UInt32, ByVal rptType As Byte, ByVal rptID As Byte, ByVal rptSize As UInt16, ByRef rptDataBytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Hid_ClntPluggedDev(ByRef dvcAdrsBytes As Byte, ByRef struSDAPinfoBytes As Byte, ByRef struLocalInfoBytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Hid_ClntUnPluggedDev(ByRef dvcAdrsBytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Hid_LEClntUnPluggedDev(ByRef dvcAdrsBytes As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_Hid_LEHostConnect(ByVal dvcHandle As UInt32, ByRef retStruBTPHIDinfoBytes As Byte) As UInt32        'returns conn handle
    End Function


    Private Sub BlueSoleil_Stru_BTPsdapPNPinfo_Init(ByRef struBytes() As Byte)

        Dim struSize As Integer = 2 + 4 + 2 + 2 + 2 + 2 + 2 + 2     '18

        ReDim struBytes(0 To struSize - 1)

    End Sub

    Private Sub BlueSoleil_Stru_BTPsdapPNPinfo_GetInfo(ByRef struBytes() As Byte, ByVal struOffset As Integer)



    End Sub

    Private Sub BlueSoleil_Stru_BTPsdapPNPinfo_SetInfo(ByRef struBytes() As Byte, ByVal struOffset As Integer)



    End Sub


    Private Sub BlueSoleil_Stru_RmtHIDSvcExtAttr_Init(ByRef struBytes() As Byte)

        Dim struSize As Integer = 4 + 2 + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 2 + 1 + 2 + 1 ' + ?

        ReDim struBytes(0 To struSize - 1)

    End Sub


    Private Sub BlueSoleil_Stru_RmtHIDSvcExtAttr_GetInfo(ByRef struBytes() As Byte, ByVal struOffset As Integer)



    End Sub

    Private Sub BlueSoleil_Stru_RmtHIDSvcExtAttr_SetInfo(ByRef struBytes() As Byte, ByVal struOffset As Integer)



    End Sub


    Private Sub BlueSoleil_Stru_BTPhidHOSTinfo_Init(ByRef struBytes() As Byte)

        Dim struSize As Integer = 0
        struSize = struSize + 2 + 4 + 2 + 2 + 2 + 2 + 2 + 2
        struSize = struSize + 4 + 2 + 2 + 2 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 2 + 1 + 2 + 1 ' + ?

        ReDim struBytes(0 To struSize - 1)

        ReDim struBytes(0 To struSize + 1000)

    End Sub


    Private Sub BlueSoleil_Stru_ShortCutProperty_Init(ByRef struBytes() As Byte)

        '    typedef struct  _BtSdkShortCutPropertyStru
        '{
        '	BTSHCHDL shc_hdl;						/* handle assigned To the shortcut instance*/
        '	BTCONNHDL conn_hdl;						/* handle assigned To the connection instance*/
        '	BTCONNHDL dev_hdl;						/* handle assigned To the device instance */
        '	BTSVCHDL svc_hdl;						/* handle assigned To the service instance */
        '	BTBOOL	by_dev_hdl;						/* BTSDK_TRUE: Specify device by dev_hdl. 
        '	Otherwise by bd_addr. */
        '	BTBOOL	by_svc_hdl;						/* BTSDK_TRUE: Specify service type by svc_hdl. 
        '	Otherwise by svc_class. */
        '	BTUINT8 bd_addr[BTSDK_BDADDR_LEN];		/* bt address Of local device */
        '	BTUINT16 svc_class;						/* service Class */
        '	BTUINT16 mask;							/* Specified which member Is To be Set Or Get. */ 
        '	BTUINT8 shc_name[BTSDK_SHORTCUT_NAME_LEN];	/* name Of the shortcut, must In UTF-8 */
        '	BTUINT8 dev_name[BTSDK_DEVNAME_LEN];	/* Name Of the device record, must be In UTF-8 */
        '	BTUINT8 svc_name[BTSDK_SERVICENAME_MAXLENGTH];	/*must In UTF-8*/
        '	BTUINT32 dev_class;						/* device Class Of the remote device*/
        '	BTBOOL is_default;						/* Is Default shortcut */
        '	BTUINT8 sec_level;						/* Security level Of this shortcut. Authentication/Encryption. */
        '	BTUINT16 shc_attrib_len;				/* the length Of shortcut attribute */
        '	BTUINT8 *pshc_attrib;					/* shortcut attribute */
        '}*PBtSdkShortCutPropertyStru, BtSdkShortCutPropertyStru;

        Dim struSize As Integer = 4 + 4 + 4 + 4 + 1 + 1 + BTSDK_BDADDR_LEN + 2 + 2 + BTSDK_SHORTCUT_NAME_LEN + BTSDK_DEVNAME_LEN + BTSDK_SERVICENAME_MAXLENGTH + 4 + 1 + 1 + 2 + 1 '+ ?



    End Sub


    Public Function BlueSoleil_HID_IsDevicePresent(ByVal dvcAddress As String) As Boolean

        If dvcAddress = "" Then Return False

        Dim addrHexBytes(0 To 0) As String
        addrHexBytes = Split(dvcAddress, ":")

        If addrHexBytes.Length <> 6 Then

            Return False
        End If

        Dim dvcAddrBytes(0 To 5) As Byte

        Dim i As Integer
        Dim bCounter As Integer = 0
        For i = 5 To 0 Step -1
            dvcAddrBytes(bCounter) = CByte(Val("&H" & addrHexBytes(i)))
        Next

        Dim retByte As Byte = Btsdk_Hid_ClntGetPlugFlag(dvcAddrBytes(0))

        Return retByte = BTSDK_TRUE

    End Function


    Public Function BlueSoleil_HID_GetReport_Input(ByVal dvcAddress As String, Optional ByVal reportID As Byte = 0) As UInt32



    End Function


    Public Function BlueSoleil_HID_ConnectDevice(ByVal dvcHandle As UInt32) As UInt32

        If dvcHandle = 0 Then Return 0

        Dim struBTPhidHOSTinfo(0 To 0) As Byte
        BlueSoleil_Stru_BTPhidHOSTinfo_Init(struBTPhidHOSTinfo)

        Dim retUInt As UInt32 = Btsdk_Hid_LEHostConnect(dvcHandle, struBTPhidHOSTinfo(0))

        'parse struBTPhidHOSTinfo

        Return retUInt

    End Function



End Module
