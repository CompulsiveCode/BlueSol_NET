'modBlueSoleil_MAP - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'This module wraps the Blue Soleil SDK functions for using the Message Access Profile.
'
'This module wraps the Win32 File IO routines so BlueSoleil can access them in a platform-agnostic way.  See the BSFileIO function(s).
'
'Apparently the iPhone does not support sending messages via the MAP profile.
'
'


Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices
Module modBlueSoleil_MAP

    Private BlueSoleil_MAP_TempFileCounter As Long = 0

    Private BlueSoleil_MAP_LastCreateFileMsgNotification As DateTime = Now

    Private Const BTSDK_OK As UInt32 = 0
    Private Const BTSDK_TRUE As Byte = 1

    Private Const BTSDK_SERVICENAME_MAXLENGTH As UInt32 = 80

    Private Const BTSDK_PBAP_MAX_DELIMITER As Byte = &H02

    Private Const BTSDK_MAX_SUPPORT_FORMAT As UInt16 = 6       '/* OPP format number */
    Private Const BTSDK_PATH_MAXLENGTH As UInt16 = 256     '/* Shall not be larger than FTP_MAX_PATH and OPP_MAX_PATH */
    Private Const BTSDK_CARDNAME_MAXLENGTH As UInt16 = 256     '/* Shall not be larger than OPP_MAX_NAME */

    Private Const BTSDK_MAP_SUP_MSG_EMAIL As Byte = 01
    Private Const BTSDK_MAP_SUP_MSG_SMSGSM As Byte = 02
    Private Const BTSDK_MAP_SUP_MSG_SMSCDMA As Byte = 04
    Private Const BTSDK_MAP_SUP_MSG_MMS As Byte = 08

    Private Const BTSDK_MAP_GFLP_MAXCOUNT As UInt16 = &H0001
    Private Const BTSDK_MAP_GFLP_STARTOFF As UInt16 = &H0002
    Private Const BTSDK_MAP_GFLP_LISTSIZE As UInt16 = &H0004

    Private Const BTSDK_MAP_PATH_LEN As Integer = 512
    Private Const BTSDK_MAP_FOLDER_LEN As Integer = 32
    Private Const BTSDK_MAP_TIME_LEN As Integer = 20
    Private Const BTSDK_MAP_MSE_TIME_LEN As Integer = 24
    Private Const BTSDK_MAP_MSGHDL_LEN As Integer = 20
    Private Const BTSDK_MAP_MSGTYPE_LEN As Integer = 16
    Private Const BTSDK_MAP_SUBJECT_LEN As Integer = 256
    Private Const BTSDK_MAP_USERNAME_LEN As Integer = 256
    Private Const BTSDK_MAP_ADDR_LEN As Integer = 256

    Private Const BTSDK_MAP_GMLP_MAXCOUNT As UInt32 = &H00000001
    Private Const BTSDK_MAP_GMLP_STARTOFF As UInt32 = &H00000002
    Private Const BTSDK_MAP_GMLP_MSGTYPE As UInt32 = &H00000004
    Private Const BTSDK_MAP_GMLP_PERIODBEGIN As UInt32 = &H00000008
    Private Const BTSDK_MAP_GMLP_PERIODEND As UInt32 = &H00000010
    Private Const BTSDK_MAP_GMLP_READSTATUS As UInt32 = &H00000020
    Private Const BTSDK_MAP_GMLP_RECIPIENT As UInt32 = &H00000040
    Private Const BTSDK_MAP_GMLP_ORIGINATOR As UInt32 = &H00000080
    Private Const BTSDK_MAP_GMLP_PRIORITY As UInt32 = &H00000100
    Private Const BTSDK_MAP_GMLP_NEWMSG As UInt32 = &H00001000
    Private Const BTSDK_MAP_GMLP_PARAMMASK As UInt32 = &H00008000
    Private Const BTSDK_MAP_GMLP_LISTSIZE As UInt32 = &H00020000
    Private Const BTSDK_MAP_GMLP_SUBJECTLENTH As UInt32 = &H00040000
    Private Const BTSDK_MAP_GMLP_MSETIME As UInt32 = &H01000000

    Private Const BTSDK_MAP_MP_SUBJECT As UInt16 = &H0001
    Private Const BTSDK_MAP_MP_DATATIME As UInt16 = &H0002
    Private Const BTSDK_MAP_MP_SENDERNAME As UInt16 = &H0004
    Private Const BTSDK_MAP_MP_SENDERADDR As UInt16 = &H0008
    Private Const BTSDK_MAP_MP_RECIPIENTNAME As UInt16 = &H0010
    Private Const BTSDK_MAP_MP_RECIPIENTADDR As UInt16 = &H0020
    Private Const BTSDK_MAP_MP_TYPE As UInt16 = &H0040
    Private Const BTSDK_MAP_MP_SIZE As UInt16 = &H0080
    Private Const BTSDK_MAP_MP_RECPSTATUS As UInt16 = &H0100
    Private Const BTSDK_MAP_MP_TEXT As UInt16 = &H0200
    Private Const BTSDK_MAP_MP_ATTACHSIZE As UInt16 = &H0400
    Private Const BTSDK_MAP_MP_PRIORITY As UInt16 = &H0800
    Private Const BTSDK_MAP_MP_READ As UInt16 = &H1000
    Private Const BTSDK_MAP_MP_SENT As UInt16 = &H2000
    Private Const BTSDK_MAP_MP_PROTECTED As UInt16 = &H4000
    Private Const BTSDK_MAP_MP_REPLY2ADDR As UInt16 = &H8000


    Private Const BTSDK_MAP_MSG_FILTER_ST_ALL As Byte = &H00
    Private Const BTSDK_MAP_MSG_FILTER_ST_UNREAD As Byte = &H01
    Private Const BTSDK_MAP_MSG_FILTER_ST_READ As Byte = &H02

    Private Const BTSDK_MAP_FILTEROUT_NO As Byte = &H00
    Private Const BTSDK_MAP_FILTEROUT_SMSGSM As Byte = &H01
    Private Const BTSDK_MAP_FILTEROUT_SMSCDMA As Byte = &H02
    Private Const BTSDK_MAP_FILTEROUT_EMAIL As Byte = &H03
    Private Const BTSDK_MAP_FILTEROUT_MMS As Byte = &H04


    Private Const BTSDK_MAP_CHARSET_NATIVE As Byte = &H00
    Private Const BTSDK_MAP_CHARSET_UTF8 As Byte = &H01

    '/* Fraction requirement - possible values of BtSdkMAPGetMsgParamStrufraction_req */
    Private Const BTSDK_MAP_FRACT_NONE As Byte = &H00
    Private Const BTSDK_MAP_FRACT_REQFIRST As Byte = &H01
    Private Const BTSDK_MAP_FRACT_REQNEXT As Byte = &H02

    '/* Fraction indication - possible values of BtSdkMAPGetMsgParamStrufraction_deliver */
    Private Const BTSDK_MAP_FRACT_RSPMORE As Byte = &H00
    Private Const BTSDK_MAP_FRACT_RSPLAST As Byte = &H01


    '/* Message status indicator value - possible values of Btsdk_MAPSetMessageStatusstatus */
    Private Const BTSDK_MAP_MSG_SETST_READ As Byte = &H02
    Private Const BTSDK_MAP_MSG_SETST_UNREAD As Byte = &H00
    Private Const BTSDK_MAP_MSG_SETST_DELETED As Byte = &H03
    Private Const BTSDK_MAP_MSG_SETST_UNDELETED As Byte = &H01


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




    Private BlueSoleil_MAP_Callback_Status_CurrSize As UInt32 = 0
    Private BlueSoleil_MAP_Callback_Status_TotalSize As UInt32 = 0

    Private BSfileIO_APP_CurrentDir As String = ""
    Private BSfileIO_APP_RootBTdir As String = ""

    Private BSfileIO_APP_ConnHandle As UInt32 = 0


    Public Event BlueSoleil_Event_MAP_MsgNotification()

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncMAPstatusCallback(ByVal first As Byte, ByVal last As Byte, ByVal ptrFileNameBytes As IntPtr, ByVal fileSize As UInt32, ByVal curSize As UInt32)
    Public delegateMAPstatusCallback As delfuncMAPstatusCallback = AddressOf BlueSoleil_MAP_Callback_Status

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfunMAPmsgNotification(ByVal svcHandle As UInt32, ByVal EvReportPtr As IntPtr)
    Public delegateMsgNotification As delfunMAPmsgNotification = AddressOf BlueSoleil_MAP_Callback_MessageNotification


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPOpenFile(ByVal arrayPtr As IntPtr) As UInt32
    Public delegateAPPOpenFile As delfuncAPPOpenFile = AddressOf BSfileIO_APP_OpenFile

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPCreateFile(ByVal arrayPtr As IntPtr) As UInt32
    Public delegateAPPCreateFile As delfuncAPPCreateFile = AddressOf BSfileIO_APP_CreateFile

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPWriteFile(ByVal fHandle As UInt32, ByVal arrayPtr As IntPtr, ByVal arrayLen As UInt32) As UInt32
    Public delegateAPPWriteFile As delfuncAPPWriteFile = AddressOf BSfileIO_APP_WriteFile

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPReadFile(ByVal fHandle As UInt32, ByVal arrayPtr As IntPtr, ByVal arrayLen As UInt32, ByVal retIsEOF As IntPtr) As UInt32
    Public delegateAPPReadFile As delfuncAPPReadFile = AddressOf BSfileIO_APP_ReadFile

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPGetFileSize(ByVal fHandle As UInt32) As UInt32
    Public delegateAPPGetFileSize As delfuncAPPGetFileSize = AddressOf BSfileIO_APP_GetFileSize

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPRewindFile(ByVal fHandle As UInt32, ByVal offset As UInt32) As Int32
    Public delegateAPPRewindFile As delfuncAPPRewindFile = AddressOf BSfileIO_APP_Rewind

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncAPPCloseFile(ByVal fHandle As UInt32)
    Public delegateAPPCloseFile As delfuncAPPCloseFile = AddressOf BSfileIO_APP_CloseFile


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPFindFolderFirst(ByVal ptrFolderNameBytes As IntPtr, ByVal ptrFolderObjStru As IntPtr) As UInt32
    Public delegateAPPFindFolderFirst As delfuncAPPFindFolderFirst = AddressOf BSfileIO_APP_FindFolder_First

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPFindFolderNext(ByVal hFind As UInt32, ByVal ptrFolderObjStru As IntPtr) As Byte
    Public delegateAPPFindFolderNext As delfuncAPPFindFolderNext = AddressOf BSfileIO_APP_FindFolder_Next

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncAPPFindFolderClose(ByVal fHandle As UInt32)
    Public delegateAPPFindFolderClose As delfuncAPPFindFolderClose = AddressOf BSfileIO_APP_FindFolder_Close


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPFindMessageFirst(ByVal ptrFilePathBytes As IntPtr, ByVal ptrMessageFilterStru As IntPtr, ByVal ptrMessageObjStru As IntPtr) As UInt32
    Public delegateAPPFindMessageFirst As delfuncAPPFindMessageFirst = AddressOf BSfileIO_APP_FindMessage_First

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPFindMessageNext(ByVal hFind As UInt32, ByVal ptrMessageFilterStru As IntPtr, ByVal ptrFolderObjStru As IntPtr) As Byte
    Public delegateAPPFindMessageNext As delfuncAPPFindMessageNext = AddressOf BSfileIO_APP_FindMessage_Next

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfuncAPPFindMessageClose(ByVal fHandle As UInt32)
    Public delegateAPPFindMessageClose As delfuncAPPFindMessageClose = AddressOf BSfileIO_APP_FindMessage_Close


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPCreateBMessageFile(ByVal ptrStru_MsgHandle As IntPtr) As Byte
    Public delegateAPPCreateBMessageFile As delfuncAPPCreateBMessageFile = AddressOf BSmsgIO_APP_CreateBMessageFile

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPOpenBMessageFile(ByVal ptrStru_GetMsgParam As IntPtr) As UInt32
    Public delegateAPPOpenBMessageFile As delfuncAPPOpenBMessageFile = AddressOf BSmsgIO_APP_OpenBMessageFile

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPPushMessage(ByVal ptrFilePathBytes As IntPtr, ByVal ptrStru_PushMsgParam As IntPtr) As UInt32
    Public delegateAPPPushMessage As delfuncAPPPushMessage = AddressOf BSmsgIO_APP_PushMessage

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPModifyMessageStatus(ByVal ptrStru_MsgStatus As IntPtr) As UInt32
    Public delegateAPPModifyMessageStatus As delfuncAPPModifyMessageStatus = AddressOf BSmsgIO_APP_ModifyMsgStatus


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPRegisterNotification(ByVal connHandle As UInt32, ByVal svcHandle As UInt32, ByVal turnOn As Byte) As Byte
    Public delegateAPPRegisterNotification As delfuncAPPRegisterNotification = AddressOf BSmsgIO_APP_RegisterNotification

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPUpdateInbox() As Byte
    Public delegateAPPUpdateInbox As delfuncAPPUpdateInbox = AddressOf BSmsgIO_APP_UpdateInbox

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function delfuncAPPGetMSETime(ByVal ptrStru_MSETime As IntPtr) As Byte
    Public delegateAPPGetMSETime As delfuncAPPGetMSETime = AddressOf BSmsgIO_APP_GetMSETime











    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPRegisterFileIORoutines(ByVal connHandle As UInt32, ByRef mapFileIORoutinesStru As Byte) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RegisterMASService(ByRef localSvcNameBytes As Byte, ByRef masServerAttribStru As Byte, ByRef masServerCBStru As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RegisterMNSService(ByRef localSvcNameBytes As Byte, ByVal functPtr_MNSMessageNotification As UInt32, ByRef mapFileIORoutinesStru As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_UnregisterMAPService(ByVal svcHandle As UInt32) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPRegisterSvrCallback(ByVal svcHandle As UInt32, ByRef masServerCBStru As Byte) As UInt32
    End Function



    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPSetNotificationRegistration(ByVal connHandle As UInt32, ByVal turnON As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPUpdateInbox(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPCancelTransfer(ByVal connHandle As UInt32) As UInt32
    End Function



    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPSetFolder(ByVal connHandle As UInt32, ByRef folderName As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPGetFolderList(ByVal connHandle As UInt32, ByRef mapGetFolderListParamStru As Byte, ByVal fHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPGetMessageList(ByVal connHandle As UInt32, ByRef mapGetMsgListParamStru As Byte, ByVal fHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPGetMessage(ByVal connHandle As UInt32, ByRef mapGetMsgParamStru As Byte, ByVal fHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPRegisterStatusCallback(ByVal connHandle As UInt32, ByVal functPtr_MAP_STATUS_INFO_CB As IntPtr) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPPushMessage(ByVal connHandle As UInt32, ByRef mapPushMsgParamStru As Byte, ByVal fHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_MAPSetMessageStatus(ByVal connHandle As UInt32, ByRef mapMsgHandleBytes As Byte, ByVal newStatus As Byte) As UInt32
    End Function

    'Btsdk_MAPRegisterStatusCallback


    'BTSDKHANDLE Btsdk_MAPStartEnumFolderList(Btsdk_ReadFile_Func func_read, Btsdk_RewindFile_Func func_rewind, BTSDKHANDLE file_hdl);
    'PBtSdkMAPFolderObjStru Btsdk_MAPEnumFolderList(BTSDKHANDLE enum_hdl, PBtSdkMAPFolderObjStru item);
    'void Btsdk_MAPEndEnumFolderList(BTSDKHANDLE enum_hdl);
    'BTINT32 Btsdk_MAPSetMessageStatus(BTCONNHDL conn_hdl, BTUINT8 *msg_hdl, BTUINT8 status);

    'BTSDKHANDLE Btsdk_MAPStartEnumMessageList(Btsdk_ReadFile_Func func_read, Btsdk_RewindFile_Func func_rewind, BTSDKHANDLE file_hdl);
    'PBtSdkMAPMsgObjStru Btsdk_MAPEnumMessageList(BTSDKHANDLE enum_hdl, PBtSdkMAPMsgObjStru item);
    'void Btsdk_MAPEndEnumMessageList(BTSDKHANDLE enum_hdl);



    Private Function BSfileIO_APP_CreateDir(ByVal arrayPtr As IntPtr) As Integer

        Debug.Print("BSfileIO_CreateDir")

        Dim tempDirName As String = ""
        tempDirName = System.Runtime.InteropServices.Marshal.PtrToStringAuto(arrayPtr)

        Dim TorF As Boolean = False

        Try
            IO.Directory.CreateDirectory(tempDirName)
            TorF = True
        Catch ex As Exception

        End Try


        If TorF = True Then
            Return 0
        Else

            Try
                If IO.Directory.Exists(tempDirName) = True Then
                    Return 0
                Else
                    Return -1
                End If
            Catch ex As Exception
                Return -1
            End Try

        End If

    End Function

    Private Function BSfileIO_APP_OpenFile(ByVal arrayPtr As IntPtr) As UInt32

        Debug.Print("BSfileIO_OpenFile")

        If arrayPtr = IntPtr.Zero Then
            Return 0
        End If

        Dim tempDirName As String = ""
        tempDirName = System.Runtime.InteropServices.Marshal.PtrToStringAuto(arrayPtr)

        Dim hFile As IntPtr = FileAPI_OpenFile(tempDirName, False)

        Return CUInt(hFile)


    End Function

    Private Function BSfileIO_APP_CreateFile(ByVal arrayPtr As IntPtr) As UInt32

        'this function seems to be fired a few times by the MNS service when a message is received.



        Debug.Print("BSfileIO_CreateFile")

        Dim tempFN As String = ""

        Dim hFile As IntPtr = IntPtr.Zero

        If arrayPtr = IntPtr.Zero Then
            'create temp file and return that.
            tempFN = System.IO.Path.GetTempFileName
            'tempFN = "C:\My Work\Media\Car Stuff\DriveLine\BlueSoleilTest\bin\Debug\MsgServer\" & BlueSoleil_MAP_TempFileCounter & ".txt"
            'BlueSoleil_MAP_TempFileCounter = BlueSoleil_MAP_TempFileCounter + 1
            '   Try
            '       IO.File.Delete(tempFN)
            '   Catch ex As Exception
            '
            '   End Try

            If Now.Subtract(BlueSoleil_MAP_LastCreateFileMsgNotification).TotalSeconds > 10 Then
                BlueSoleil_MAP_LastCreateFileMsgNotification = Now
                'RaiseEvent BlueSoleil_Event_MAP_MsgNotification()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_MAP_MsgNotification())
                t.Start()
            End If
            BlueSoleil_MAP_LastCreateFileMsgNotification = Now

            hFile = FileAPI_OpenFile(tempFN, False)
        Else
            tempFN = System.Runtime.InteropServices.Marshal.PtrToStringAuto(arrayPtr)
            hFile = FileAPI_OpenFile(tempFN, True)
        End If



        Debug.Print("hFile = " & hFile.ToInt64)

        Return CUInt(hFile)


    End Function

    Private Function BSfileIO_APP_WriteFile(ByVal fHandle As UInt32, ByVal arrayPtr As IntPtr, ByVal arrayLen As UInt32) As UInt32

        Debug.Print("BSfileIO_WriteFile Len = " & arrayLen & "  Ptr = " & arrayPtr.ToInt64 & "  handle = " & fHandle)

        If fHandle = 0 Then
            Return 0
        End If

        If arrayLen = 0 Then        'this is important.
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



    Private Function BSfileIO_APP_ReadFile(ByVal fHandle As UInt32, ByVal arrayPtr As IntPtr, ByVal arrayLen As UInt32, ByVal retIsEOF As IntPtr) As UInt32

        Debug.Print("BSfileIO_ReadFile Len=" & arrayLen)

        If fHandle = 0 Then
            Return 0
        End If

        If arrayLen = 0 Then
            Return 1            '
        End If

        Dim tempBytesArray(0 To 0) As Byte
        ReDim tempBytesArray(0 To CInt(arrayLen - 1))


        Dim TorF As Boolean = FileAPI_GetBytes(CType(fHandle, IntPtr), -1, tempBytesArray.Length, tempBytesArray)





        If TorF = True Then 'put the bytes back to the pointer.
            System.Runtime.InteropServices.Marshal.Copy(tempBytesArray, 0, arrayPtr, tempBytesArray.Length)
        End If

        Dim newEOFval As Byte = 0
        If FileAPI_GetCurrentOffset(CType(fHandle, IntPtr)) >= FileAPI_GetFileSize(CType(fHandle, IntPtr)) Then
            newEOFval = 1
        End If
        If tempBytesArray.Length < arrayLen Then
            newEOFval = 1
        End If
        System.Runtime.InteropServices.Marshal.WriteByte(retIsEOF, newEOFval)


        If TorF = True Then
            Debug.Print("Return = " & tempBytesArray.Length)
            Return CUInt(tempBytesArray.Length)
            'Return 1
        Else
            Debug.Print("Return = 0")
            Return 0
        End If

    End Function



    Private Sub BSfileIO_APP_CloseFile(ByVal fHandle As UInt32)

        Debug.Print("BSfileIO_CloseFile " & fHandle)

        Try
            FileAPI_CloseFile(CType(fHandle, IntPtr))
        Catch ex As Exception

        End Try

    End Sub


    Private Function BSfileIO_APP_GetFileSize(ByVal fHandle As UInt32) As UInt32

        Debug.Print("BSfileIO_GetFileSize")

        Dim retLong As Long = FileAPI_GetFileSize(CType(fHandle, IntPtr))

        If retLong < 1 Then retLong = 0

        Debug.Print("Return = " & retLong)

        Return CUInt(retLong)

    End Function

    Private Function BSfileIO_APP_Rewind(ByVal fHandle As UInt32, ByVal offset As UInt32) As Int32

        Debug.Print("BSfileIO_Rewind - offset = " & offset)

        Dim TorF As Boolean = FileAPI_SetFileOffset(CType(fHandle, IntPtr), offset)

        Debug.Print("Return = " & TorF)
        If TorF = True Then
            Return 0
        Else
            Return -1
        End If

    End Function


    Private Function BSfileIO_APP_FindMessage_First(ByVal arrayPtr As IntPtr, ByVal ptrMessageFilterObjStru As IntPtr, ByVal ptrMessageObjStru As IntPtr) As UInt32

        Debug.Print("BSfileIO_APP_FindMessage_First")

        Dim tempDirName As String = ""
        tempDirName = System.Runtime.InteropServices.Marshal.PtrToStringAuto(arrayPtr)



        If Strings.Right(tempDirName, 1) <> "\" Then tempDirName = tempDirName & "\"
        tempDirName = tempDirName & "*.msg"

        Dim retFileName As String = ""
        Dim retIsDir As Boolean = False
        Dim hFind As IntPtr = FindFileAPI_FindFirst(tempDirName, retFileName, retIsDir)
        Dim findTorF As Boolean = (hFind <> IntPtr.Zero)
        Dim fullFN As String = ""

        Do
            fullFN = tempDirName
            If Strings.Right(fullFN, 1) <> "\" Then fullFN = fullFN & "\"
            fullFN = fullFN & retFileName

            If retFileName <> "." And retFileName <> ".." Then
                If retIsDir = False Then
                    findTorF = True
                    Exit Do
                End If
            End If

            findTorF = FindFileAPI_FindNext(hFind, retFileName, retIsDir)

            If findTorF = False Then
                FindFileAPI_CloseHandle(CType(hFind, IntPtr))
                Exit Do
            End If

        Loop

        If findTorF = False Then
            Return 0
        End If

        'populate ptrFolderObjStru

        Dim struMessageObj(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapMessageObj(struMessageObj)
        Marshal.Copy(struMessageObj, 0, ptrMessageObjStru, struMessageObj.Length)

        Return CUInt(hFind)

    End Function




    Private Function BSfileIO_APP_FindMessage_Next(ByVal hFind As UInt32, ByVal ptrMessageFilterObjStru As IntPtr, ByVal ptrMessageObjStru As IntPtr) As Byte

        Debug.Print("BSfileIO_APP_FindMessage_Next")


        Dim retFileName As String = ""
        Dim retIsDir As Boolean = False

        Dim findTorF As Boolean = (hFind <> 0)
        findTorF = FindFileAPI_FindNext(CType(hFind, IntPtr), retFileName, retIsDir)


        Do

            If retFileName <> "." And retFileName <> ".." Then
                If retIsDir = False Then
                    findTorF = True
                    Exit Do
                End If
            End If



            If findTorF = False Then
                FindFileAPI_CloseHandle(CType(hFind, IntPtr))
                Exit Do
            End If

        Loop

        If findTorF = False Then
            Return 0
        End If

        'populate ptrFolderObjStru

        Dim struMessageObj(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapMessageObj(struMessageObj)
        Marshal.Copy(struMessageObj, 0, ptrMessageObjStru, struMessageObj.Length)

        Return 1


    End Function


    Private Function BSfileIO_APP_FindFolder_First(ByVal arrayPtr As IntPtr, ByVal ptrFolderObjStru As IntPtr) As UInt32

        Debug.Print("BSfileIO_APP_FindFolder_First")

        Dim tempDirName As String = ""
        tempDirName = System.Runtime.InteropServices.Marshal.PtrToStringAuto(arrayPtr)



        If Strings.Right(tempDirName, 1) <> "\" Then tempDirName = tempDirName & "\"
        tempDirName = tempDirName & "*.*"

        Dim retFileName As String = ""
        Dim retIsDir As Boolean = False
        Dim hFind As IntPtr = FindFileAPI_FindFirst(tempDirName, retFileName, retIsDir)
        Dim findTorF As Boolean = (hFind <> IntPtr.Zero)
        Dim fullFN As String = ""

        Do
            fullFN = tempDirName
            If Strings.Right(fullFN, 1) <> "\" Then fullFN = fullFN & "\"
            fullFN = fullFN & retFileName

            If retFileName <> "." And retFileName <> ".." Then
                If retIsDir = True Then
                    findTorF = True
                    Exit Do
                End If
            End If

            findTorF = FindFileAPI_FindNext(hFind, retFileName, retIsDir)

            If findTorF = False Then
                FindFileAPI_CloseHandle(CType(hFind, IntPtr))
                Exit Do
            End If

        Loop

        If findTorF = False Then
            Return 0
        End If

        'populate ptrFolderObjStru

        Dim struFolderObj(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapFolderObj(struFolderObj, retFileName)
        Marshal.Copy(struFolderObj, 0, ptrFolderObjStru, struFolderObj.Length)

        Return CUInt(hFind)

    End Function



    Private Function BSfileIO_APP_FindFolder_Next(ByVal hFind As UInt32, ByVal ptrFolderObjStru As IntPtr) As Byte

        Debug.Print("BSfileIO_APP_FindFolder_Next")





        Dim retFileName As String = ""
        Dim retIsDir As Boolean = False
        Dim findTorF As Boolean = FindFileAPI_FindNext(CType(hFind, IntPtr), retFileName, retIsDir)
        Dim fullFN As String = ""

        Do

            If retFileName <> "." And retFileName <> ".." Then
                If retIsDir = True Then
                    findTorF = True
                    Exit Do
                End If
            End If

            findTorF = FindFileAPI_FindNext(CType(hFind, IntPtr), retFileName, retIsDir)

            If findTorF = False Then Exit Do

        Loop

        If findTorF = False Then
            Return 0
        End If

        'populate ptrFolderObjStru

        Dim struFolderObj(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapFolderObj(struFolderObj, retFileName)
        Marshal.Copy(struFolderObj, 0, ptrFolderObjStru, struFolderObj.Length)

        Return 1

    End Function


    Private Sub BSfileIO_APP_FindFolder_Close(ByVal hFind As UInt32)

        Debug.Print("BSfileIO_APP_FindFolder_Close")

        FindFileAPI_CloseHandle(CType(hFind, IntPtr))

    End Sub

    Private Sub BSfileIO_APP_FindMessage_Close(ByVal hFind As UInt32)

        Debug.Print("BSfileIO_APP_FindMessage_Close")

        FindFileAPI_CloseHandle(CType(hFind, IntPtr))

    End Sub

    Private Function BSfileIO_APP_ChangeDir(ByVal arrayPtr As IntPtr) As Integer

        Debug.Print("BSfileIO_ChangeDir")

        Dim tempDirName As String = ""
        tempDirName = System.Runtime.InteropServices.Marshal.PtrToStringAuto(arrayPtr)

        Try
            IO.Directory.SetCurrentDirectory(tempDirName)

        Catch ex As Exception

        End Try

        BSfileIO_APP_CurrentDir = tempDirName

        Try
            If IO.Directory.Exists(tempDirName) = True Then
                Return 0
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try




    End Function

    Private Function BSmsgIO_APP_CreateBMessageFile(ByVal ptrStru_MsgHandle As IntPtr) As Byte

        Debug.Print("BSmsgIO_APP_CreateBMessageFile")

        Return 1

    End Function


    Private Function BSmsgIO_APP_OpenBMessageFile(ByVal ptrStru_GetMsgParam As IntPtr) As UInt32

        Debug.Print("BSmsgIO_APP_OpenBMessageFile")

        'get msg handle from structure.
        'find msghandle.msg in BSfileIO_APP_RootBTdir


        'return the open file handle.
        Return 0

    End Function

    Private Function BSmsgIO_APP_PushMessage(ByVal arrayPtr_CurPath As IntPtr, ByVal ptrStru_PushMsgParam As IntPtr) As Byte

        Debug.Print("BSmsgIO_APP_PushMessage")

        Return 1

    End Function

    Private Function BSmsgIO_APP_ModifyMsgStatus(ByVal ptrStru_MsgStatus As IntPtr) As Byte

        Debug.Print("BSmsgIO_APP_ModifyMsgStatus")


        Return 1

    End Function

    Private Sub BlueSoleil_MAP_InitStruBytes_masLocalServerAttrStru(ByRef inpByteArray() As Byte, ByVal mapRootPathLocal As String)

        Dim sizeOfStru As Integer = 4 + 2 + (BTSDK_PATH_MAXLENGTH + 1) + (BTSDK_PBAP_MAX_DELIMITER + 1) + 1 + 1
        ReDim inpByteArray(0 To sizeOfStru - 1)

        Dim tempBytes(0 To 3) As Byte

        Dim currByteIdx As Integer = 0

        'structure length.
        tempBytes = BitConverter.GetBytes(sizeOfStru)
        Array.Copy(tempBytes, 0, inpByteArray, currByteIdx, tempBytes.Length)
        currByteIdx = currByteIdx + tempBytes.Length

        'mask?  reserved.
        currByteIdx = currByteIdx + 2

        'root path
        Dim pathBytes(0 To Len(mapRootPathLocal)) As Byte
        pathBytes = System.Text.UTF8Encoding.UTF8.GetBytes(mapRootPathLocal & Chr(0))
        Array.Copy(pathBytes, 0, inpByteArray, currByteIdx, pathBytes.Length)
        currByteIdx = currByteIdx + (BTSDK_PATH_MAXLENGTH + 1)

        'path delimiter
        inpByteArray(currByteIdx) = CByte(Asc("\"))
        currByteIdx = currByteIdx + (BTSDK_PBAP_MAX_DELIMITER + 1)

        'instance ID
        inpByteArray(currByteIdx) = 0
        currByteIdx = currByteIdx + 1

        'supported msg types.
        Dim msgTypes As Byte = BTSDK_MAP_SUP_MSG_EMAIL Or BTSDK_MAP_SUP_MSG_MMS Or BTSDK_MAP_SUP_MSG_SMSCDMA Or BTSDK_MAP_SUP_MSG_SMSGSM
        inpByteArray(currByteIdx) = msgTypes


    End Sub



    Private Sub BlueSoleil_MAP_InitStruBytes_masServerCBstru(ByRef inpByteArray() As Byte, ByRef struFindFolderRoutines() As Byte, ByRef struFindMsgRoutines() As Byte, ByRef struFileIOroutines() As Byte, ByRef struMsgIOroutines() As Byte, ByRef struMsgStatusRoutines() As Byte)


        '	BtSdkMAPFindFolderRoutinesStru  find_folder_rtns;           ;4 x 3
        '       Btsdk_MAP_FindFirstFolder_Func  find_first_folder;
        '       Btsdk_MAP_FindNextFolder_Func   find_next_folder;
        '       Btsdk_MAP_FindCloseFolder_Func  find_folder_close;

        '   BtSdkMAPFindMsgRoutinesStru     find_msg_rtns;              ;4 x 3
        '       Btsdk_MAP_FindFirstMessage_Func  find_first_folder;
        '       Btsdk_MAP_FindNextMessage_Func   find_next_folder;
        '       Btsdk_MAP_FindCloseMessage_Func  find_folder_close;

        '   BtSdkMAPFileIORoutinesStru      file_io_rtns;               ;4 x 7
        '       Btsdk_OpenFile_Func         open_file;
        '       Btsdk_CreateFile_Func       create_file;
        '       Btsdk_WriteFile_Func        write_file;
        '       Btsdk_ReadFile_Func         read_file;
        '       Btsdk_GetFileSize_Func  get_file_size;
        '       Btsdk_RewindFile_Func   rewind_file;
        '       Btsdk_CloseFile_Func        close_file;

        '   BtSdkMAPMsgIORoutinesStru       msg_io_rtns;                ;4 x 4
        '       Btsdk_ModifyMsgStatus_Func  modify_msg_status;
        '       Btsdk_CreateBMsgFile_Func   create_bmsg_file;
        '       Btsdk_OpenBMsgFile_Func     open_bmsg_file;
        '       Btsdk_PushMsg_Func          push_msg;

        '   BtSdkMAPMSEStatusRoutinesStru   mse_status_rtns;            ;4 x 3
        '       Btsdk_RegisterNotification_Func register_notification;
        '       Btsdk_UnpdateInbox_Func         update_inbox;
        '       Btsdk_GetMSETime_Func           get_mse_time;


        Dim sizeOfStru As Integer = (4 * 3) + (4 * 3) + (4 * 7) + (4 * 4) + (4 * 3)
        ReDim inpByteArray(0 To sizeOfStru - 1)

        Dim currByteIdx As Integer = 0

        Array.Copy(struFindFolderRoutines, 0, inpByteArray, currByteIdx, struFindFolderRoutines.Length)
        currByteIdx = currByteIdx + struFindFolderRoutines.Length

        Array.Copy(struFindMsgRoutines, 0, inpByteArray, currByteIdx, struFindMsgRoutines.Length)
        currByteIdx = currByteIdx + struFindMsgRoutines.Length

        Array.Copy(struFileIOroutines, 0, inpByteArray, currByteIdx, struFileIOroutines.Length)
        currByteIdx = currByteIdx + struFileIOroutines.Length

        Array.Copy(struMsgIOroutines, 0, inpByteArray, currByteIdx, struMsgIOroutines.Length)
        currByteIdx = currByteIdx + struMsgIOroutines.Length

        Array.Copy(struMsgStatusRoutines, 0, inpByteArray, currByteIdx, struMsgStatusRoutines.Length)
        currByteIdx = currByteIdx + struMsgStatusRoutines.Length


    End Sub

    Private Sub BlueSoleil_MAP_InitStruBytes_mapGetMessageListParams(ByRef inpByteArray() As Byte, ByVal getUnreadOnly As Boolean)

        Dim struSize As Integer = 4 + BTSDK_MAP_FOLDER_LEN + 2 + 2 + 4 + BTSDK_MAP_TIME_LEN + BTSDK_MAP_TIME_LEN + BTSDK_MAP_USERNAME_LEN + BTSDK_MAP_USERNAME_LEN + 1 + 1 + 1 + 1 + 2 + 1 + BTSDK_MAP_MSE_TIME_LEN
        ReDim inpByteArray(0 To struSize - 1)

        Dim gmlMask As UInt32 = BTSDK_MAP_GMLP_MAXCOUNT Or BTSDK_MAP_GMLP_STARTOFF Or BTSDK_MAP_GMLP_READSTATUS Or BTSDK_MAP_GMLP_NEWMSG 'Or BTSDK_MAP_GFLP_LISTSIZE 'Or BTSDK_MAP_GMLP_MSGTYPE

        Dim gmlParams As UInt32 = BTSDK_MAP_MP_SENDERNAME Or BTSDK_MAP_MP_TEXT Or BTSDK_MAP_MP_SUBJECT ' Or BTSDK_MAP_MP_READ 'Or BTSDK_MAP_MP_PROTECTED

        Dim tempBytes(0 To 0) As Byte

        Dim currOutIdx As Integer = 0

        ReDim tempBytes(0 To 3)
        tempBytes = BitConverter.GetBytes(gmlMask)
        Array.Copy(tempBytes, 0, inpByteArray, currOutIdx, tempBytes.Length)
        currOutIdx = currOutIdx + tempBytes.Length

        currOutIdx = currOutIdx + BTSDK_MAP_FOLDER_LEN   'should be able to skip remote folder name since we set the path elsewhere.

        inpByteArray(currOutIdx) = &HFF             'max count.
        inpByteArray(currOutIdx + 1) = &HFF
        currOutIdx = currOutIdx + 2

        inpByteArray(currOutIdx) = 0             'start offset.
        inpByteArray(currOutIdx + 1) = 0
        currOutIdx = currOutIdx + 2

        ReDim tempBytes(0 To 3)
        tempBytes = BitConverter.GetBytes(gmlParams)
        Array.Copy(tempBytes, 0, inpByteArray, currOutIdx, tempBytes.Length)
        currOutIdx = currOutIdx + tempBytes.Length

        currOutIdx = currOutIdx + BTSDK_MAP_TIME_LEN   'should be able to skip time period start since we are not filtering.
        currOutIdx = currOutIdx + BTSDK_MAP_TIME_LEN   'should be able to skip time period end since we are not filtering.

        currOutIdx = currOutIdx + BTSDK_MAP_USERNAME_LEN   'should be able to skip sender name since we are not filtering.
        currOutIdx = currOutIdx + BTSDK_MAP_USERNAME_LEN   'should be able to skip recipient name since we are not filtering.

        inpByteArray(currOutIdx) = 0
        currOutIdx = currOutIdx + 1                    ' msg type..


        If getUnreadOnly = True Then                    'read_unread_status
            inpByteArray(currOutIdx) = BTSDK_MAP_MSG_FILTER_ST_UNREAD
        Else
            inpByteArray(currOutIdx) = BTSDK_MAP_MSG_FILTER_ST_ALL
        End If
        currOutIdx = currOutIdx + 1

        currOutIdx = currOutIdx + 1        'should be able to skip priority since we are not filtering.

        currOutIdx = currOutIdx + 1        'should be able to skip subject length since we are not filtering.

        currOutIdx = currOutIdx + 2        'should be able to skip list size since it is a return value.

        currOutIdx = currOutIdx + 1        'should be able to skip new-msg-indicator since it is a return value.

        currOutIdx = currOutIdx + BTSDK_MAP_MSE_TIME_LEN        'should be able to skip server time since it is a return value.


    End Sub


    Private Sub BlueSoleil_MAP_InitStruBytes_mapGetMessageParam(ByRef inpByteArray() As Byte, ByVal msgHandle As String, ByVal getAttachments As Boolean)

        'handle, charset_flag, attachment_bool, fraction_req, fraction_del



        Dim struSize As Integer = BTSDK_MAP_MSGHDL_LEN + 1 + 1 + 1 + 1
        ReDim inpByteArray(0 To struSize - 1)

        'put msg handle starting in byte zero.
        Dim tempBytes(0 To 0) As Byte
        tempBytes = System.Text.UTF8Encoding.UTF8.GetBytes(msgHandle & Chr(0))
        ReDim Preserve tempBytes(0 To BTSDK_MAP_MSGHDL_LEN - 1)
        Array.Copy(tempBytes, 0, inpByteArray, 0, tempBytes.Length)


        'assign charset
        inpByteArray(BTSDK_MAP_MSGHDL_LEN) = BTSDK_MAP_CHARSET_UTF8

        'attachment?  
        If getAttachments = True Then
            inpByteArray(BTSDK_MAP_MSGHDL_LEN + 1) = 1
        End If

        ''?


    End Sub

    Private Sub BlueSoleil_MAP_InitStruBytes_mapPushMessageParam(ByRef inpByteArray() As Byte, ByVal remoteFolder As String)

        'folder, save_copy, retry, charset, msgHandle



        Dim struSize As Integer = BTSDK_MAP_FOLDER_LEN + 1 + 1 + 1 + BTSDK_MAP_MSGHDL_LEN
        ReDim inpByteArray(0 To struSize - 1)

        'put folder starting in byte zero...
        Dim tempBytes(0 To 0) As Byte

        tempBytes = System.Text.UTF8Encoding.UTF8.GetBytes(remoteFolder & Chr(0))
        ReDim Preserve tempBytes(0 To BTSDK_MAP_FOLDER_LEN - 1)
        Array.Copy(tempBytes, 0, inpByteArray, 0, tempBytes.Length)


        'save copy.
        ' inpByteArray(BTSDK_MAP_FOLDER_LEN) = 1

        'retry?  hell yes if necessary.
        ' inpByteArray(BTSDK_MAP_FOLDER_LEN + 1) = 1

        'charset.
        inpByteArray(BTSDK_MAP_FOLDER_LEN + 2) = BTSDK_MAP_CHARSET_UTF8

        ''?


    End Sub


    Public Function BlueSoleil_MAP_PushMessage_BMSG(ByVal connHandle As UInt32, ByVal BMSGfilename As String) As Boolean

        If IO.File.Exists(BMSGfilename) = False Then
            Return False
        End If

        BlueSoleil_MAP_RegisterFileIOroutines(connHandle)
        BlueSoleil_MAP_RegisterStatusCallback(connHandle)


        Dim retUInt As UInt32

        Dim dirBytes(0 To 0) As Byte
        retUInt = Btsdk_MAPSetFolder(connHandle, dirBytes(0))


        dirBytes = System.Text.UTF8Encoding.UTF8.GetBytes("telecom" & Chr(0))
        retUInt = Btsdk_MAPSetFolder(connHandle, dirBytes(0))

        dirBytes = System.Text.UTF8Encoding.UTF8.GetBytes("msg" & Chr(0))
        retUInt = Btsdk_MAPSetFolder(connHandle, dirBytes(0))



        Dim bmsgFileHandle As IntPtr = FileAPI_OpenFile(BMSGfilename, False)
        Dim bmsgFileHandle_INT32 As UInt32 = CUInt(bmsgFileHandle)

        Dim struPushMsgParam(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapPushMessageParam(struPushMsgParam, "outbox")

        retUInt = Btsdk_MAPPushMessage(connHandle, struPushMsgParam(0), bmsgFileHandle_INT32)

        FileAPI_CloseFile(bmsgFileHandle)

        If retUInt = 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Function BlueSoleil_MAP_XML_ConvertDateTimeString(ByVal inpDateTimeStr As String) As DateTime

        Dim inpYear As Integer = CInt(Val(Mid(inpDateTimeStr, 1, 4)))
        Dim inpMonth As Integer = CInt(Val(Mid(inpDateTimeStr, 5, 2)))
        Dim inpDay As Integer = CInt(Val(Mid(inpDateTimeStr, 7, 2)))

        Dim inpHour As Integer = CInt(Val(Mid(inpDateTimeStr, 10, 2)))
        Dim inpMinute As Integer = CInt(Val(Mid(inpDateTimeStr, 12, 2)))
        Dim inpSecond As Integer = CInt(Val(Mid(inpDateTimeStr, 14, 2)))

        'Dim inpZoneOffset As Integer = CInt(Val(Mid(inpDateTimeStr, 17, 5)))

        Dim retDateTime As New DateTime(inpYear, inpMonth, inpDay, inpHour, inpMinute, inpSecond)

        Return retDateTime

    End Function



    Public Function BlueSoleil_MAP_XML_GetFolderListInfo(ByVal fnMsgList As String, ByRef retFolders() As String) As Integer

        'returns the number of messages.
        Dim retMsgCount As Integer = 0

        'read XML file.


        Dim inpXMLreader As New Xml.XmlTextReader(fnMsgList)

        Dim lineNodeName As String = ""
        Dim lineNodeText As String = ""


        Dim outItemCount As Integer = 0

        Dim tempStr As String = ""

        Dim TorF As Boolean = True
        Do
            If inpXMLreader.EOF = True Then Exit Do

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

                    Case "FOLDER"
                        ReDim Preserve retFolders(0 To outItemCount)
                        tempStr = inpXMLreader.GetAttribute("name")
                        retFolders(outItemCount) = tempStr
                        outItemCount = outItemCount + 1

                End Select
            End If
        Loop


        inpXMLreader.Close()

        Return outItemCount

    End Function



    Public Function BlueSoleil_MAP_XML_GetMessageListInfo(ByVal fnMsgList As String, ByRef retHandles() As String, ByRef retMsgSubjects() As String, ByRef retDateTimes() As DateTime, ByRef retSenderNames() As String, ByRef retSenderAddresses() As String, ByRef retRecipAddresses() As String, ByRef retMsgTypes() As String, ByRef retMsgSizes() As Integer, ByRef retAttachmentSizes() As Integer, ByRef retMsgIsRead() As Boolean) As Integer

        'returns the number of messages.
        Dim retMsgCount As Integer = 0

        'read XML file.


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


                    Case "MSG"

                        ReDim Preserve retHandles(0 To retMsgCount)
                        ReDim Preserve retMsgSubjects(0 To retMsgCount)
                        ReDim Preserve retDateTimes(0 To retMsgCount)
                        ReDim Preserve retSenderNames(0 To retMsgCount)
                        ReDim Preserve retSenderAddresses(0 To retMsgCount)
                        ReDim Preserve retRecipAddresses(0 To retMsgCount)
                        ReDim Preserve retMsgTypes(0 To retMsgCount)
                        ReDim Preserve retMsgSizes(0 To retMsgCount)
                        ReDim Preserve retAttachmentSizes(0 To retMsgCount)
                        ReDim Preserve retMsgIsRead(0 To retMsgCount)


                        tempStr = inpXMLreader.GetAttribute("handle")
                        retHandles(retMsgCount) = tempStr

                        tempStr = inpXMLreader.GetAttribute("subject")
                        retMsgSubjects(retMsgCount) = tempStr

                        tempStr = inpXMLreader.GetAttribute("datetime")
                        If tempStr <> "" Then
                            retDateTimes(retMsgCount) = BlueSoleil_MAP_XML_ConvertDateTimeString(tempStr)
                        End If

                        tempStr = inpXMLreader.GetAttribute("sender_name")
                        retSenderNames(retMsgCount) = tempStr

                        tempStr = inpXMLreader.GetAttribute("sender_addressing")
                        retSenderAddresses(retMsgCount) = tempStr

                        tempStr = inpXMLreader.GetAttribute("recipient_addressing")
                        retRecipAddresses(retMsgCount) = tempStr

                        tempStr = inpXMLreader.GetAttribute("type")
                        retMsgTypes(retMsgCount) = tempStr

                        tempStr = inpXMLreader.GetAttribute("size")
                        retMsgSizes(retMsgCount) = CInt(Val(tempStr))

                        tempStr = inpXMLreader.GetAttribute("attachment_size")
                        retAttachmentSizes(retMsgCount) = CInt(Val(tempStr))

                        tempStr = inpXMLreader.GetAttribute("read")
                        If UCase(tempStr) = "NO" Then
                            retMsgIsRead(retMsgCount) = False
                        Else
                            retMsgIsRead(retMsgCount) = True
                        End If


                        retMsgCount = retMsgCount + 1

                End Select

                'ElseIf inpXMLreader.NodeType = XmlNodeType.Text Then
                '    lineNodeName = inpXMLreader.Name
                '    lineNodeName = UCase(String_CleanSimpleXMLtext(lineNodeName))
                '    Select Case lineNodeName
                '    Case "WOEID"
                '    ReDim Preserve retWOEIDs(0 To retWOEIDcount)
                '    retWOEIDs(retWOEIDcount) = tempStr
                '    retWOEIDcount = retWOEIDcount + 1
                '    End Select

            End If

        Loop


        inpXMLreader.Close()

        Return retMsgCount

    End Function

    Private Sub BlueSoleil_MAP_InitStruBytes_mapMessageObj(ByRef inpByteArray() As Byte)

        Dim struSize As Integer = BTSDK_MAP_MSGHDL_LEN + 4 + 4 + 4 + BTSDK_MAP_SUBJECT_LEN + BTSDK_MAP_USERNAME_LEN + BTSDK_MAP_ADDR_LEN + BTSDK_MAP_ADDR_LEN + BTSDK_MAP_USERNAME_LEN + BTSDK_MAP_ADDR_LEN + BTSDK_MAP_MSGTYPE_LEN + BTSDK_MAP_TIME_LEN + 1 + 1 + 1 + 1 + 1 + 1
        ReDim inpByteArray(0 To struSize - 1)

        '!!!


    End Sub



    Private Sub BlueSoleil_MAP_InitStruBytes_mapMessageFilter(ByRef inpByteArray() As Byte)

        Dim struSize As Integer = 4 + 4 + BTSDK_MAP_TIME_LEN + BTSDK_MAP_TIME_LEN + BTSDK_MAP_USERNAME_LEN + BTSDK_MAP_USERNAME_LEN + 1 + 1 + 1 + 1
        ReDim inpByteArray(0 To struSize - 1)

        '!!!


    End Sub


    Private Function BSmsgIO_APP_RegisterNotification(ByVal connHandle As UInt32, ByVal masSvcHandle As UInt32, ByVal turnOn As Byte) As Byte

        Debug.Print("BSfileIO_APP_RegisterNotification")

        'guessing here.  probably wrong, as I think this a server-side function.

        'BlueSoleil_MAP_EnableNotifications(connHandle, (turnOn = 1))

        Return 1

    End Function

    Private Function BSmsgIO_APP_UpdateInbox() As Byte

        Debug.Print("BSfileIO_APP_UpdateInbox")

        'guessing here.  probably wrong, as I think this a server-side function.

        'BlueSoleil_MAP_UpdateInbox(BSfileIO_APP_ConnHandle)

        Return 1

    End Function

    Private Function BSmsgIO_APP_GetMSETime(ByVal ptrStru_MSETime As IntPtr) As Byte

        Debug.Print("BSfileIO_APP_GetMSETime")


        Dim tempTimeStr As String = ""
        Dim tempBytes(0 To 0) As Byte
        Dim tempDate As DateTime = Now
        tempTimeStr = Format(tempDate.Year, "0000") & Format(tempDate.Month, "00" & Format(tempDate.Day, "00")) & "T" & Format(tempDate.Hour, "00") & Format(tempDate.Minute, "00") & Format(tempDate.Second, "00")
        tempTimeStr = tempTimeStr & "Z+0800"

        tempBytes = System.Text.UTF8Encoding.UTF8.GetBytes(tempTimeStr & Chr(0))
        ReDim Preserve tempBytes(0 To BTSDK_MAP_MSE_TIME_LEN - 1)



        Marshal.Copy(tempBytes, 0, ptrStru_MSETime, tempBytes.Length)

        Return 1

    End Function



    Private Sub BlueSoleil_MAP_InitStruBytes_mapMessageStatusRoutines(ByRef inpByteArray() As Byte, ByVal functPtr_APPRegisterNotification As UInt32, ByVal functPtr_APPUpdateInbox As UInt32, ByVal functPtr_APPGetMSETime As UInt32)

        Dim sizeOfStru As Integer = 3 * 4

        ReDim inpByteArray(0 To sizeOfStru - 1)

        Dim tempBytes(0 To 3) As Byte


        tempBytes = BitConverter.GetBytes(functPtr_APPRegisterNotification)
        Array.Copy(tempBytes, 0, inpByteArray, 0, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPUpdateInbox)
        Array.Copy(tempBytes, 0, inpByteArray, 4, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPGetMSETime)
        Array.Copy(tempBytes, 0, inpByteArray, 8, 4)


    End Sub


    Private Sub BlueSoleil_MAP_InitStruBytes_mapMessageIORoutines(ByRef inpByteArray() As Byte, ByVal functPtr_APPModifyMsgStatus As UInt32, ByVal functPtr_APPCreateBMsgFile As UInt32, ByVal functPtr_APPOpenBMsgFile As UInt32, ByVal functPtr_APPPushMsg As UInt32)

        Dim sizeOfStru As Integer = 4 * 4

        ReDim inpByteArray(0 To sizeOfStru - 1)

        Dim tempBytes(0 To 3) As Byte


        tempBytes = BitConverter.GetBytes(functPtr_APPModifyMsgStatus)
        Array.Copy(tempBytes, 0, inpByteArray, 0, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPCreateBMsgFile)
        Array.Copy(tempBytes, 0, inpByteArray, 4, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPOpenBMsgFile)
        Array.Copy(tempBytes, 0, inpByteArray, 8, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPPushMsg)
        Array.Copy(tempBytes, 0, inpByteArray, 12, 4)

    End Sub



    Private Sub BlueSoleil_MAP_InitStruBytes_mapFolderObj(ByRef inpByteArray() As Byte, ByVal folderName As String)

        'get the folder info and populate.

        Dim struSize As Integer = 4 + BTSDK_MAP_FOLDER_LEN + BTSDK_MAP_TIME_LEN + BTSDK_MAP_TIME_LEN + BTSDK_MAP_TIME_LEN
        ReDim inpByteArray(0 To struSize - 1)

        Dim dInfo As New IO.DirectoryInfo(folderName)

        Dim tempTimeStr As String = ""

        Dim tempBytes(0 To 0) As Byte

        Dim currByteIdx As Integer = 0


        'size
        Dim dirSize As Long = FileAPI_GetDirectorySize(folderName, False)
        If dirSize > 2 ^ 31 Then
            dirSize = CLng(2 ^ 31)
        End If
        Dim dirSizeInt As Integer = CInt(dirSize)
        ReDim tempBytes(0 To 3)
        tempBytes = BitConverter.GetBytes(dirSizeInt)
        Array.Copy(tempBytes, 0, inpByteArray, currByteIdx, tempBytes.Length)
        currByteIdx = currByteIdx + tempBytes.Length


        'dir title (not full path)
        Dim shortPathName As String = IO.Path.GetFileName(folderName)
        tempBytes = System.Text.UTF8Encoding.UTF8.GetBytes(shortPathName & Chr(0))
        ReDim Preserve tempBytes(0 To BTSDK_MAP_FOLDER_LEN - 1)
        Array.Copy(tempBytes, 0, inpByteArray, currByteIdx, tempBytes.Length)
        currByteIdx = currByteIdx + tempBytes.Length


        'dir create time
        tempTimeStr = Format(dInfo.CreationTime.Year, "0000") & Format(dInfo.CreationTime.Month, "00" & Format(dInfo.CreationTime.Day, "00")) & "T" & Format(dInfo.CreationTime.Hour, "00") & Format(dInfo.CreationTime.Minute, "00") & Format(dInfo.CreationTime.Second, "00")
        tempBytes = System.Text.UTF8Encoding.UTF8.GetBytes(tempTimeStr & Chr(0))
        ReDim Preserve tempBytes(0 To BTSDK_MAP_TIME_LEN - 1)
        Array.Copy(tempBytes, 0, inpByteArray, currByteIdx, tempBytes.Length)
        currByteIdx = currByteIdx + tempBytes.Length


        'dir access time
        tempTimeStr = Format(dInfo.LastAccessTime.Year, "0000") & Format(dInfo.LastAccessTime.Month, "00" & Format(dInfo.LastAccessTime.Day, "00")) & "T" & Format(dInfo.LastAccessTime.Hour, "00") & Format(dInfo.LastAccessTime.Minute, "00") & Format(dInfo.LastAccessTime.Second, "00")
        tempBytes = System.Text.UTF8Encoding.UTF8.GetBytes(tempTimeStr & Chr(0))
        ReDim Preserve tempBytes(0 To BTSDK_MAP_TIME_LEN - 1)
        Array.Copy(tempBytes, 0, inpByteArray, currByteIdx, tempBytes.Length)
        currByteIdx = currByteIdx + tempBytes.Length


        'dir modified time
        tempTimeStr = Format(dInfo.LastWriteTime.Year, "0000") & Format(dInfo.LastWriteTime.Month, "00" & Format(dInfo.LastWriteTime.Day, "00")) & "T" & Format(dInfo.LastWriteTime.Hour, "00") & Format(dInfo.LastWriteTime.Minute, "00") & Format(dInfo.LastWriteTime.Second, "00")
        tempBytes = System.Text.UTF8Encoding.UTF8.GetBytes(tempTimeStr & Chr(0))
        ReDim Preserve tempBytes(0 To BTSDK_MAP_TIME_LEN - 1)
        Array.Copy(tempBytes, 0, inpByteArray, currByteIdx, tempBytes.Length)
        currByteIdx = currByteIdx + tempBytes.Length


    End Sub

    Private Sub BlueSoleil_MAP_InitStruBytes_mapGetFolderListParams(ByRef inpByteArray() As Byte)

        Dim struSize As Integer = 2 + 2 + 2 + 2
        ReDim inpByteArray(0 To struSize - 1)

        Dim gflMask As UInt16 = BTSDK_MAP_GFLP_MAXCOUNT Or BTSDK_MAP_GFLP_STARTOFF Or BTSDK_MAP_GFLP_LISTSIZE

        'much like the PBAP param structure, the 16 bit values appear to have the byte order reversed, or maybe I'm not reading the notation correctly.

        inpByteArray(1) = 0
        inpByteArray(0) = CByte(gflMask)

        inpByteArray(3) = &HFF '0 
        inpByteArray(2) = &HFF '1 

        inpByteArray(4) = 0
        inpByteArray(5) = 0

        inpByteArray(6) = 0
        inpByteArray(7) = 0

    End Sub


    Private Sub BlueSoleil_MAP_InitStruBytes_mapFindMessageRoutines(ByRef inpByteArray() As Byte, ByVal functPtr_APPFindMessageFirst As UInt32, ByVal functPtr_APPFindMessageNext As UInt32, ByVal functPtr_APPFindMessageClose As UInt32)

        Dim sizeOfStru As Integer = 3 * 4

        ReDim inpByteArray(0 To sizeOfStru - 1)

        Dim tempBytes(0 To 3) As Byte


        tempBytes = BitConverter.GetBytes(functPtr_APPFindMessageFirst)
        Array.Copy(tempBytes, 0, inpByteArray, 0, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPFindMessageNext)
        Array.Copy(tempBytes, 0, inpByteArray, 4, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPFindMessageClose)
        Array.Copy(tempBytes, 0, inpByteArray, 8, 4)

    End Sub

    Private Sub BlueSoleil_MAP_InitStruBytes_mapFindFolderRoutines(ByRef inpByteArray() As Byte, ByVal functPtr_APPFindFolderFirst As UInt32, ByVal functPtr_APPFindFolderNext As UInt32, ByVal functPtr_APPFindFolderClose As UInt32)

        Dim sizeOfStru As Integer = 3 * 4

        ReDim inpByteArray(0 To sizeOfStru - 1)

        Dim tempBytes(0 To 3) As Byte


        tempBytes = BitConverter.GetBytes(functPtr_APPFindFolderFirst)
        Array.Copy(tempBytes, 0, inpByteArray, 0, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPFindFolderNext)
        Array.Copy(tempBytes, 0, inpByteArray, 4, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPFindFolderClose)
        Array.Copy(tempBytes, 0, inpByteArray, 8, 4)

    End Sub

    Private Sub BlueSoleil_MAP_InitStruBytes_mapFileIOroutines(ByRef inpByteArray() As Byte, ByVal functPtr_APPOpenFile As UInt32, ByVal functPtr_APPCreateFile As UInt32, ByVal functPtr_APPWriteFile As UInt32, ByVal functPtr_APPReadFile As UInt32, ByVal functPtr_APPGetFileSize As UInt32, ByVal functPtr_APPRewindFile As UInt32, ByVal functPtr_APPCloseFile As UInt32)

        Dim sizeOfStru As Integer = 7 * 4

        ReDim inpByteArray(0 To sizeOfStru - 1)

        Dim tempBytes(0 To 3) As Byte

        tempBytes = BitConverter.GetBytes(functPtr_APPOpenFile)
        Array.Copy(tempBytes, 0, inpByteArray, 0, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPCreateFile)
        Array.Copy(tempBytes, 0, inpByteArray, 4, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPWriteFile)
        Array.Copy(tempBytes, 0, inpByteArray, 8, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPReadFile)
        Array.Copy(tempBytes, 0, inpByteArray, 12, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPGetFileSize)
        Array.Copy(tempBytes, 0, inpByteArray, 16, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPRewindFile)
        Array.Copy(tempBytes, 0, inpByteArray, 20, 4)

        tempBytes = BitConverter.GetBytes(functPtr_APPCloseFile)
        Array.Copy(tempBytes, 0, inpByteArray, 24, 4)


    End Sub





    Public Function BlueSoleil_MAP_UnregisterServers(ByVal mnsServiceHandle As UInt32, ByVal masServiceHandle As UInt32) As Boolean



        Dim retUInt As UInt32 = 0
        If mnsServiceHandle <> 0 Then
            retUInt = Btsdk_UnregisterMAPService(mnsServiceHandle)

        End If

        If masServiceHandle <> 0 Then
            retUInt = Btsdk_UnregisterMAPService(masServiceHandle)

        End If

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_MAP_RegisterNotificationService() As UInt32

        Dim funcPtr_MsgNotification As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateMsgNotification)

        Dim retUInt As UInt32

        retUInt = BlueSoleil_MAP_RegisterMessageNotificationService("BSmapMNS", CUInt(funcPtr_MsgNotification))
        Return retUInt

    End Function



    Private Function BlueSoleil_MAP_RegisterStatusCallback(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        delegateMAPstatusCallback = AddressOf BlueSoleil_MAP_Callback_Status

        Dim retUInt As UInt32

        Dim functPtr_StatusCallback As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateMAPstatusCallback)
        retUInt = Btsdk_MAPRegisterStatusCallback(connHandle, functPtr_StatusCallback)

        Return (retUInt = BTSDK_OK)

    End Function


    Private Function BlueSoleil_MAP_UnregisterStatusCallback(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        Dim retUInt As UInt32
        Dim functPtr_StatusCallback As IntPtr = IntPtr.Zero
        retUInt = Btsdk_MAPRegisterStatusCallback(connHandle, functPtr_StatusCallback)

        Return (retUInt = 0)

    End Function

    Private Function BlueSoleil_MAP_RegisterFileIOroutines(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False


        ' Pinner_PinDelegate(delegateAPPOpenFile)
        ' Pinner_PinDelegate(delegateAPPCreateFile)
        ' Pinner_PinDelegate(delegateAPPWriteFile)
        ' Pinner_PinDelegate(delegateAPPReadFile)
        ' Pinner_PinDelegate(delegateAPPGetFileSize)
        ' Pinner_PinDelegate(delegateAPPRewindFile)
        ' Pinner_PinDelegate(delegateAPPCloseFile)
        BSfileIO_APP_ConnHandle = connHandle

        Dim functPtr_APPopenFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPOpenFile)
        Dim functPtr_APPcreateFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPCreateFile)
        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)
        Dim functPtr_APPreadFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPReadFile)
        Dim functPtr_APPgetFileSize As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPGetFileSize)
        Dim functPtr_APPrewindFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPRewindFile)
        Dim functPtr_APPcloseFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPCloseFile)

        Dim retUInt As UInt32

        'initialize IO routine pointer(s) for BS to use to write the VCF file.  we only need to provide a pointer to APP_WriteFile (which wraps the Win32 WriteFile API)
        Dim mapIOroutinesBytes(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapFileIOroutines(mapIOroutinesBytes, CUInt(functPtr_APPopenFile), CUInt(functPtr_APPcreateFile), CUInt(functPtr_APPwriteFile), CUInt(functPtr_APPreadFile), CUInt(functPtr_APPgetFileSize), CUInt(functPtr_APPrewindFile), CUInt(functPtr_APPcloseFile))

        retUInt = Btsdk_MAPRegisterFileIORoutines(connHandle, mapIOroutinesBytes(0))

        Return (retUInt = BTSDK_OK)

    End Function




    Private Function BlueSoleil_MAP_RegisterMessageAccessService(ByVal nameToRegisterServiceAs As String, ByVal localBluetoothFolder As String) As UInt32

        'returns local service handle.
        '

        'init structure 1.
        Dim srvrAttribStruBytes(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_masLocalServerAttrStru(srvrAttribStruBytes, localBluetoothFolder)


        'init structure 2.
        Dim srvrCallbacksStruBytes(0 To 0) As Byte

        'and sub-structures.
        Dim struCB_FindFolderRoutines(0 To 0) As Byte
        Dim struCB_FindMsgRoutines(0 To 0) As Byte
        Dim struCB_FileIORoutines(0 To 0) As Byte
        Dim struCB_MsgIORoutines(0 To 0) As Byte
        Dim struCB_MsgStatusRoutines(0 To 0) As Byte


        Dim functPtr_APPFindFolderFirst As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPFindFolderFirst)
        Dim functPtr_APPFindFolderNext As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPFindFolderNext)
        Dim functPtr_APPFindFolderClose As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPFindFolderClose)
        BlueSoleil_MAP_InitStruBytes_mapFindFolderRoutines(struCB_FindFolderRoutines, CUInt(functPtr_APPFindFolderFirst), CUInt(functPtr_APPFindFolderNext), CUInt(functPtr_APPFindFolderClose))

        Dim functPtr_APPFindMessageFirst As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPFindMessageFirst)
        Dim functPtr_APPFindMessageNext As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPFindMessageNext)
        Dim functPtr_APPFindMessageClose As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPFindMessageClose)
        BlueSoleil_MAP_InitStruBytes_mapFindMessageRoutines(struCB_FindMsgRoutines, CUInt(functPtr_APPFindMessageFirst), CUInt(functPtr_APPFindMessageNext), CUInt(functPtr_APPFindMessageClose))

        Dim functPtr_APPopenFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPOpenFile)
        Dim functPtr_APPcreateFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPCreateFile)
        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)
        Dim functPtr_APPreadFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPReadFile)
        Dim functPtr_APPgetFileSize As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPGetFileSize)
        Dim functPtr_APPrewindFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPRewindFile)
        Dim functPtr_APPcloseFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPCloseFile)
        BlueSoleil_MAP_InitStruBytes_mapFileIOroutines(struCB_FileIORoutines, CUInt(functPtr_APPopenFile), CUInt(functPtr_APPcreateFile), CUInt(functPtr_APPwriteFile), CUInt(functPtr_APPreadFile), CUInt(functPtr_APPgetFileSize), CUInt(functPtr_APPrewindFile), CUInt(functPtr_APPcloseFile))


        'msg io routines.
        '
        '!!!
        Dim functPtr_APPModifyMsgStatus As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPModifyMessageStatus)
        Dim functPtr_APPCreateBMsgFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPCreateBMessageFile)
        Dim functPtr_APPOpenBMsgFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPOpenBMessageFile)
        Dim functPtr_APPPushMsg As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPPushMessage)
        BlueSoleil_MAP_InitStruBytes_mapMessageIORoutines(struCB_MsgIORoutines, CUInt(functPtr_APPModifyMsgStatus), CUInt(functPtr_APPCreateBMsgFile), CUInt(functPtr_APPOpenBMsgFile), CUInt(functPtr_APPPushMsg))


        Dim functPtr_APPRegisterNotification As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPRegisterNotification)
        Dim functPtr_APPUpdateInbox As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPUpdateInbox)
        Dim functPtr_APPGetMSETime As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPGetMSETime)
        BlueSoleil_MAP_InitStruBytes_mapMessageStatusRoutines(struCB_MsgStatusRoutines, CUInt(functPtr_APPRegisterNotification), CUInt(functPtr_APPUpdateInbox), CUInt(functPtr_APPGetMSETime))

        BlueSoleil_MAP_InitStruBytes_masServerCBstru(srvrCallbacksStruBytes, struCB_FindFolderRoutines, struCB_FindMsgRoutines, struCB_FileIORoutines, struCB_MsgIORoutines, struCB_MsgStatusRoutines)



        Dim svcNameBytes(0 To 0) As Byte
        svcNameBytes = System.Text.UTF8Encoding.UTF8.GetBytes(nameToRegisterServiceAs & Chr(0))


        Dim retUInt As UInt32
        retUInt = Btsdk_RegisterMASService(svcNameBytes(0), srvrAttribStruBytes(0), srvrCallbacksStruBytes(0))

        If retUInt = 0 Then
            'some error?
            Return 0

        End If

        Return retUInt

    End Function

    Private Function BlueSoleil_MAP_RegisterMessageNotificationService(ByVal nameToRegisterServiceAs As String, ByVal funcPtr_MessageNotification As UInt32) As UInt32

        'returns local service handle.

        Dim functPtr_APPopenFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPOpenFile)
        Dim functPtr_APPcreateFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPCreateFile)
        Dim functPtr_APPwriteFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPWriteFile)
        Dim functPtr_APPreadFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPReadFile)
        Dim functPtr_APPgetFileSize As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPGetFileSize)
        Dim functPtr_APPrewindFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPRewindFile)
        Dim functPtr_APPcloseFile As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(delegateAPPCloseFile)

        Dim retUInt As UInt32

        'initialize IO routine pointer(s) for BS to use to write files.
        Dim mapIOroutinesBytes(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapFileIOroutines(mapIOroutinesBytes, CUInt(functPtr_APPopenFile), CUInt(functPtr_APPcreateFile), CUInt(functPtr_APPwriteFile), CUInt(functPtr_APPreadFile), CUInt(functPtr_APPgetFileSize), CUInt(functPtr_APPrewindFile), CUInt(functPtr_APPcloseFile))

        Dim svcNameBytes(0 To 0) As Byte
        svcNameBytes = System.Text.UTF8Encoding.UTF8.GetBytes(nameToRegisterServiceAs)
        ReDim Preserve svcNameBytes(0 To BTSDK_SERVICENAME_MAXLENGTH - 1)

        retUInt = Btsdk_RegisterMNSService(svcNameBytes(0), funcPtr_MessageNotification, mapIOroutinesBytes(0))

        If retUInt = 0 Then
            'some error?
            Return 0

        End If


        Return retUInt

    End Function


    Public Function BlueSoleil_MAP_EnableNotifications(ByVal connHandle As UInt32, ByVal enableNots As Boolean) As Boolean

        If connHandle = 0 Then Return False

        Dim tfByte As Byte = 0
        If enableNots = True Then tfByte = 1

        Dim retUInt As UInt32 = Btsdk_MAPSetNotificationRegistration(connHandle, tfByte)

        If retUInt = 0 Then

            retUInt = Btsdk_MAPUpdateInbox(connHandle)

            Return True
        Else
            Return False
        End If

    End Function



    Public Function BlueSoleil_MAP_UpdateInbox(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        Dim retUInt As UInt32
        retUInt = Btsdk_MAPUpdateInbox(connHandle)

        Return (retUInt = 0)

    End Function


    Private Function BSfileIO_APP_GetRootDir(ByVal ptrToWriteTo As IntPtr) As String

        Debug.Print("BSfileIO_APP_GetRootDir  ptr = " & ptrToWriteTo.ToInt64)

        Dim retStr As String = BSfileIO_APP_RootBTdir

        Dim pthBytes(0 To 0) As Byte
        pthBytes = System.Text.UTF8Encoding.UTF8.GetBytes(retStr & Chr(0))

        Marshal.Copy(pthBytes, 0, ptrToWriteTo, pthBytes.Length)

        Return retStr & Chr(0)

    End Function



    Public Function BlueSoleil_MAP_PullFolderList(ByVal connHandle As UInt32, ByVal dirlistFilenameToWriteTo As String) As Boolean

        If connHandle = 0 Then Return False


        BSfileIO_APP_RootBTdir = IO.Path.GetDirectoryName(dirlistFilenameToWriteTo)

        Dim retInt32 As UInt32


        BlueSoleil_MAP_RegisterFileIOroutines(connHandle)
        BlueSoleil_MAP_RegisterStatusCallback(connHandle)



        'set primary folder.
        Dim dirBytes(0 To 0) As Byte
        retInt32 = Btsdk_MAPSetFolder(connHandle, dirBytes(0))

        dirBytes = System.Text.UTF8Encoding.UTF8.GetBytes("telecom" & Chr(0))
        retInt32 = Btsdk_MAPSetFolder(connHandle, dirBytes(0))

        dirBytes = System.Text.UTF8Encoding.UTF8.GetBytes("msg" & Chr(0))
        retInt32 = Btsdk_MAPSetFolder(connHandle, dirBytes(0))


        ''''''''''''''''''''''

        'If 1 = 2 Then

        'get folder list.
        Dim struGetFolderListParams(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapGetFolderListParams(struGetFolderListParams)

        If IO.File.Exists(dirlistFilenameToWriteTo) = True Then
            Try
                IO.File.Delete(dirlistFilenameToWriteTo)
            Catch ex As Exception

            End Try
        End If

        'create file for message list, and provide handle.
        Dim hFile_FolderList As IntPtr = FileAPI_OpenFile(dirlistFilenameToWriteTo, True)
        If hFile_FolderList.ToInt64 < 1 Then
            Return False
        End If

        Dim intHFile_FolderList As UInt32 = CType(hFile_FolderList, UInt32)
        retInt32 = Btsdk_MAPGetFolderList(connHandle, struGetFolderListParams(0), intHFile_FolderList)
        FileAPI_CloseFile(hFile_FolderList)

        '''''''''''''''''''''''''''''''''''''''


        If retInt32 <> 0 Then
            Try
                IO.File.Delete(dirlistFilenameToWriteTo)
            Catch ex As Exception

            End Try
        End If


        If retInt32 = 0 Then
            Return True
        Else
            Return False
        End If




    End Function



    Public Function BlueSoleil_MAP_PullMessageList(ByVal connHandle As UInt32, ByVal remoteFolderName As String, ByVal msglistFilenameToWriteTo As String, ByVal retrieveUNREADonly As Boolean) As Boolean

        If connHandle = 0 Then Return False


        BlueSoleil_MAP_UpdateInbox(connHandle)

        BSfileIO_APP_RootBTdir = IO.Path.GetDirectoryName(msglistFilenameToWriteTo)

        Dim retInt32 As UInt32


        BlueSoleil_MAP_RegisterFileIOroutines(connHandle)
        BlueSoleil_MAP_RegisterStatusCallback(connHandle)



        'set folder.
        Dim dirBytes(0 To 0) As Byte
        retInt32 = Btsdk_MAPSetFolder(connHandle, dirBytes(0))

        'telecom/msg/




        dirBytes = System.Text.UTF8Encoding.UTF8.GetBytes("telecom" & Chr(0))
        retInt32 = Btsdk_MAPSetFolder(connHandle, dirBytes(0))

        dirBytes = System.Text.UTF8Encoding.UTF8.GetBytes("msg" & Chr(0))
        retInt32 = Btsdk_MAPSetFolder(connHandle, dirBytes(0))


        'set folder.
        dirBytes = System.Text.UTF8Encoding.UTF8.GetBytes(remoteFolderName & Chr(0))
        retInt32 = Btsdk_MAPSetFolder(connHandle, dirBytes(0))


        'get message list.
        If IO.File.Exists(msglistFilenameToWriteTo) = True Then
            Try
                IO.File.Delete(msglistFilenameToWriteTo)
            Catch ex As Exception

            End Try
        End If

        Dim struGetMessageListParams(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapGetMessageListParams(struGetMessageListParams, retrieveUNREADonly)

        Dim hFile_MessageList As IntPtr = FileAPI_OpenFile(msglistFilenameToWriteTo, True)
        If hFile_MessageList.ToInt64 < 1 Then
            Return False
        End If

        Dim intHFile_MessageList As UInt32 = CType(hFile_MessageList, UInt32)
        retInt32 = Btsdk_MAPGetMessageList(connHandle, struGetMessageListParams(0), intHFile_MessageList)

        FileAPI_CloseFile(hFile_MessageList)



        If retInt32 = BTSDK_ER_REQUEST_TIMEOUT Then
            'one more try.
            Try
                IO.File.Delete(msglistFilenameToWriteTo)
            Catch ex As Exception

            End Try
            hFile_MessageList = FileAPI_OpenFile(msglistFilenameToWriteTo, True)
            If hFile_MessageList.ToInt64 < 1 Then
                Return False
            End If

            intHFile_MessageList = CType(hFile_MessageList, UInt32)
            retInt32 = Btsdk_MAPGetMessageList(connHandle, struGetMessageListParams(0), intHFile_MessageList)

            FileAPI_CloseFile(hFile_MessageList)

        End If


        If retInt32 <> 0 Then
            Try
                IO.File.Delete(msglistFilenameToWriteTo)
            Catch ex As Exception

            End Try
        End If

        If retInt32 = 0 Then
            Return True
        Else
            Return False
        End If


    End Function


    Public Function BlueSoleil_MAP_PullMessage(ByVal connHandle As UInt32, ByVal msgHandle As String, ByVal msgFilenameToWriteTo As String, ByVal getAttachments As Boolean) As Boolean

        If connHandle = 0 Then Return False

        BlueSoleil_MAP_RegisterFileIOroutines(connHandle)
        BlueSoleil_MAP_RegisterStatusCallback(connHandle)


        Dim struMessageParam(0 To 0) As Byte
        BlueSoleil_MAP_InitStruBytes_mapGetMessageParam(struMessageParam, msgHandle, getAttachments)

        Dim hFile_MessageFile As IntPtr = FileAPI_OpenFile(msgFilenameToWriteTo, True)
        If hFile_MessageFile.ToInt64 < 1 Then
            Return False
        End If

        Dim intHFile_MessageList As UInt32 = CType(hFile_MessageFile, UInt32)

        Dim retUInt As UInt32
        retUInt = Btsdk_MAPGetMessage(connHandle, struMessageParam(0), intHFile_MessageList)
        FileAPI_CloseFile(hFile_MessageFile)

        If retUInt <> 0 Then
            Try
                IO.File.Delete(msgFilenameToWriteTo)
            Catch ex As Exception

            End Try
        End If

        Return (retUInt = 0)


    End Function



    Private Sub BlueSoleil_MAP_Callback_MessageNotification(ByVal svcHandle As UInt32, ByVal ptr_EvReportObjStru As IntPtr)

        'gee, i sure would love to see this event fire some day.  

        Debug.Print("MsgNotification!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")

        'RaiseEvent BlueSoleil_Event_MAP_MsgNotification()
        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_MAP_MsgNotification())
        t.Start()

    End Sub

    Private Sub BlueSoleil_MAP_Callback_Status(ByVal first As Byte, ByVal last As Byte, ByVal ptrFileNameBytes As IntPtr, ByVal fileSize As UInt32, ByVal curSize As UInt32)

        'this callback is called by BlueSoleil during MAP transfers.

        Debug.Print("MAP Callback.  first = " & first & "   last = " & last & "   filesize = " & fileSize & "    cursize = " & curSize)


        If last = 0 Then

            If first <> 0 Then
                'start of transfer.  
                BlueSoleil_MAP_Callback_Status_CurrSize = curSize
                BlueSoleil_MAP_Callback_Status_TotalSize = fileSize

                Debug.Print("Start MAP xfer.  0 of " & fileSize & " bytes.")

            Else
                BlueSoleil_MAP_Callback_Status_CurrSize = BlueSoleil_MAP_Callback_Status_CurrSize + curSize
                Debug.Print("Continue MAP xfer.  " & BlueSoleil_MAP_Callback_Status_CurrSize & " of " & fileSize & " bytes.")
            End If

        Else

            'add final bytes.
            BlueSoleil_MAP_Callback_Status_CurrSize = BlueSoleil_MAP_Callback_Status_CurrSize + curSize
            Debug.Print("Finish MAP xfer.  " & BlueSoleil_MAP_Callback_Status_CurrSize & " of " & BlueSoleil_MAP_Callback_Status_TotalSize & " bytes.")

            'do whatever cuz we're done..

            'RaiseEvent BlueSoleil_Event_MAP_TransferComplete()
            ' Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_MAP_TransferComplete())
            ' t.Start()

            'reset.
            ' BlueSoleil_MAP_Callback_Status_TotalSize = 0
            ' BlueSoleil_MAP_Callback_Status_CurrSize = 0

        End If

    End Sub


    Public Function BlueSoleil_MAP_SetMessageStatus_READ(ByVal connHandle As UInt32, ByVal msgHandle As String) As Boolean

        If connHandle = 0 Then Return False

        Dim msgHandleBytes(0 To 0) As Byte
        msgHandleBytes = System.Text.UTF8Encoding.UTF8.GetBytes(msgHandle & Chr(0))

        Dim retUInt As UInt32 = Btsdk_MAPSetMessageStatus(connHandle, msgHandleBytes(0), BTSDK_MAP_MSG_SETST_READ)

        If retUInt = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function BlueSoleil_MAP_SetMessageStatus_UNREAD(ByVal connHandle As UInt32, ByVal msgHandle As String) As Boolean

        If connHandle = 0 Then Return False

        Dim msgHandleBytes(0 To 0) As Byte
        msgHandleBytes = System.Text.UTF8Encoding.UTF8.GetBytes(msgHandle & Chr(0))

        Dim retUInt As UInt32 = Btsdk_MAPSetMessageStatus(connHandle, msgHandleBytes(0), BTSDK_MAP_MSG_SETST_UNREAD)

        If retUInt = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function BlueSoleil_MAP_SetMessageStatus_DELETED(ByVal connHandle As UInt32, ByVal msgHandle As String) As Boolean

        If connHandle = 0 Then Return False

        Dim msgHandleBytes(0 To 0) As Byte
        msgHandleBytes = System.Text.UTF8Encoding.UTF8.GetBytes(msgHandle & Chr(0))

        Dim retUInt As UInt32 = Btsdk_MAPSetMessageStatus(connHandle, msgHandleBytes(0), BTSDK_MAP_MSG_SETST_DELETED)

        If retUInt = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function BlueSoleil_MAP_SetMessageStatus_UNDELETED(ByVal connHandle As UInt32, ByVal msgHandle As String) As Boolean

        If connHandle = 0 Then Return False

        Dim msgHandleBytes(0 To 0) As Byte
        msgHandleBytes = System.Text.UTF8Encoding.UTF8.GetBytes(msgHandle & Chr(0))

        Dim retUInt As UInt32 = Btsdk_MAPSetMessageStatus(connHandle, msgHandleBytes(0), BTSDK_MAP_MSG_SETST_UNDELETED)

        If retUInt = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function

End Module
