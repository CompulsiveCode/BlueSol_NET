'modBlueSoleil_HFP - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

Module modBlueSoleil_HFP


    Public BlueSoleil_HFP_LastBatteryVal As Integer = -1
    Public BlueSoleil_HFP_LastSignalVal As Integer = -1

    Public Event BlueSoleil_Event_HFP_Ringing()
    Public Event BlueSoleil_Event_HFP_Standby()
    Public Event BlueSoleil_Event_HFP_OngoingCall()
    Public Event BlueSoleil_Event_HFP_OutgoingCall()

    Public Event BlueSoleil_Event_HFP_ConnectionReleased()

    Public Event BlueSoleil_Event_HFP_NetworkAvailable()
    Public Event BlueSoleil_Event_HFP_NetworkUnavailable()
    Public Event BlueSoleil_Event_HFP_NetworkOperatorName(ByVal networkName As String)
    Public Event BlueSoleil_Event_HFP_StartRoaming()
    Public Event BlueSoleil_Event_HFP_StopRoaming()

    Public Event BlueSoleil_Event_HFP_CurrentCallInfo(ByVal phoneNo As String, ByVal callIsIncoming As Boolean, ByVal callIsFax As Boolean, ByVal callIsData As Boolean, ByVal callIsMultiParty As Boolean)

    Public Event BlueSoleil_Event_HFP_CallerID(ByVal phoneNo As String, ByVal phoneName As String)
    Public Event BlueSoleil_Event_HFP_SubscriberPhoneNo(ByVal phoneNo As String, ByVal phoneName As String)

    Public Event BlueSoleil_Event_HFP_ModelName(ByVal theName As String)
    Public Event BlueSoleil_Event_HFP_ManufacturerName(ByVal theName As String)
    Public Event BlueSoleil_Event_HFP_ExtCmdInd(ByVal theName As String)

    Public Event BlueSoleil_Event_HFP_SignalQuality(ByVal currPct As Double)
    Public Event BlueSoleil_Event_HFP_BatteryCharge(ByVal currPct As Double)

    Public Event BlueSoleil_Event_HFP_MicVolume(ByVal currPct As Double)
    Public Event BlueSoleil_Event_HFP_SpeakerVolume(ByVal currPct As Double)

    Public Event BlueSoleil_Event_HFP_VoiceCmdActivated()
    Public Event BlueSoleil_Event_HFP_VoiceCmdDeactivated()

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfunc_HFPeventAGCallback(ByVal connHandle As UInt32, ByVal hfpEvent As UInt16, ByVal funcParam As IntPtr, ByVal paramLen As UInt16)
    Public delegateHFPeventAGCallback As delfunc_HFPeventAGCallback = AddressOf BlueSoleil_HFP_Callback_Status_AG_Func

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub delfunc_HFPeventAPCallback(ByVal connHandle As UInt32, ByVal hfpEvent As UInt16, ByVal funcParam As IntPtr, ByVal paramLen As UInt16)
    Public delegateHFPeventAPCallback As delfunc_HFPeventAPCallback = AddressOf BlueSoleil_HFP_Callback_Status_AP_Func


    Private Const BTSDK_OK As UInt32 = 0
    Private Const BTSDK_TRUE As Byte = 1

    Private Const BTSDK_SERVICENAME_MAXLENGTH As UInt16 = 80
    Private Const BTSDK_MAX_SUPPORT_FORMAT As UInt16 = 6       '/* OPP format number */
    Private Const BTSDK_PATH_MAXLENGTH As UInt16 = 256     '/* Shall not be larger than FTP_MAX_PATH and OPP_MAX_PATH */
    Private Const BTSDK_CARDNAME_MAXLENGTH As UInt16 = 256     '/* Shall not be larger than OPP_MAX_NAME */
    Private Const BTSDK_PACKETTYPE_MAXNUM As UInt16 = 10      '/* PAN supported network packet type */

    Private Const BTSDK_CLS_HANDSFREE As UInt32 = &H111E
    Private Const BTSDK_CLS_HANDSFREE_AG As UInt32 = &H111F
    Private Const BTSDK_CLS_HEADSET As UInt32 = &H1108
    Private Const BTSDK_CLS_HEADSET_AG As UInt32 = &H1112



    '/* BRSF feature mask ID for AG*/
    Private Const BTSDK_AG_BRSF_3WAYCALL As UInt32 = &H1 '/* Three-way calling */
    Private Const BTSDK_AG_BRSF_NREC As UInt32 = &H2 '/* EC and/or NR function */
    Private Const BTSDK_AG_BRSF_BVRA As UInt32 = &H4 '/* Voice recognition function */
    Private Const BTSDK_AG_BRSF_INBANDRING As UInt32 = &H8 '/* In-band ring tone capability */
    Private Const BTSDK_AG_BRSF_BINP As UInt32 = &H10 '/* Attach a number to a voice tag */
    Private Const BTSDK_AG_BRSF_REJECT_CALL As UInt32 = &H20 '/* Ability to reject a call */
    Private Const BTSDK_AG_BRSF_ENHANCED_CALLSTATUS As UInt32 = &H40 '/* Enhanced call status */
    Private Const BTSDK_AG_BRSF_ENHANCED_CALLCONTROL As UInt32 = &H80 '/* Enhanced call control */
    Private Const BTSDK_AG_BRSF_EXTENDED_ERRORRESULT As UInt32 = &H100 '/* Extended Error Result Codes */
    Private Const BTSDK_AG_BRSF_ALL As UInt32 = &H1FF '/* Support all the upper features */ 

    '/* BRSF feature mask ID for HF */
    Private Const BTSDK_HF_BRSF_NREC As UInt32 = &H1 '/* EC and/or NR function */
    Private Const BTSDK_HF_BRSF_3WAYCALL As UInt32 = &H2 '/* Call waiting and 3-way calling */
    Private Const BTSDK_HF_BRSF_CLIP As UInt32 = &H4 '/* CLI presentation capability */
    Private Const BTSDK_HF_BRSF_BVRA As UInt32 = &H8 '/* Voice recognition activation */
    Private Const BTSDK_HF_BRSF_RMTVOLCTRL As UInt32 = &H10 '/* Remote volume control */
    Private Const BTSDK_HF_BRSF_ENHANCED_CALLSTATUS As UInt32 = &H20 '/* Enhanced call status */
    Private Const BTSDK_HF_BRSF_ENHANCED_CALLCONTROL As UInt32 = &H40 '/* Enhanced call control */
    Private Const BTSDK_HF_BRSF_ALL As UInt32 = &H7F '/* Support all the upper features */ 

    '/* HSP/HFP AG specific event. */
    Private Const BTSDK_APP_EV_AGAP_BASE As UInt16 = &H900



    '/* Macros for HFP/HSP AG */
    '/* Parameters for Btsdk_AGAP_Init */
    Private Const BTSDK_AGAP_FEA_3WAY_CALLING As UInt32 = &H1
    Private Const BTSDK_AGAP_FEA_NREC As UInt32 = &H2
    Private Const BTSDK_AGAP_FEA_VOICE_RECOG As UInt32 = &H4
    Private Const BTSDK_AGAP_FEA_INBAND_RING As UInt32 = &H8
    Private Const BTSDK_AGAP_FEA_VOICETAG_PHONE_NUM As UInt32 = &H10
    Private Const BTSDK_AGAP_FEA_REJ_CALL As UInt32 = &H20
    Private Const BTSDK_AGAP_SCO_PKT_HV1 As UInt32 = &H20
    Private Const BTSDK_AGAP_SCO_PKT_HV2 As UInt32 = &H40
    Private Const BTSDK_AGAP_SCO_PKT_HV3 As UInt32 = &H80

    '/* Available status for Btsdk_AGAP_GetStatus. */
    Private Const BTSDK_AGAP_STATUS_GENERATE_INBAND_RINGTONE As UInt32 = &H1  '/* whether AG is capable of generating in-band ring tone */
    Private Const BTSDK_AGAP_STATUS_AUDIO_CONN_ONGOING As UInt32 = &H2  '/* whether audio connection with remote device is ongoing */

    '/* Possible AG state of Btsdk_AGAP_GetAGState*/
    Private Const BTSDK_AGAP_ST_IDLE As UInt32 = &H1     '/*before service level connection is established*/
    Private Const BTSDK_AGAP_ST_STANDBY As UInt32 = &H2     '/*service level connection is established*/
    Private Const BTSDK_AGAP_ST_RINGING As UInt32 = &H3     '/*ringing*/
    Private Const BTSDK_AGAP_ST_OUTGOINGCALL As UInt32 = &H4     '/*outgoing call*/
    Private Const BTSDK_AGAP_ST_ONGOINGCALL As UInt32 = &H5     '/*ongoing call*/
    Private Const BTSDK_AGAP_ST_BVRA As UInt32 = &H6     '/*voice recognition is ongoing*/
    Private Const BTSDK_AGAP_ST_VOVG As UInt32 = &H7
    Private Const BTSDK_AGAP_ST_HELDINCOMINGCALL As UInt32 = &H8 '/*the incoming call is held*/
    Private Const BTSDK_AGAP_ST_THREEWAYCALLING As UInt32 = &H9 '/*three way calling*/

    '/* Current state mask code for function Btsdk_AGAP_SetCurIndicatorVal. */
    Private Const BTSDK_AGAP_INDICATOR_SVC_UNAVAILABLE As UInt32 = &H0
    Private Const BTSDK_AGAP_INDICATOR_SVC_AVAILABLE As UInt32 = &H1
    Private Const BTSDK_AGAP_INDICATOR_ACTIVE As UInt32 = &H2
    Private Const BTSDK_AGAP_INDICATOR_INCOMING As UInt32 = &H4
    Private Const BTSDK_AGAP_INDICATOR_DIALING As UInt32 = &H8
    Private Const BTSDK_AGAP_INDICATOR_ALERTING As UInt32 = &H10

    '/* Possible "features" parameter of Btsdk_HFAP_Init */ 
    Private Const BTSDK_HFAP_FEA_NREC As UInt32 = &H1
    Private Const BTSDK_HFAP_FEA_3WAY_CALLING As UInt32 = &H2
    Private Const BTSDK_HFAP_FEA_CALLING_LINE_NUM As UInt32 = &H4
    Private Const BTSDK_HFAP_FEA_VOICE_RECOG As UInt32 = &H8
    Private Const BTSDK_HFAP_FEA_RMT_VOL_CTRL As UInt32 = &H10

    '/* Possible "sco_pkt_type" parameter of Btsdk_HFAP_Init */ 
    Private Const BTSDK_HFAP_SCO_PKT_HV1 As UInt32 = &H20
    Private Const BTSDK_HFAP_SCO_PKT_HV2 As UInt32 = &H40
    Private Const BTSDK_HFAP_SCO_PKT_HV3 As UInt32 = &H80

    '/* Available status from function Btsdk_HFAP_GetStatus. */
    Private Const BTSDK_HFAP_STATUS_LOCAL_GENERATE_RINGTONE As UInt32 = &H1  '/* whether HF device need to generate its own in-band ring tone */
    Private Const BTSDK_HFAP_STATUS_AUDIO_CONN_ONGOING As UInt32 = &H2  '/* whether audio connection with remote device is ongoing */

    '/* Three way calling mode */
    Private Const BTSDK_HFAP_3WAY_MOD0 As UInt32 = 0 ' '/*Set busy tone for a waiting call; Release the held call*/
    Private Const BTSDK_HFAP_3WAY_MOD1 As UInt32 = 1 ' '/*Release activate call & accept held/waiting call*/
    Private Const BTSDK_HFAP_3WAY_MOD2 As UInt32 = 2 ' '/*Swap between active call and held call; Place active call on held; Place held call on active*/
    Private Const BTSDK_HFAP_3WAY_MOD3 As UInt32 = 3 ' '/*Add a held call to the conversation*/
    Private Const BTSDK_HFAP_3WAY_MOD4 As UInt32 = 4 ' '/*Connects the two calls and disconnects the subscriber from both calls (Explicit Call Transfer)*/

    '/* AG Type of the call, possible values of HFP_EV_ANSWER_CALL_REQ and HFP_EV_CANCEL_CALL_REQ event parameter */
    Private Const BTSDK_HFP_TYPE_ALL_CALLS As UInt32 = &H1 '/* (Release) all the existing calls */
    Private Const BTSDK_HFP_TYPE_INCOMING_CALL As UInt32 = &H2 '/* (Reject or accept) the incoming call */ 
    Private Const BTSDK_HFP_TYPE_HELDINCOMING_CALL As UInt32 = &H3 '/* (Reject or accept) the Held incoming call */
    Private Const BTSDK_HFP_TYPE_OUTGOING_CALL As UInt32 = &H4 '/* (Release) the outgoing call */
    Private Const BTSDK_HFP_TYPE_ONGOING_CALL As UInt32 = &H5 '/* (Release) the ongoing call */


    '/*-----------------------------------------------------------------------------
    '/* 					 CME Error Code and Standard Error Code for APP			 */
    '/*---------------------------------------------------------------------------*/
    '/* This CME ERROR Code is only for APP Reference. More Code reference to GSM Spec. */
    Private Const BTSDK_HFP_CMEERR_AGFAILURE As UInt32 = 0  '/* +CME ERROR:0 - AG failure */
    Private Const BTSDK_HFP_CMEERR_NOCONN2PHONE As UInt32 = 1  '/* +CME ERROR:1 - no connection to phone */
    Private Const BTSDK_HFP_CMEERR_OPERATION_NOTALLOWED As UInt32 = 3  '/* +CME ERROR:3 - operation not allowed */
    Private Const BTSDK_HFP_CMEERR_OPERATION_NOTSUPPORTED As UInt32 = 4  '/* +CME ERROR:4 - operation not supported */
    Private Const BTSDK_HFP_CMEERR_PHSIMPIN_REQUIRED As UInt32 = 5  '/* +CME ERROR:5 - PH-SIM PIN required */
    Private Const BTSDK_HFP_CMEERR_SIMNOT_INSERTED As UInt32 = 10 '/* +CME ERROR:10 - SIM not inserted */
    Private Const BTSDK_HFP_CMEERR_SIMPIN_REQUIRED As UInt32 = 11 '/* +CME ERROR:11 - SIM PIN required */
    Private Const BTSDK_HFP_CMEERR_SIMPUK_REQUIRED As UInt32 = 12 '/* +CME ERROR:12 - SIM PUK required */
    Private Const BTSDK_HFP_CMEERR_SIM_FAILURE As UInt32 = 13 '/* +CME ERROR:13 - SIM failure */
    Private Const BTSDK_HFP_CMEERR_SIM_BUSY As UInt32 = 14 '/* +CME ERROR:14 - SIM busy */
    Private Const BTSDK_HFP_CMEERR_INCORRECT_PASSWORD As UInt32 = 16 '/* +CME ERROR:16 - incorrect password */
    Private Const BTSDK_HFP_CMEERR_SIMPIN2_REQUIRED As UInt32 = 17 '/* +CME ERROR:17 - SIM PIN2 required */
    Private Const BTSDK_HFP_CMEERR_SIMPUK2_REQUIRED As UInt32 = 18 '/* +CME ERROR:18 - SIM PUK2 required */
    Private Const BTSDK_HFP_CMEERR_MEMORY_FULL As UInt32 = 20 '/* +CME ERROR:20 - memory full */
    Private Const BTSDK_HFP_CMEERR_INVALID_INDEX As UInt32 = 21 '/* +CME ERROR:21 - invalid index */
    Private Const BTSDK_HFP_CMEERR_MEMORY_FAILURE As UInt32 = 23 '/* +CME ERROR:23 - memory failure */
    Private Const BTSDK_HFP_CMEERR_TEXTSTRING_TOOLONG As UInt32 = 24 '/* +CME ERROR:24 - text string too long */
    Private Const BTSDK_HFP_CMEERR_INVALID_CHAR_INTEXTSTRING As UInt32 = 25 '/* +CME ERROR:25 - invalid characters in text string */
    Private Const BTSDK_HFP_CMEERR_DIAL_STRING_TOOLONG As UInt32 = 26 '/* +CME ERROR:26 - dial string too long */
    Private Const BTSDK_HFP_CMEERR_INVALID_CHAR_INDIALSTRING As UInt32 = 27 '/* +CME ERROR:27 - invalid characters in dial string */
    Private Const BTSDK_HFP_CMEERR_NETWORK_NOSERVICE As UInt32 = 30 '/* +CME ERROR:30 - no network service */
    Private Const BTSDK_HFP_CMEERR_NETWORK_TIMEOUT As UInt32 = 31 '/* +CME ERROR:31 - network timeout */
    Private Const BTSDK_HFP_CMEERR_EMERGENCYCALL_ONLY As UInt32 = 32 '/* +CME ERROR:32 - Network not allowed, emergency calls only */

    '/* APP specific error code. */
    Private Const BTSDK_HFP_APPERR_TIMEOUT As UInt32 = 200 '/* Wait for response timeout */

    '/* Standard error result code. */
    Private Const BTSDK_HFP_STDERR_ERROR As UInt32 = 201 '/* result code: ERROR */
    Private Const BTSDK_HFP_STDRR_NOCARRIER As UInt32 = 202 '/* result code: NO CARRIER */
    Private Const BTSDK_HFP_STDERR_BUSY As UInt32 = 203 '/* result code: BUSY */
    Private Const BTSDK_HFP_STDERR_NOANSWER As UInt32 = 204 '/* result code: NO ANSWER */
    Private Const BTSDK_HFP_STDERR_DELAYED As UInt32 = 205 '/* result code: DELAYED */
    Private Const BTSDK_HFP_STDERR_BLACKLISTED As UInt32 = 206 '/* result code: BLACKLISTED */
    Private Const BTSDK_HFP_OK As UInt32 = 255 '/* result code: OK */






    '/* HFP_SetState Callback to Application Event Co de */
    '/* SLC - Both AG and HF */
    Public Const BTSDK_HFP_EV_SLC_ESTABLISHED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 1                '/* HFP Service Level connection established. Parameter: public const BTSDK_HFP_ConnInfoStru */
    Public Const BTSDK_HFP_EV_SLC_RELEASED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 2                   '/* SPP connection released. Parameter: public const BTSDK_HFP_ConnInfoStru */

    '/* SCO - Both AG and HF  */
    Public Const BTSDK_HFP_EV_AUDIO_CONN_ESTABLISHED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 3      '/* SCO audio connection established */
    Public Const BTSDK_HFP_EV_AUDIO_CONN_RELEASED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 4         '/* SCO audio connection released */

    '/* Status Changed Indication */
    Public Const BTSDK_HFP_EV_STANDBY_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 5                       '/* STANDBY Menu, the incoming call or outgoing call or ongoing call is canceled  */
    Public Const BTSDK_HFP_EV_ONGOINGCALL_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 6                   '/* ONGOING-CALL Menu, a call (incoming call or outgoing call) is established (ongoing) */
    Public Const BTSDK_HFP_EV_RINGING_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 7                       '/* RINGING Menu, a call is incoming. Parameter: BTBOOL - in-band ring tone or not.   */
    Public Const BTSDK_HFP_EV_OUTGOINGCALL_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 8                  '/* OUTGOING-CALL Menu, an outgoing call is being established, 3Way in Guideline P91 */
    Public Const BTSDK_HFP_EV_CALLHELD_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 9                      '/* BTRH-HOLD Menu, +BTRH:0, AT+BTRH=0, incoming call is put on hold */
    Public Const BTSDK_HFP_EV_CALL_WAITING_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 10                 '/* Call Waiting Menu, +CCWA, When Call=Active, call waiting notification. Parameter: public const BTSDK_HFP_PhoneInfoStru */
    Public Const BTSDK_HFP_EV_TBUSY_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 11                        '/* GSM Network Remote Busy, TBusy Timer Activated */

    '/* AG & HF APP General Event Indication */
    Public Const BTSDK_HFP_EV_GENERATE_INBAND_RINGTONE_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 12      '/* AG Only, Generate the in-band ring tone */
    Public Const BTSDK_HFP_EV_TERMINATE_LOCAL_RINGTONE_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 13      '/* Terminate local generated ring tone or the in-band ring tone */
    Public Const BTSDK_HFP_EV_VOICE_RECOGN_ACTIVATED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 14        '/* +BVRA:1, voice recognition activated indication or HF request to start voice recognition procedure */
    Public Const BTSDK_HFP_EV_VOICE_RECOGN_DEACTIVATED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 15      '/* +BVRA:0, voice recognition deactivated indication or requests AG to deactivate the voice recognition procedure */
    Public Const BTSDK_HFP_EV_NETWORK_AVAILABLE_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 16             '/* +CIEV:<service><value>, cellular network is available */
    Public Const BTSDK_HFP_EV_NETWORK_UNAVAILABLE_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 17           '/* +CIEV:<service><value>, cellular network is unavailable */
    Public Const BTSDK_HFP_EV_ROAMING_RESET_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 18                 '/* +CIEV:<roam><value>, roaming is not active */
    Public Const BTSDK_HFP_EV_ROAMING_ACTIVE_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 19                '/* +CIEV:<roam><value>, a roaming is active */
    Public Const BTSDK_HFP_EV_SIGNAL_STRENGTH_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 20               '/* +CIEV:<signal><value>, signal strength indication. Parameter: BTUINT8 - indicator value */	
    Public Const BTSDK_HFP_EV_BATTERY_CHARGE_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 21                '/* +CIEV:<battchg><value>, battery charge indication. Parameter: BTUINT8 - indicator value  */
    Public Const BTSDK_HFP_EV_CHLDHELD_ACTIVATED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 22            '/* +CIEV:<callheld><1>, Call on CHLD Held to be or has been actived. */
    Public Const BTSDK_HFP_EV_CHLDHELD_RELEASED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 23             '/* +CIEV:<callheld><0>, Call on CHLD Held to be or has been released. */	
    Public Const BTSDK_HFP_EV_MICVOL_CHANGED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 24                '/* +VGM, AT+VGM, microphone volume changed indication */
    Public Const BTSDK_HFP_EV_SPKVOL_CHANGED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 25                '/* +VGS, AT+VGS, speaker volume changed indication */

    '/* OK and Error Code - HF only */
    Public Const BTSDK_HFP_EV_ATCMD_RESULT As UInt16 = BTSDK_APP_EV_AGAP_BASE + 26                      '/* HF Received OK, Error/+CME ERROR from AG or Wait for AG Response Timeout. Parameter: public const BTSDK_HFP_ATCmdResultStru */

    '/* To HF APP, Call Related, AG Send information to HF */
    Public Const BTSDK_HFP_EV_CLIP_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 27                          '/* +CLIP, Phone Number Indication. Parameter: public const BTSDK_HFP_PhoneInfoStru */
    Public Const BTSDK_HFP_EV_CURRENT_CALLS_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 28                 '/* +CLCC, the current calls of AG. Parameter: public const BTSDK_HFP_CLCCInfoStru */
    Public Const BTSDK_HFP_EV_NETWORK_OPERATOR_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 29              '/* +COPS, the current network operator name of AG. Parameter: public const BTSDK_HFP_COPSInfoStru */
    Public Const BTSDK_HFP_EV_SUBSCRIBER_NUMBER_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 30             '/* +CNUM, the subscriber number of AG. Parameter: public const BTSDK_HFP_PhoneInfoStru */
    Public Const BTSDK_HFP_EV_VOICETAG_PHONE_NUM_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 31            '/* +BINP, AG inputted phone number for voice-tag; requests AG to input a phone number for the voice-tag at the HF side. Parameter: public const BTSDK_HFP_PhoneInfoStru */

    '/* AG APP, HF Request or Indicate AG */
    Public Const BTSDK_HFP_EV_CURRENT_CALLS_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 32                 '/* AT+CLCC, query the list of current calls in AG. */
    Public Const BTSDK_HFP_EV_NETWORK_OPERATOR_FORMAT_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 33       '/* AT+COPS=3,0, indicate app the network operator name should be set to long alphanumeric */
    Public Const BTSDK_HFP_EV_NETWORK_OPERATOR_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 34              '/* AT+COPS?, requests AG to respond with +COPS response indicating the currently selected operator */
    Public Const BTSDK_HFP_EV_SUBSCRIBER_NUMBER_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 35             '/* AT+CNUM, query the AG subscriber number information. */
    Public Const BTSDK_HFP_EV_VOICETAG_PHONE_NUM_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 36            '/* AT+BINP, requests AG to input a phone number for the voice-tag at the HF */
    Public Const BTSDK_HFP_EV_CUR_INDICATOR_VAL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 37             '/* AT+CIND?, get the current indicator during the service level connection initialization procedure */
    Public Const BTSDK_HFP_EV_HF_DIAL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 38                       '/* ATD, instructs AG to dial the specific phone number. Parameter: (HFP only) BTUINT8* - phone number */
    Public Const BTSDK_HFP_EV_HF_MEM_DIAL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 39                   '/* ATD>, instructs AG to dial the phone number indexed by the specific memory location of SIM card. Parameter: BTUINT8* - memory location */
    Public Const BTSDK_HFP_EV_HF_LASTNUM_REDIAL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 40             '/* AT+BLDN, instructs AG to redial the last dialed phone number */
    Public Const BTSDK_HFP_EV_MANUFACTURER_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 41                  '/* AT+CGMI, requests AG to respond with the Manufacturer ID */
    Public Const BTSDK_HFP_EV_MODEL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 42                         '/* AT+CGMM, requests AG to respond with the Model ID */
    Public Const BTSDK_HFP_EV_NREC_DISABLE_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 43                  '/* AT+NREC=0, requests AG to disable NREC function */
    Public Const BTSDK_HFP_EV_DTMF_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 44                          '/* AT+VTS, instructs AG to transmit the specific DTMF code. Parameter: BTUINT8 - DTMF code */
    Public Const BTSDK_HFP_EV_ANSWER_CALL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 45                   '/* inform AG app to answer the call. Parameter: BTUINT8 - One of public const BTSDK_HFP_TYPE_INCOMING_CALL, public const BTSDK_HFP_TYPE_HELDINCOMING_CALL. */	
    Public Const BTSDK_HFP_EV_CANCEL_CALL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 46                   '/* inform AG app to cancel the call. Parameter: BTUINT8 - One of public const BTSDK_HFP_TYPE_ALL_CALLS, public const BTSDK_HFP_TYPE_INCOMING_CALL, public const BTSDK_HFP_TYPE_HELDINCOMING_CALL, public const BTSDK_HFP_TYPE_OUTGOING_CALL, public const BTSDK_HFP_TYPE_ONGOING_CALL. */	
    Public Const BTSDK_HFP_EV_HOLD_CALL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 47                     '/* inform AG app to hold the incoming call (AT+BTRH=0) */

    '/* AG APP, 3-Way Calling */
    Public Const BTSDK_HFP_EV_REJECTWAITINGCALL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 48             '/* AT+CHLD=0, Release all held calls or reject waiting call. */	
    Public Const BTSDK_HFP_EV_ACPTWAIT_RELEASEACTIVE_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 49        '/* AT+CHLD=1, Accept the held or waiting call and release all avtive calls. Parameter: BTUINT8 - value of <idx>*/
    Public Const BTSDK_HFP_EV_HOLDACTIVECALL_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 50                '/* AT+CHLD=2, Held Specified Active Call.  Parameter: BTUINT8 - value of <idx>*/
    Public Const BTSDK_HFP_EV_ADD_ONEHELDCALL_2ACTIVE_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 51       '/* AT+CHLD=3, Add One CHLD Held Call to Active Call. */
    Public Const BTSDK_HFP_EV_LEAVE3WAYCALLING_REQ As UInt16 = BTSDK_APP_EV_AGAP_BASE + 52              '/* AT+CHLD=4, Leave The 3-Way Calling. */

    '/* Extended */
    Public Const BTSDK_HFP_EV_EXTEND_CMD_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 53                    '/* indicate app extend command received. Parameter: BTUINT8* - Full extended AT command or result code. */
    Public Const BTSDK_HFP_EV_PRE_SCO_CONNECTION_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 54           '/* indicate app to create SCO connection. Parameter: public const BTSDK_AGAP_PreSCOConnIndStru. */
    Public Const BTSDK_HFP_EV_SPP_ESTABLISHED_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 55               '/* SPP connection created. Parameter: public const BTSDK_HFP_ConnInfoStru. added 2008-7-3 */
    Public Const BTSDK_HFP_EV_HF_MANUFACTURERID_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 56             '/* ManufacturerID indication. Parameter: BTUINT8* - Manufacturer ID of the AG device, a null-terminated ASCII string. */
    Public Const BTSDK_HFP_EV_HF_MODELID_IND As UInt16 = BTSDK_APP_EV_AGAP_BASE + 57                    '/* ModelID indication.  Parameter: BTUINT8* - Model ID of the AG device, a null-terminated ASCII string. */




    '/* AG Action Reason */
    Private Const BTSDK_HFP_CANCELED_ALLCALL As UInt32 = &H1    '/* AG released all calls or GSM Service Unavailable */
    Private Const BTSDK_HFP_CANCELED_CALLSETUP As UInt32 = &H2    '/* AG or GSM Release the Incoming Call or Outgoing Call */
    Private Const BTSDK_HFP_CANCELED_LASTCALL As UInt32 = &H3    '/* AG or GSM Release Last Call in Call=1 */

    Private Const BTSDK_HFP_AG_PRIVATE_MODE As UInt32 = &H5    '/* Answer the Outgoing/Incoming Call on AG */
    Private Const BTSDK_HFP_AG_HANDSFREE_MODE As UInt32 = &H6    '/* Answer the Outgoing/Incoming Call on HF */



    '/* Possible received events from GSM/CDMA cellular network */
    Private Const BTSDK_AGAP_NETWORK_RMT_IS_BUSY As UInt32 = &H1
    Private Const BTSDK_AGAP_NETWORK_ALERTING_RMT As UInt32 = &H2
    Private Const BTSDK_AGAP_NETWORK_INCOMING_CALL As UInt32 = &H3
    Private Const BTSDK_AGAP_NETWORK_RMT_ANSWER_CALL As UInt32 = &H4
    Private Const BTSDK_AGAP_NETWORK_SVC_UNAVAILABLE As UInt32 = &H5
    Private Const BTSDK_AGAP_NETWORK_SVC_AVAILABLE As UInt32 = &H6
    Private Const BTSDK_AGAP_NETWORK_SIGNAL_STRENGTH As UInt32 = &H7
    Private Const BTSDK_AGAP_NETWORK_ROAMING_RESET As UInt32 = &H8
    Private Const BTSDK_AGAP_NETWORK_ROAMING_ACTIVE As UInt32 = &H9

    Private Const BTSDK_HFP_CMD_GROUP1 As UInt32 = &H8000        '/* AT Command will response directly by OK */
    Private Const BTSDK_HFP_CMD_CHLD_0 As UInt32 = (BTSDK_HFP_CMD_GROUP1 Or &HB)                     '/* AT+CHLD=0 Held Call Release */
    Private Const BTSDK_HFP_CMD_CHLD_1 As UInt32 = (BTSDK_HFP_CMD_GROUP1 Or &HC)                     '/* AT+CHLD=1 Release Specified Active Call */
    Private Const BTSDK_HFP_CMD_CHLD_2 As UInt32 = (BTSDK_HFP_CMD_GROUP1 Or &HD)                     '/* AT+CHLD=2 Call Held or Active/Held Position Swap */
    Private Const BTSDK_HFP_CMD_CHLD_3 As UInt32 = (BTSDK_HFP_CMD_GROUP1 Or &HE)                     '/* AT+CHLD=3 Adds a held call to the conversation */
    Private Const BTSDK_HFP_CMD_CHLD_4 As UInt32 = (BTSDK_HFP_CMD_GROUP1 Or &HF)                     '/* AT+CHLD=4 Connects the two calls and disconnects the subscriber from both calls */

    '/* HF Device state*/
    Public Const BTSDK_HFAP_ST_IDLE As UInt32 = &H1     '/*before service level connection is established*/
    Public Const BTSDK_HFAP_ST_STANDBY As UInt32 = &H2     '/*service level connection is established*/
    Public Const BTSDK_HFAP_ST_RINGING As UInt32 = &H3     '/*ringing*/
    Public Const BTSDK_HFAP_ST_OUTGOINGCALL As UInt32 = &H4     '/*outgoing call*/
    Public Const BTSDK_HFAP_ST_ONGOINGCALL As UInt32 = &H5     '/*ongoing call*/
    Public Const BTSDK_HFAP_ST_BVRA As UInt32 = &H6     '/*voice recognition is ongoing*/
    Public Const BTSDK_HFAP_ST_VOVG As UInt32 = &H7
    Public Const BTSDK_HFAP_ST_HELDINCOMINGCALL As UInt32 = &H8 '/*the incoming call is held*/




    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AGAP_IsAudioConnExisted(ByRef connBool As Byte) As UInt32        'used to figure if call audio is on pc or phone.
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_GetSubscriberNumber(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_NetworkOperatorReq(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AGAP_GetAGState(ByRef agState As UInt16) As UInt32
    End Function



    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_APPRegCbk4ThirdParty(ByVal ptrFunc_HFPevent As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AGAP_APPRegCbk4ThirdParty(ByVal ptrFunc_HFPevent As UInt32) As UInt32
    End Function





    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_Dial(ByVal connHandle As UInt32, ByRef arrayPhoneNumberBytes As Byte, ByVal arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_LastNumRedial(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_CancelCall(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_AnswerCall(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_SetMicVol(ByVal connHandle As UInt32, ByVal micVol0to15 As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AGAP_SetMicVol(ByVal connHandle As UInt32, ByVal micVol0to15 As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_SetSpkVol(ByVal connHandle As UInt32, ByVal spkVol0to15 As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AGAP_SetSpkVol(ByVal connHandle As UInt32, ByVal spkVol0to15 As Byte) As UInt32
    End Function



    'Btsdk_AGAP_SetSpkVol

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_GetCurrHFState(ByRef agState As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_GetManufacturerID(ByVal connHandle As UInt32, ByRef array256bytes As Byte, ByRef arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_GetModelID(ByVal connHandle As UInt32, ByRef array256bytes As Byte, ByRef arrayLen As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_TxDTMF(ByVal connHandle As UInt32, ByVal dtmfChar As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_VoiceRecognitionReq(ByVal connHandle As UInt32, ByVal ONorOFF As Byte) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_SetWaveInDevice(ByRef arrayBytes_DvcName As Byte, ByVal arrayLen As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_SetWaveOutDevice(ByRef arrayBytes_DvcName As Byte, ByVal arrayLen As UInt32) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFP_ExtendCmd(ByVal connHandle As UInt32, ByRef array256bytes As Byte, ByVal arrayLen As UInt16, ByVal timeoutSecs As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_HFAP_AudioConnTrans(ByVal connHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_RegisterHFPService(ByRef array80bytes As Byte, ByVal svcClass As UInt16, ByVal svcFeatures As UInt16) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_UnregisterHFPService(ByVal svcHandle As UInt32) As UInt32
    End Function



    Private Function BlueSoleil_HFP_Callback_Status_AG_Add(ByVal ptrFunc_HFPEventCallback_AG As IntPtr) As Boolean

        Dim retBool As Boolean = False

        Dim retUInt32 As UInt32

        If ptrFunc_HFPEventCallback_AG = IntPtr.Zero Then

            retUInt32 = Btsdk_AGAP_APPRegCbk4ThirdParty(CUInt(ptrFunc_HFPEventCallback_AG))
        Else
            retUInt32 = Btsdk_AGAP_APPRegCbk4ThirdParty(CUInt(ptrFunc_HFPEventCallback_AG))
        End If



        If retUInt32 = BTSDK_OK Then
            retBool = True
        Else
            retBool = False
        End If

        Return retBool

    End Function



    Private Function BlueSoleil_HFP_Callback_Status_AP_Add(ByVal ptrFunc_HFPEventCallback_AP As IntPtr) As Boolean

        Dim retBool As Boolean = False

        Dim retUInt32 As UInt32

        If ptrFunc_HFPEventCallback_AP = IntPtr.Zero Then

            retUInt32 = Btsdk_HFAP_APPRegCbk4ThirdParty(CUInt(ptrFunc_HFPEventCallback_AP))
        Else
            retUInt32 = Btsdk_HFAP_APPRegCbk4ThirdParty(CUInt(ptrFunc_HFPEventCallback_AP))
        End If



        If retUInt32 = BTSDK_OK Then
            retBool = True
        Else
            retBool = False
        End If

        Return retBool

    End Function

    Private Sub BlueSoleil_HFP_Callback_Status_AP_Func_ParseCOPSinfoStru(ByVal ptrStruPhoneInfo As IntPtr, ByVal paramLen As UInt16, ByRef retName As String)

        'expecting 3+ bytes.  (or maybe more?)   '1 byte mode, 1 byte format, 1 byte name-length, name

        '3 bytes, plus the value of ptrStruPhoneInfo(2)

        retName = ""

        Dim evtData(0 To 2) As Byte
        Marshal.Copy(ptrStruPhoneInfo, evtData, 0, evtData.Length)

        Dim nameLen As Byte = evtData(2)

        Debug.Print("NameLen = " & nameLen)

        If nameLen > 0 Then
            ReDim evtData(0 To 2 + nameLen)
            Marshal.Copy(ptrStruPhoneInfo, evtData, 0, evtData.Length)
            retName = System.Text.Encoding.UTF8.GetString(evtData, 3, nameLen)

        End If

    End Sub


    Private Sub BlueSoleil_HFP_Callback_Status_AP_Func_ParseCLCCinfoStru(ByVal ptrStruCLCCinfo As IntPtr, ByVal paramLen As UInt16, ByRef clccPhoneNumber As String, ByRef clccIsIncoming As Boolean, ByRef clccIsFax As Boolean, ByRef clccIsData As Boolean, ByRef clccIsMultiParty As Boolean)

        'expecting 7+ bytes.  (or maybe more?)   '1 byte idx, 1 byte direction, 1 byte status, 1 byte mode, 1 byte multiparty, 1 byte type, 1 byte number-length, phone number

        '7 bytes, plus the value of ptrStruCLCCinfo(6)

        clccPhoneNumber = ""

        Dim evtData(0 To 6) As Byte
        Marshal.Copy(ptrStruCLCCinfo, evtData, 0, evtData.Length)

        clccIsIncoming = (evtData(1) = 1)
        clccIsData = (evtData(3) = 1)
        clccIsFax = (evtData(3) = 2)
        clccIsMultiParty = (evtData(4) = 1)

        Dim numLen As Byte = evtData(6)

        Debug.Print("NumLen = " & numLen)

        If numLen > 0 Then
            ReDim evtData(0 To 6 + numLen)
            Marshal.Copy(ptrStruCLCCinfo, evtData, 0, evtData.Length)
            clccPhoneNumber = System.Text.Encoding.UTF8.GetString(evtData, 7, numLen)

        End If

    End Sub


    Private Sub BlueSoleil_HFP_Callback_Status_AP_Func_ParseExtCmdInd(ByVal ptrExtCmdInd As IntPtr, ByVal paramLen As UInt16)

        Dim cmdResponse As String = ""
        cmdResponse = Marshal.PtrToStringAnsi(ptrExtCmdInd, paramLen)

        Debug.Print(cmdResponse)

        'parse the string.
        'if CBC, fire batterycharge event.  if CSQ, fire signalquality event, etc.

        Dim plusPos As Integer = InStr(1, cmdResponse, "+")
        Dim colonPos As Integer = InStr(plusPos + 1, cmdResponse, ":")
        Dim cmdType As String = Mid(cmdResponse, plusPos + 1, colonPos - plusPos - 1)

        Select Case cmdType
            Case "CBC"      'plug-state, pct

            Case "CSQ"      'sig quality, range 0-31, or 99 if unavailable or not detected.

            Case "CCLK"

            Case "CIND"     'supported indicator types

            Case "CGSN"

            Case "CMER"

            Case "CGMI"

            Case "CGMM"

            Case "CGMR"

            Case "CGSN"

            Case "CLIP"

            Case "COPN"

            Case "COPS"     ' , , network operator name 

            Case "CGREG"    'registed to home network?

            Case "CGATT"    'attached to network?

            Case "CSCA"     'sms service address

            Case "CGPSINF"  'probably never gonna happen.

        End Select

        'RaiseEvent BlueSoleil_Event_HFP_ExtCmdInd(cmdResultStr)
        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_ExtCmdInd(cmdResponse))
        t.Start()

    End Sub

    Private Sub BlueSoleil_HFP_Callback_Status_AP_Func_ParsePhoneInfoStru(ByVal ptrStruPhoneInfo As IntPtr, ByVal paramLen As UInt16, ByRef retPhoneNum As String, ByRef retName As String)

        'expecting 36+ bytes.   '1 byte type, 1 byte service, 1 byte numLen, 32 byte phone number, 1 byte name len, name

        '36 bytes, plus the value of ptrStruPhoneInfo(35)

        '
        retPhoneNum = ""
        retName = ""

        Dim evtData(0 To 35) As Byte
        Marshal.Copy(ptrStruPhoneInfo, evtData, 0, evtData.Length)

        Dim piType As Byte = evtData(0)         'format of phone number.
        Dim piService As Byte = evtData(1)      '4 = voice, 5 = fax.
        Dim numLen As Byte = evtData(2)

        Dim nameLen As Byte = evtData(35)

        Debug.Print("PhoneInfoStru - NumLen = " & numLen & "   NameLen = " & nameLen)

        If numLen = 0 Then numLen = 32
        If numLen > 32 Then numLen = 32

        Dim phoneNumBytes(0 To 0) As Byte
        If numLen > 0 Then
            retPhoneNum = System.Text.Encoding.UTF8.GetString(evtData, 3, numLen)
        End If

        If nameLen > 0 Then
            ReDim evtData(0 To 35 + nameLen)
            Marshal.Copy(ptrStruPhoneInfo, evtData, 0, evtData.Length)
            retName = System.Text.Encoding.UTF8.GetString(evtData, 36, nameLen)

        End If


    End Sub

    Private Sub BlueSoleil_HFP_Callback_Status_AG_Func(ByVal connHandle As UInt32, ByVal hfpEvent As UInt16, ByVal funcParam As IntPtr, ByVal paramLen As UInt16)



        Debug.Print("HFP AG Callback  evt = " & hfpEvent & "  paramLen = " & paramLen)

        Return

        Select Case hfpEvent

            Case BTSDK_HFP_EV_EXTEND_CMD_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_EXTEND_CMD_IND")
                If paramLen > 0 Then

                    BlueSoleil_HFP_Callback_Status_AP_Func_ParseExtCmdInd(funcParam, paramLen)



                End If

            Case BTSDK_HFP_EV_SIGNAL_STRENGTH_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_SIGNAL_STRENGTH_IND")
                Dim tempSignal As Byte
                tempSignal = Marshal.ReadByte(funcParam)
                If tempSignal < 6 Then
                    BlueSoleil_HFP_LastSignalVal = tempSignal
                    Dim currSignalPct As Double = tempSignal * 20

                    'RaiseEvent BlueSoleil_Event_HFP_SignalQuality(currSignalPct)
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_SignalQuality(currSignalPct))
                    t.Start()
                Else
                    tempSignal = tempSignal
                End If


            Case BTSDK_HFP_EV_BATTERY_CHARGE_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_BATTERY_CHARGE_IND")
                Dim tempBattery As Byte
                tempBattery = Marshal.ReadByte(funcParam)
                Debug.Print("Byte = " & tempBattery)
                If tempBattery < 6 Then
                    BlueSoleil_HFP_LastBatteryVal = tempBattery
                    Dim currBatteryPct As Double = tempBattery * 20

                    'RaiseEvent BlueSoleil_Event_HFP_BatteryCharge(currBatteryPct)
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_BatteryCharge(currBatteryPct))
                    t.Start()
                Else
                    tempBattery = tempBattery
                End If



            Case BTSDK_HFP_EV_SLC_RELEASED_IND
                'HFP connection released?!
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_SLC_RELEASED_IND")
                'RaiseEvent BlueSoleil_Event_HFP_ConnectionReleased()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_ConnectionReleased())
                t.Start()


            Case BTSDK_HFP_EV_AUDIO_CONN_ESTABLISHED_IND
                'use this to get conn handle to SCO connection.


            Case BTSDK_HFP_EV_AUDIO_CONN_RELEASED_IND
                'sco conn released.


            Case BTSDK_HFP_EV_VOICE_RECOGN_ACTIVATED_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_VOICE_RECOGN_ACTIVATED_IND")
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_VoiceCmdActivated())
                t.Start()

            Case BTSDK_HFP_EV_VOICE_RECOGN_DEACTIVATED_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_VOICE_RECOGN_DEACTIVATED_IND")
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_VoiceCmdDeactivated())
                t.Start()



            Case BTSDK_HFP_EV_ROAMING_ACTIVE_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_ROAMING_ACTIVE_IND")
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_StartRoaming())
                t.Start()

            Case BTSDK_HFP_EV_ROAMING_RESET_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_ROAMING_RESET_IND")
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_StopRoaming())
                t.Start()





            Case BTSDK_HFP_EV_CLIP_IND
                'caller id
                'param = Btsdk_HFP_PhoneInfoStru
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_CLIP_IND")
                Dim clipPhoneNo As String = ""
                Dim clipName As String = ""
                BlueSoleil_HFP_Callback_Status_AP_Func_ParsePhoneInfoStru(funcParam, paramLen, clipPhoneNo, clipName)
                'RaiseEvent Bluesoleil_Event_HFP_CallerID(clipPhoneNo, clipName)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_CallerID(clipPhoneNo, clipName))
                t.Start()

            Case BTSDK_HFP_EV_SUBSCRIBER_NUMBER_IND
                'local phone no.
                'param = Btsdk_HFP_PhoneInfoStru
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_SUBSCRIBER_NUMBER_IND")
                Dim subscriberPhoneNo As String = ""
                Dim subscriberName As String = ""
                BlueSoleil_HFP_Callback_Status_AP_Func_ParsePhoneInfoStru(funcParam, paramLen, subscriberPhoneNo, subscriberName)
                'RaiseEvent Bluesoleil_Event_HFP_SubscriberPhoneNo(subscriberPhoneNo, subscriberName)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_SubscriberPhoneNo(subscriberPhoneNo, subscriberName))
                t.Start()

            Case BTSDK_HFP_EV_NETWORK_AVAILABLE_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_NETWORK_AVAILABLE_IND")
                'RaiseEvent BlueSoleil_Event_HFP_NetworkAvailable()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_NetworkAvailable())
                t.Start()


            Case BTSDK_HFP_EV_NETWORK_UNAVAILABLE_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_NETWORK_UNAVAILABLE_IND")
                'RaiseEvent BlueSoleil_Event_HFP_NetworkUnavailable()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_NetworkUnavailable())
                t.Start()

            Case BTSDK_HFP_EV_NETWORK_OPERATOR_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_NETWORK_OPERATOR_IND")
                Dim netopName As String = ""
                BlueSoleil_HFP_Callback_Status_AP_Func_ParseCOPSinfoStru(funcParam, paramLen, netopName)
                'RaiseEvent BlueSoleil_Event_HFP_NetworkOperatorName(netopName)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_NetworkOperatorName(netopName))
                t.Start()

            Case BTSDK_HFP_EV_MICVOL_CHANGED_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_MICVOL_CHANGED_IND")
                Dim tempVol As Byte
                Dim tempVolPct As Double
                tempVol = Marshal.ReadByte(funcParam)
                tempVol = tempVol
                tempVolPct = 100 * 15 / tempVol
                If tempVolPct < 0 Then tempVolPct = 0
                If tempVolPct > 100 Then tempVolPct = 100
                'RaiseEvent BlueSoleil_Event_HFP_MicVolume(tempVolPct)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_MicVolume(tempVolPct))
                t.Start()

            Case BTSDK_HFP_EV_SPKVOL_CHANGED_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_SPKVOL_CHANGED_IND")
                Dim tempVol As Byte
                Dim tempVolPct As Double
                tempVol = Marshal.ReadByte(funcParam)
                tempVol = tempVol
                tempVolPct = 100 * 15 / tempVol
                If tempVolPct < 0 Then tempVolPct = 0
                If tempVolPct > 100 Then tempVolPct = 100
                'RaiseEvent BlueSoleil_Event_HFP_SpeakerVolume(tempVolPct)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_SpeakerVolume(tempVolPct))
                t.Start()

            Case BTSDK_HFP_EV_HF_MODELID_IND

                If paramLen > 0 Then
                    Dim modelStr As String = ""
                    Dim modelBytes(0 To paramLen - 1) As Byte
                    Marshal.Copy(funcParam, modelBytes, 0, modelBytes.Length)
                    modelStr = System.Text.Encoding.UTF8.GetString(modelBytes)
                    'RaiseEvent BlueSoleil_Event_HFP_ModelName(modelStr)
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_ModelName(modelStr))
                    t.Start()
                End If

            Case BTSDK_HFP_EV_HF_MANUFACTURERID_IND

                If paramLen > 0 Then
                    Dim manuStr As String = ""
                    Dim manuBytes(0 To paramLen - 1) As Byte
                    Marshal.Copy(funcParam, manuBytes, 0, manuBytes.Length)
                    manuStr = System.Text.Encoding.UTF8.GetString(manuBytes)
                    'RaiseEvent BlueSoleil_Event_HFP_ManufacturerName(manuStr)
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_ManufacturerName(manuStr))
                    t.Start()
                End If


            Case BTSDK_HFP_EV_STANDBY_IND
                'RaiseEvent BlueSoleil_Event_HFP_Standby()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_Standby())
                t.Start()

            Case BTSDK_HFP_EV_ONGOINGCALL_IND
                'RaiseEvent BlueSoleil_Event_HFP_OngoingCall()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_OngoingCall())
                t.Start()

            Case BTSDK_HFP_EV_RINGING_IND
                'RaiseEvent BlueSoleil_Event_HFP_Ringing()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_Ringing())
                t.Start()

            Case BTSDK_HFP_EV_OUTGOINGCALL_IND
                'RaiseEvent BlueSoleil_Event_HFP_OutgoingCall()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_OutgoingCall())
                t.Start()

            Case BTSDK_HFP_EV_ATCMD_RESULT
                'param = Btsdk_HFP_ATCmdResultStru
                If paramLen > 0 Then
                    Dim cmdResultStr As String = ""
                    ' Dim modelBytes(0 To paramLen - 1) As Byte
                    ' Marshal.Copy(funcParam, modelBytes, 0, modelBytes.Length)
                    ' cmdResultStr = System.Text.Encoding.UTF8.GetString(modelBytes)
                    ' Debug.Print("BTSDK_HFP_EV_ATCMD_RESULT = " & cmdResultStr)
                End If


            Case Else
                ' MsgBox(connHandle)

        End Select


        Return 'BTSDK_OK

    End Sub




    Private Sub BlueSoleil_HFP_Callback_Status_AP_Func(ByVal connHandle As UInt32, ByVal hfpEvent As UInt16, ByVal funcParam As IntPtr, ByVal paramLen As UInt16)

        Debug.Print("HFP AP Callback  evt = " & hfpEvent & "  paramLen = " & paramLen)



        Select Case hfpEvent

            Case BTSDK_HFP_EV_EXTEND_CMD_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_EXTEND_CMD_IND")
                If paramLen > 0 Then

                    BlueSoleil_HFP_Callback_Status_AP_Func_ParseExtCmdInd(funcParam, paramLen)


                End If

            Case BTSDK_HFP_EV_SIGNAL_STRENGTH_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_SIGNAL_STRENGTH_IND")
                Dim tempSignal As Byte
                tempSignal = Marshal.ReadByte(funcParam)
                If tempSignal < 6 Then
                    BlueSoleil_HFP_LastSignalVal = tempSignal
                    Dim currSignalPct As Double = tempSignal * 20

                    'RaiseEvent BlueSoleil_Event_HFP_SignalQuality(currSignalPct)
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_SignalQuality(currSignalPct))
                    t.Start()
                Else
                    tempSignal = tempSignal
                End If


            Case BTSDK_HFP_EV_BATTERY_CHARGE_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_BATTERY_CHARGE_IND")
                Dim tempBattery As Byte
                tempBattery = Marshal.ReadByte(funcParam)
                If tempBattery < 6 Then
                    BlueSoleil_HFP_LastBatteryVal = tempBattery
                    Dim currBatteryPct As Double = tempBattery * 20

                    'RaiseEvent BlueSoleil_Event_HFP_BatteryCharge(currBatteryPct)
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_BatteryCharge(currBatteryPct))
                    t.Start()
                Else
                    tempBattery = tempBattery
                End If



            Case BTSDK_HFP_EV_SLC_RELEASED_IND
                'HFP connection released?!
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_SLC_RELEASED_IND")
                'RaiseEvent BlueSoleil_Event_HFP_ConnectionReleased()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_ConnectionReleased())
                t.Start()


            Case BTSDK_HFP_EV_AUDIO_CONN_ESTABLISHED_IND
                'use this to get conn handle to SCO connection.


            Case BTSDK_HFP_EV_AUDIO_CONN_RELEASED_IND
                'sco conn released.


            Case BTSDK_HFP_EV_VOICE_RECOGN_ACTIVATED_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_VOICE_RECOGN_ACTIVATED_IND")
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_VoiceCmdActivated())
                t.Start()

            Case BTSDK_HFP_EV_VOICE_RECOGN_DEACTIVATED_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_VOICE_RECOGN_DEACTIVATED_IND")
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_VoiceCmdDeactivated())
                t.Start()



            Case BTSDK_HFP_EV_ROAMING_ACTIVE_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_ROAMING_ACTIVE_IND")
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_StartRoaming())
                t.Start()

            Case BTSDK_HFP_EV_ROAMING_RESET_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_ROAMING_RESET_IND")
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_StopRoaming())
                t.Start()


            Case BTSDK_HFP_EV_CURRENT_CALLS_IND
                'current calls
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_CURRENT_CALLS_IND")
                Dim clccPhoneNo As String = ""
                Dim clccIsIncoming As Boolean = False
                Dim clccIsFax As Boolean = False
                Dim clccIsData As Boolean = False
                Dim clccIsMultiParty As Boolean = False
                BlueSoleil_HFP_Callback_Status_AP_Func_ParseCLCCinfoStru(funcParam, paramLen, clccPhoneNo, clccIsIncoming, clccIsFax, clccIsData, clccIsMultiParty)
                'RaiseEvent Bluesoleil_Event_HFP_CallerID(clipPhoneNo, clipName)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_CurrentCallInfo(clccPhoneNo, clccIsIncoming, clccIsFax, clccIsData, clccIsMultiParty))
                t.Start()



            Case BTSDK_HFP_EV_CLIP_IND
                'caller id
                'param = Btsdk_HFP_PhoneInfoStru
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_CLIP_IND")
                Dim clipPhoneNo As String = ""
                Dim clipName As String = ""
                BlueSoleil_HFP_Callback_Status_AP_Func_ParsePhoneInfoStru(funcParam, paramLen, clipPhoneNo, clipName)
                'RaiseEvent Bluesoleil_Event_HFP_CallerID(clipPhoneNo, clipName)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_CallerID(clipPhoneNo, clipName))
                t.Start()

            Case BTSDK_HFP_EV_SUBSCRIBER_NUMBER_IND
                'local phone no.
                'param = Btsdk_HFP_PhoneInfoStru
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_SUBSCRIBER_NUMBER_IND")
                Dim subscriberPhoneNo As String = ""
                Dim subscriberName As String = ""
                BlueSoleil_HFP_Callback_Status_AP_Func_ParsePhoneInfoStru(funcParam, paramLen, subscriberPhoneNo, subscriberName)
                'RaiseEvent Bluesoleil_Event_HFP_SubscriberPhoneNo(subscriberPhoneNo, subscriberName)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_SubscriberPhoneNo(subscriberPhoneNo, subscriberName))
                t.Start()

            Case BTSDK_HFP_EV_NETWORK_AVAILABLE_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_NETWORK_AVAILABLE_IND")
                'RaiseEvent BlueSoleil_Event_HFP_NetworkAvailable()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_NetworkAvailable())
                t.Start()


            Case BTSDK_HFP_EV_NETWORK_UNAVAILABLE_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_NETWORK_UNAVAILABLE_IND")
                'RaiseEvent BlueSoleil_Event_HFP_NetworkUnavailable()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_NetworkUnavailable())
                t.Start()

            Case BTSDK_HFP_EV_NETWORK_OPERATOR_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_NETWORK_OPERATOR_IND")
                Dim netopName As String = ""
                BlueSoleil_HFP_Callback_Status_AP_Func_ParseCOPSinfoStru(funcParam, paramLen, netopName)
                'RaiseEvent BlueSoleil_Event_HFP_NetworkOperatorName(netopName)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_NetworkOperatorName(netopName))
                t.Start()

            Case BTSDK_HFP_EV_MICVOL_CHANGED_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_MICVOL_CHANGED_IND")
                Dim tempVol As Byte
                Dim tempVolPct As Double
                tempVol = Marshal.ReadByte(funcParam)
                tempVol = tempVol
                tempVolPct = 100 * 15 / tempVol
                If tempVolPct < 0 Then tempVolPct = 0
                If tempVolPct > 100 Then tempVolPct = 100
                'RaiseEvent BlueSoleil_Event_HFP_MicVolume(tempVolPct)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_MicVolume(tempVolPct))
                t.Start()

            Case BTSDK_HFP_EV_SPKVOL_CHANGED_IND
                Debug.Print("HFP Callback  evt = BTSDK_HFP_EV_SPKVOL_CHANGED_IND")
                Dim tempVol As Byte
                Dim tempVolPct As Double
                tempVol = Marshal.ReadByte(funcParam)
                tempVol = tempVol
                tempVolPct = 100 * 15 / tempVol
                If tempVolPct < 0 Then tempVolPct = 0
                If tempVolPct > 100 Then tempVolPct = 100
                'RaiseEvent BlueSoleil_Event_HFP_SpeakerVolume(tempVolPct)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_SpeakerVolume(tempVolPct))
                t.Start()

            Case BTSDK_HFP_EV_HF_MODELID_IND

                If paramLen > 0 Then
                    Dim modelStr As String = ""
                    Dim modelBytes(0 To paramLen - 1) As Byte
                    Marshal.Copy(funcParam, modelBytes, 0, modelBytes.Length)
                    modelStr = System.Text.Encoding.UTF8.GetString(modelBytes)
                    'RaiseEvent BlueSoleil_Event_HFP_ModelName(modelStr)
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_ModelName(modelStr))
                    t.Start()
                End If

            Case BTSDK_HFP_EV_HF_MANUFACTURERID_IND
                If paramLen > 0 Then
                    Dim manuStr As String = ""
                    Dim manuBytes(0 To paramLen - 1) As Byte
                    Marshal.Copy(funcParam, manuBytes, 0, manuBytes.Length)
                    manuStr = System.Text.Encoding.UTF8.GetString(manuBytes)
                    'RaiseEvent BlueSoleil_Event_HFP_ManufacturerName(manuStr)
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_ManufacturerName(manuStr))
                    t.Start()
                End If


            Case BTSDK_HFP_EV_STANDBY_IND
                'RaiseEvent BlueSoleil_Event_HFP_Standby()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_Standby())
                t.Start()

            Case BTSDK_HFP_EV_ONGOINGCALL_IND
                'RaiseEvent BlueSoleil_Event_HFP_OngoingCall()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_OngoingCall())
                t.Start()

            Case BTSDK_HFP_EV_RINGING_IND
                'RaiseEvent BlueSoleil_Event_HFP_Ringing()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_Ringing())
                t.Start()

            Case BTSDK_HFP_EV_OUTGOINGCALL_IND
                'RaiseEvent BlueSoleil_Event_HFP_OutgoingCall()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_HFP_OutgoingCall())
                t.Start()

            Case BTSDK_HFP_EV_ATCMD_RESULT
                'param = Btsdk_HFP_ATCmdResultStru
                If paramLen > 0 Then
                    Dim cmdResultStr As String = ""
                    ' Dim modelBytes(0 To paramLen - 1) As Byte
                    ' Marshal.Copy(funcParam, modelBytes, 0, modelBytes.Length)
                    ' cmdResultStr = System.Text.Encoding.UTF8.GetString(modelBytes)
                    ' Debug.Print("BTSDK_HFP_EV_ATCMD_RESULT = " & cmdResultStr)
                End If


            Case Else
                ' MsgBox(connHandle)

        End Select


        Return 'BTSDK_OK

    End Sub

    Public Sub BlueSoleil_HFP_RegisterCallbacks()

        ' Dim functPtrAG As IntPtr = Marshal.GetFunctionPointerForDelegate(delegateHFPeventAGCallback)
        ' BlueSoleil_HFP_Callback_Status_AG_Add(functPtrAG)

        Dim functPtrAP As IntPtr = Marshal.GetFunctionPointerForDelegate(delegateHFPeventAPCallback)
        BlueSoleil_HFP_Callback_Status_AP_Add(functPtrAP)

    End Sub



    Public Sub BlueSoleil_HFP_UnregisterCallbacks()

        Dim functPtr As IntPtr = IntPtr.Zero

        ' BlueSoleil_HFP_Callback_Status_AG_Add(functPtr)

        BlueSoleil_HFP_Callback_Status_AP_Add(functPtr)

    End Sub




    Public Function BlueSoleil_HFP_RegisterService_HandsFreeUnit(ByVal svcName As String) As UInt32

        Dim svcNameBytes(0 To BTSDK_SERVICENAME_MAXLENGTH - 1) As Byte

        Dim tempBytes(0 To 0) As Byte
        tempBytes = System.Text.Encoding.UTF8.GetBytes(svcName)
        If tempBytes.Length > svcNameBytes.Length Then
            Return 0
        End If
        Array.Copy(tempBytes, 0, svcNameBytes, 0, tempBytes.Length)

        Dim flags As UInt16 = BTSDK_HF_BRSF_3WAYCALL Or BTSDK_HF_BRSF_CLIP Or BTSDK_HF_BRSF_BVRA Or BTSDK_HF_BRSF_RMTVOLCTRL Or BTSDK_HF_BRSF_ENHANCED_CALLSTATUS Or BTSDK_HF_BRSF_ENHANCED_CALLCONTROL

        flags = BTSDK_HF_BRSF_ALL


        Dim retUInt32 As UInt32 = Btsdk_RegisterHFPService(svcNameBytes(0), BTSDK_CLS_HANDSFREE, flags)

        Return retUInt32

    End Function


    Public Function BlueSoleil_HFP_RegisterService_HandsFreeAudioGateway(ByVal svcName As String) As UInt32

        Dim svcNameBytes(0 To BTSDK_SERVICENAME_MAXLENGTH - 1) As Byte

        Dim tempBytes(0 To 0) As Byte
        tempBytes = System.Text.Encoding.UTF8.GetBytes(svcName)
        If tempBytes.Length > svcNameBytes.Length Then
            Return 0
        End If
        Array.Copy(tempBytes, 0, svcNameBytes, 0, tempBytes.Length)

        Dim flags As UInt16 = BTSDK_AG_BRSF_3WAYCALL Or BTSDK_AG_BRSF_BVRA Or BTSDK_AG_BRSF_BINP Or BTSDK_AG_BRSF_REJECT_CALL Or BTSDK_AG_BRSF_ENHANCED_CALLSTATUS Or BTSDK_AG_BRSF_ENHANCED_CALLCONTROL Or BTSDK_AG_BRSF_EXTENDED_ERRORRESULT

        flags = BTSDK_AG_BRSF_ALL

        Dim retUInt32 As UInt32 = Btsdk_RegisterHFPService(svcNameBytes(0), BTSDK_CLS_HANDSFREE_AG, flags)

        Return retUInt32

    End Function




    Public Function BlueSoleil_HFP_RegisterService_HeadSetUnit(ByVal svcName As String) As UInt32

        Dim svcNameBytes(0 To BTSDK_SERVICENAME_MAXLENGTH - 1) As Byte

        Dim tempBytes(0 To 0) As Byte
        tempBytes = System.Text.Encoding.UTF8.GetBytes(svcName)
        If tempBytes.Length > svcNameBytes.Length Then
            Return 0
        End If
        Array.Copy(tempBytes, 0, svcNameBytes, 0, tempBytes.Length)


        Dim retUInt32 As UInt32 = Btsdk_RegisterHFPService(svcNameBytes(0), BTSDK_CLS_HEADSET, 0)

        Return retUInt32

    End Function

    Public Function BlueSoleil_HFP_RegisterService_HeadSetAudioGateway(ByVal svcName As String) As UInt32

        Dim svcNameBytes(0 To BTSDK_SERVICENAME_MAXLENGTH - 1) As Byte

        Dim tempBytes(0 To 0) As Byte
        tempBytes = System.Text.Encoding.UTF8.GetBytes(svcName & Chr(0))
        If tempBytes.Length > svcNameBytes.Length Then
            Return 0
        End If
        Array.Copy(tempBytes, 0, svcNameBytes, 0, tempBytes.Length)


        Dim retUInt32 As UInt32 = Btsdk_RegisterHFPService(svcNameBytes(0), BTSDK_CLS_HEADSET_AG, 0)

        Return retUInt32

    End Function

    Public Function BlueSoleil_HFP_UnregisterService(ByVal svcHandle As UInt32) As Boolean

        If svcHandle = 0 Then Return True

        Dim retUInt32 As UInt32 = Btsdk_UnregisterHFPService(svcHandle)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function





    Public Function BlueSoleil_HFP_SendATcmd(ByVal connHandle As UInt32, ByVal atCmdString As String, Optional ByVal timeoutMilliSec As UInteger = 0) As Boolean

        If connHandle = 0 Then Return False


        If atCmdString = "" Then Return False

        If Strings.Right(atCmdString, 1) <> vbCr Then atCmdString = atCmdString & vbCr

        Dim atCmdStringLen As UShort = CUShort(atCmdString.Length)

        Dim atCmdBytes(0 To 255) As Byte

        Dim tempBytes(0 To 0) As Byte
        tempBytes = System.Text.Encoding.UTF8.GetBytes(atCmdString & Chr(0))
        If tempBytes.Length > atCmdBytes.Length Then
            Return False
        End If
        '  Array.Copy(tempBytes, 0, atCmdBytes, 0, tempBytes.Length)



        Dim retUInt32 As UInt32 = Btsdk_HFP_ExtendCmd(connHandle, tempBytes(0), atCmdStringLen, timeoutMilliSec)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_HFP_SendRequest_GetSubscriberNumber(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        Dim retUInt As UInt32
        retUInt = Btsdk_HFAP_GetSubscriberNumber(connHandle)

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_HFP_SendRequest_GetNetworkOperator(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        Dim retUInt As UInt32
        retUInt = Btsdk_HFAP_NetworkOperatorReq(connHandle)

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_HFP_GetManufacturer(ByVal connHandle As UInt32, ByRef retManufacturer As String) As Boolean

        retManufacturer = ""

        If connHandle = 0 Then Return False

        Dim retUInt32 As UInt32


        Dim retCount As UInt16 = 255
        Dim byteArray(0 To 255) As Byte

        ReDim byteArray(0 To retCount - 1)
        retUInt32 = Btsdk_HFAP_GetManufacturerID(connHandle, byteArray(0), retCount)

        If retCount < 1 Then
            retManufacturer = ""
            Return False
        End If

        ReDim Preserve byteArray(0 To retCount - 1)

        If retUInt32 = BTSDK_OK Then
            retManufacturer = System.Text.Encoding.UTF8.GetString(byteArray)
            retManufacturer = Replace(retManufacturer, Chr(0), "")
            If Left(retManufacturer, 3) = "+CG" Then
                retManufacturer = Mid(retManufacturer, 7)
            End If
            retManufacturer = Trim(retManufacturer)

            Return True
        Else
            retManufacturer = ""
            Return False
        End If

    End Function


    Public Function BlueSoleil_HFP_GetModel(ByVal connHandle As UInt32, ByRef retModel As String) As Boolean

        retModel = ""
        If connHandle = 0 Then Return False

        Dim retUInt32 As UInt32


        Dim retCount As UInt16 = 255
        Dim byteArray(0 To retCount - 1) As Byte


        retUInt32 = Btsdk_HFAP_GetModelID(connHandle, byteArray(0), retCount)

        If retCount < 1 Then
            retModel = ""
            Return False
        End If

        ReDim Preserve byteArray(0 To retCount - 1)

        If retUInt32 = BTSDK_OK Then
            retModel = System.Text.Encoding.UTF8.GetString(byteArray)
            retModel = Replace(retModel, Chr(0), "")
            If Left(retModel, 3) = "+CG" Then
                retModel = Mid(retModel, 7)
            End If
            retModel = Trim(retModel)
            Return True
        Else
            retModel = ""
            Return False
        End If

    End Function



    Public Function BlueSoleil_HFP_SetWaveInDevice(ByVal waveInDeviceName As String) As Boolean

        Dim byteArray(0 To 0) As Byte

        Dim retUInt32 As UInt32

        Dim tempBytes(0 To 0) As Byte

        If waveInDeviceName = "" Then
            retUInt32 = Btsdk_HFAP_SetWaveInDevice(tempBytes(0), 0) 'not sure.
        Else
            tempBytes = System.Text.Encoding.UTF8.GetBytes(waveInDeviceName & Chr(0))
            retUInt32 = Btsdk_HFAP_SetWaveInDevice(tempBytes(0), CUInt(tempBytes.Length))
        End If

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function




    Public Function BlueSoleil_HFP_SetWaveOutDevice(ByVal waveOutDeviceName As String) As Boolean

        Dim byteArray(0 To 0) As Byte
        Dim dvcNameLen As Integer = waveOutDeviceName.Length

        Dim retUInt32 As UInt32

        Dim tempBytes(0 To 0) As Byte

        If waveOutDeviceName = "" Then
            retUInt32 = Btsdk_HFAP_SetWaveOutDevice(tempBytes(0), 0) 'not sure.
        Else
            tempBytes = System.Text.Encoding.UTF8.GetBytes(waveOutDeviceName & Chr(0))
            retUInt32 = Btsdk_HFAP_SetWaveOutDevice(tempBytes(0), CUInt(tempBytes.Length))
        End If

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function



    Public Function BlueSoleil_HFP_GetState(ByVal connHandle As UInt32, ByRef retState As UInt16) As Boolean

        retState = 0
        If connHandle = 0 Then Return False

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_GetCurrHFState(retState)

        If retUInt32 = BTSDK_OK Then
            Return True
        End If

        Application.DoEvents()
        retUInt32 = Btsdk_HFAP_GetCurrHFState(retState)
        If retUInt32 = BTSDK_OK Then
            Return True
        End If


        Return False



    End Function




    Public Function BlueSoleil_HFP_SetVoiceRecognitionState(ByVal connHandle As UInt32, ByVal enableVR As Boolean) As Boolean

        If connHandle = 0 Then Return False

        Dim parmByte As Byte = 0
        If enableVR = True Then parmByte = 1

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_VoiceRecognitionReq(connHandle, parmByte)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_HFP_SendDTMF(ByVal connHandle As UInt32, ByVal dtmfChar As String) As Boolean

        If connHandle = 0 Then Return False

        dtmfChar = UCase(Left(dtmfChar, 1))
        Select Case dtmfChar
            Case "A", "B", "C" : dtmfChar = "2"
            Case "D", "E", "F" : dtmfChar = "3"
            Case "G", "H", "I" : dtmfChar = "4"
            Case "J", "K", "L" : dtmfChar = "5"
            Case "M", "N", "O" : dtmfChar = "6"
            Case "P", "Q", "R", "S" : dtmfChar = "7"
            Case "T", "U", "V" : dtmfChar = "8"
            Case "W", "X", "Y", "Z" : dtmfChar = "9"

        End Select

        Dim allowedChars As String = "0123456789*#+"

        If InStr(1, allowedChars, dtmfChar) = 0 Then
            Return False
        End If


        Dim tempByte As Byte
        tempByte = CByte(Asc(dtmfChar))

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_TxDTMF(connHandle, tempByte)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function



    Public Function BlueSoleil_HFP_Dial(ByVal connHandle As UInt32, ByVal phoneNumber As String) As Boolean

        If connHandle = 0 Then Return False

        phoneNumber = Replace(phoneNumber, " ", "")

        If phoneNumber = "" Then Return False

        Dim tempBytes(0 To 0) As Byte
        tempBytes = System.Text.Encoding.UTF8.GetBytes(phoneNumber & Chr(0))

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_Dial(connHandle, tempBytes(0), CUShort(tempBytes.Length - 1))

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If


    End Function



    Public Function BlueSoleil_HFP_SetMicVol(ByVal connHandle As UInt32, ByVal micVolPct As Double) As Boolean

        If connHandle = 0 Then Return False

        Dim tempVal As Integer = CInt(micVolPct * 15 / 100)
        Dim micVolByte As Byte = CByte(tempVal)
        If micVolByte > 15 Then micVolByte = 15
        If micVolByte < 0 Then micVolByte = 0

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_SetMicVol(connHandle, micVolByte)

        ' retUInt32 = Btsdk_AGAP_SetMicVol(connHandle, micVolByte)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function



    Public Function BlueSoleil_HFP_SetSpeakerVol(ByVal connHandle As UInt32, ByVal spkVolPct As Double) As Boolean

        If connHandle = 0 Then Return False

        Dim tempVal As Integer = CInt(spkVolPct * 15 / 100)
        Dim micVolByte As Byte = CByte(tempVal)
        If micVolByte > 15 Then micVolByte = 15
        If micVolByte < 0 Then micVolByte = 0

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_SetSpkVol(connHandle, micVolByte)

        ' retUInt32 = Btsdk_AGAP_SetSpkVol(connHandle, micVolByte)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function



    Public Function BlueSoleil_HFP_Redial(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_LastNumRedial(connHandle)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function BlueSoleil_HFP_HangUp(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_CancelCall(connHandle)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_HFP_AnswerCall(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_AnswerCall(connHandle)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Function BlueSoleil_HFP_TransferAudioConnection(ByVal connHandle As UInt32) As Boolean

        If connHandle = 0 Then Return False

        'this sounds like it transfers the audio between PC and phone.

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_HFAP_AudioConnTrans(connHandle)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If

    End Function



    Public Function BlueSoleil_HFP_GetHFPstateDesc(ByVal hfpState As UInt16) As String

        Dim stateStr As String = "Unknown"

        Select Case hfpState


            Case BTSDK_HFAP_ST_IDLE
                stateStr = "Service Not Connected"       'no service?

            Case BTSDK_HFAP_ST_STANDBY
                stateStr = "Ready"    'service established

            Case BTSDK_HFAP_ST_RINGING
                stateStr = "Ringing"

            Case BTSDK_HFAP_ST_OUTGOINGCALL
                stateStr = "Outgoing Call"

            Case BTSDK_HFAP_ST_ONGOINGCALL
                stateStr = "Ongoing Call"

            Case BTSDK_HFAP_ST_BVRA
                stateStr = "Voice Recognition Ongoing"

            Case BTSDK_HFAP_ST_VOVG
                stateStr = "VOVG"

            Case BTSDK_HFAP_ST_HELDINCOMINGCALL
                stateStr = "Incoming Call Is Held"

            Case Else
                stateStr = "Unknown"

        End Select

        Return stateStr

    End Function

    Public Function BlueSoleil_AudioGateway_GetAGAPstateDesc(ByVal hfpState As UInt16) As String

        Dim stateStr As String = "Unknown"

        Select Case hfpState

            Case BTSDK_AGAP_ST_IDLE
                stateStr = "Service Not Connected"       'no service?

            Case BTSDK_AGAP_ST_STANDBY
                stateStr = "Ready"    'service established

            Case BTSDK_AGAP_ST_RINGING
                stateStr = "Ringing"

            Case BTSDK_AGAP_ST_OUTGOINGCALL
                stateStr = "Outgoing Call"

            Case BTSDK_AGAP_ST_ONGOINGCALL
                stateStr = "Ongoing Call"

            Case BTSDK_HFAP_ST_BVRA
                stateStr = "Voice Recognition Ongoing"

            Case BTSDK_HFAP_ST_VOVG
                stateStr = "VOVG"

            Case BTSDK_AGAP_ST_HELDINCOMINGCALL
                stateStr = "Incoming Call Is Held"



        End Select

        Return stateStr

    End Function


    Public Function BlueSoleil_AudioGateway_GetState(ByVal connHandle As UInt32, ByRef retState As UInt16) As Boolean

        Dim retUInt32 As UInt32
        retUInt32 = Btsdk_AGAP_GetAGState(retState)

        If retUInt32 = BTSDK_OK Then
            Return True
        Else
            Return False
        End If


    End Function

End Module
