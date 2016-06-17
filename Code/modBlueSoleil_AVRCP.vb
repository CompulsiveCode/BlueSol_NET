'modBlueSoleil_AVRCP - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices
Module modBlueSoleil_AVRCP


    Private Const BTSDK_OK As UInt32 = 0
    Private Const BTSDK_TRUE As Byte = 1
    Private Const BTSDK_FALSE As Byte = 0



    '/* Notification event IDs, possible values of BtSdkRegisterNotifiReqStru::event_id. */
    Private Const BTSDK_AVRCP_EVENT_PLAYBACK_STATUS_CHANGED As Byte = &H01
    Private Const BTSDK_AVRCP_EVENT_TRACK_CHANGED As Byte = &H02
    Private Const BTSDK_AVRCP_EVENT_TRACK_REACHED_END As Byte = &H03    '/* If any action (e.g. GetElementAttributes) Is undertaken On the CT As reaction To the EVENT_TRACK_REACHED_END, the CT should register the EVENT_TRACK_REACHED_END again before initiating this action In order To Get informed about intermediate changes regarding the track status. */
    Private Const BTSDK_AVRCP_EVENT_TRACK_REACHED_START As Byte = &H04  '/* If any action (e.g. GetElementAttributes) Is undertaken On the CT As reaction To the EVENT_TRACK_REACHED_START, the CT should register the EVENT_TRACK_REACHED_START again before initiating this action In order To Get informed about intermediate changes regarding the track status. */
    Private Const BTSDK_AVRCP_EVENT_PLAYBACK_POS_CHANGED As Byte = &H05
    Private Const BTSDK_AVRCP_EVENT_BATT_STATUS_CHANGED As Byte = &H06
    Private Const BTSDK_AVRCP_EVENT_SYSTEM_STATUS_CHANGED As Byte = &H07
    Private Const BTSDK_AVRCP_EVENT_PLAYER_APPLICATION_SETTING_CHANGED As Byte = &H08
    Private Const BTSDK_AVRCP_EVENT_NOW_PLAYING_CONTENT_CHANGED As Byte = &H09 '/* If the NowPlaying folder Is browsed As reaction To the EVENT_NOW_PLAYING_CONTENT_CHANGED, the CT should register the EVENT_NOW_PLAYING_CONTENT_CHANGED again before browsing the NowPlaying folder In order To Get informed about intermediate changes In that folder. */
    Private Const BTSDK_AVRCP_EVENT_AVAILABLE_PLAYERS_CHANGED As Byte = &H0A    '/* If the Media Player List Is browsed As reaction To the EVENT_AVAILABLE_PLAYERS_CHANGED, the CT should register the EVENT_AVAILABLE_PLAYERS_CHANGED again before browsing the Media Player list In order To Get informed about intermediate changes Of the available players. */
    Private Const BTSDK_AVRCP_EVENT_ADDRESSED_PLAYER_CHANGED As Byte = &H0B
    Private Const BTSDK_AVRCP_EVENT_UIDS_CHANGED As Byte = &H0C        '/* If the Media Player Virtual Filesystem Is browsed As reaction To the EVENT_UIDS_CHANGED, the CT should register the EVENT_UIDS_CHANGED again before browsing the Media Player Virtual Filesystem In order To Get informed about intermediate changes within the fileystem. */
    Private Const BTSDK_AVRCP_EVENT_VOLUME_CHANGED As Byte = &H0D


    '/*AV/C Panel Commands operation_id*/
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_POWER As UInt32 = &H40
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_VOLUME_UP As UInt32 = &H41
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_VOLUME_DOWN As UInt32 = &H42
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_MUTE As UInt32 = &H43
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_PLAY As UInt32 = &H44
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_STOP As UInt32 = &H45
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_PAUSE As UInt32 = &H46
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_RECORD As UInt32 = &H47
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_REWIND As UInt32 = &H48
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_FAST_FORWARD As UInt32 = &H49
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_EJECT As UInt32 = &H4A
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_FORWARD As UInt32 = &H4B
    Private Const BTSDK_AVRCP_OPID_AVC_PANEL_BACKWARD As UInt32 = &H4C

    '/*button state(0: pressed 1: released)*/
    '/*used by Btsdk_AVRCP_Passthrough_Cmd_Func parameter state_flag*/
    Private Const BTSDK_AVRCP_BUTTON_STATE_PRESSED As UInt32 = 0
    Private Const BTSDK_AVRCP_BUTTON_STATE_RELEASED As UInt32 = 1





    '/* AVRCP specific event. */
    Private Const BTSDK_APP_EV_AVRCP_BASE As UInt32 = &HB00
    '/* AVRCP TG specific event. */
    Private Const BTSDK_APP_EV_AVTG_BASE As UInt32 = BTSDK_APP_EV_AVRCP_BASE
    Private Const BTSDK_APP_EV_AVTG_ATTACHPLAYER_IND As UInt32 = (BTSDK_APP_EV_AVTG_BASE + &H1)
    Private Const BTSDK_APP_EV_AVRCP_DETACHPLAYER_IND As UInt32 = (BTSDK_APP_EV_AVTG_BASE + &H3)
    Private Const BTSDK_APP_EV_AVRCP_IND_CONN As UInt32 = BTSDK_APP_EV_AVTG_ATTACHPLAYER_IND
    Private Const BTSDK_APP_EV_AVRCP_IND_DISCONN As UInt32 = BTSDK_APP_EV_AVRCP_DETACHPLAYER_IND
    Private Const BTSDK_APP_EV_AVRCP_IND_CONN_CFM As UInt32 = (BTSDK_APP_EV_AVTG_BASE + &H02)
    Private Const BTSDK_APP_EV_AVRCP_PASSTHROUGH_IND As UInt32 = (BTSDK_APP_EV_AVTG_BASE + &H06)
    Private Const BTSDK_APP_EV_AVRCP_VENDORDEP_IND As UInt32 = (BTSDK_APP_EV_AVTG_BASE + &H07)
    Private Const BTSDK_APP_EV_AVRCP_METADATA_IND As UInt32 = (BTSDK_APP_EV_AVTG_BASE + &H0D)
    Private Const BTSDK_APP_EV_AVRCP_GROUPNAV_IND As UInt32 = (BTSDK_APP_EV_AVTG_BASE + &H0F)

    '/*AVRCP CT specific event. */
    Private Const BTSDK_APP_EV_AVCT_BASE As UInt32 = BTSDK_APP_EV_AVRCP_BASE
    Private Const BTSDK_APP_EV_AVRCP_UNITINFO_RSP As UInt32 = (BTSDK_APP_EV_AVCT_BASE + &H08)
    Private Const BTSDK_APP_EV_AVRCP_SUBUNITINFO_RSP As UInt32 = (BTSDK_APP_EV_AVCT_BASE + &H09)
    Private Const BTSDK_APP_EV_AVRCP_PASSTHROUGH_RSP As UInt32 = (BTSDK_APP_EV_AVCT_BASE + &H0A)
    Private Const BTSDK_APP_EV_AVRCP_VENDORDEP_RSP As UInt32 = (BTSDK_APP_EV_AVCT_BASE + &H0B)
    Private Const BTSDK_APP_EV_AVRCP_METADATA_RSP As UInt32 = (BTSDK_APP_EV_AVCT_BASE + &H0C)
    Private Const BTSDK_APP_EV_AVRCP_GROUPNAV_RSP As UInt32 = (BTSDK_APP_EV_AVCT_BASE + &H0E)

    '/* AVRCP CT change notification events */
    Private Const BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE As UInt16 = (BTSDK_APP_EV_AVRCP_BASE + &H80)
    Private Const BTSDK_APP_EV_AVRCP_PLAYBACK_STATUS_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H01)
    Private Const BTSDK_APP_EV_AVRCP_TRACK_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H02)
    Private Const BTSDK_APP_EV_AVRCP_TRACK_REACHED_END_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H03)
    Private Const BTSDK_APP_EV_AVRCP_TRACK_REACHED_START_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H04)
    Private Const BTSDK_APP_EV_AVRCP_PLAYBACK_POS_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H05)
    Private Const BTSDK_APP_EV_AVRCP_BATT_STATUS_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H06)
    Private Const BTSDK_APP_EV_AVRCP_SYSTEM_STATUS_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H07)
    Private Const BTSDK_APP_EV_AVRCP_PLAYER_APPLICATION_SETTING_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H08)
    Private Const BTSDK_APP_EV_AVRCP_NOW_PLAYING_CONTENT_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H09)
    Private Const BTSDK_APP_EV_AVRCP_AVAILABLE_PLAYERS_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H0A)
    Private Const BTSDK_APP_EV_AVRCP_ADDRESSED_PLAYER_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H0B)
    Private Const BTSDK_APP_EV_AVRCP_UIDS_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H0C)
    Private Const BTSDK_APP_EV_AVRCP_VOLUME_CHANGED_NOTIF As UInt16 = (BTSDK_APP_EV_AVRCP_CT_NOTIF_BASE + &H0D)

    '/* AVRCP CT AV/C & Browsing specific event */
    Private Const BTSDK_APP_EV_AVRCP_CT_METARSP_BASE As UInt16 = &HD00
    Private Const BTSDK_APP_EV_AVRCP_GET_CAPABILITIES_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H10)
    Private Const BTSDK_APP_EV_AVRCP_LIST_PLAYER_SETTING_ATTR_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H11)
    Private Const BTSDK_APP_EV_AVRCP_LIST_PLAYER_SETTING_VALUES_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H12)
    Private Const BTSDK_APP_EV_AVRCP_GET_CURRENTPLAYER_SETTING_VALUE_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H13)
    Private Const BTSDK_APP_EV_AVRCP_SET_CURRENTPLAYER_SETTING_VALUE_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H14)
    Private Const BTSDK_APP_EV_AVRCP_GET_PLAYER_SETTING_ATTR_TEXT_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H15)
    Private Const BTSDK_APP_EV_AVRCP_GET_PLAYER_SETTING_VALUE_TEXT_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H16)
    Private Const BTSDK_APP_EV_AVRCP_INFORM_CHARACTERSET_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H17)
    Private Const BTSDK_APP_EV_AVRCP_INFORM_BATTERYSTATUS_OF_CT_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H18)
    Private Const BTSDK_APP_EV_AVRCP_GET_ELEMENT_ATTR_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H20)
    Private Const BTSDK_APP_EV_AVRCP_GET_PLAY_STATUS_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H30)
    Private Const BTSDK_APP_EV_AVRCP_SET_ABSOLUTE_VOLUME_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H50)
    Private Const BTSDK_APP_EV_AVRCP_SET_ADDRESSED_PLAYER_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H60)
    Private Const BTSDK_APP_EV_AVRCP_SET_BROWSED_PLAYER_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H70)
    Private Const BTSDK_APP_EV_AVRCP_GET_FOLDER_ITEMS_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H71)
    Private Const BTSDK_APP_EV_AVRCP_CHANGE_PATH_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H72)
    Private Const BTSDK_APP_EV_AVRCP_GET_ITEM_ATTRIBUTES_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H73)
    Private Const BTSDK_APP_EV_AVRCP_PLAY_ITEM_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H74)
    Private Const BTSDK_APP_EV_AVRCP_SEARCH_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H80)
    Private Const BTSDK_APP_EV_AVRCP_ADDTO_NOWPLAYING_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &H90)
    Private Const BTSDK_APP_EV_AVRCP_GENERAL_REJECT_RSP As UInt16 = (BTSDK_APP_EV_AVRCP_CT_METARSP_BASE + &HA0)


    '/* Current status of playing */
    Private Const BTSDK_AVRCP_PLAYSTATUS_STOPPED As Byte = &H00
    Private Const BTSDK_AVRCP_PLAYSTATUS_PLAYING As Byte = &H01
    Private Const BTSDK_AVRCP_PLAYSTATUS_PAUSED As Byte = &H02
    Private Const BTSDK_AVRCP_PLAYSTATUS_FWD_SEEK As Byte = &H03
    Private Const BTSDK_AVRCP_PLAYSTATUS_REV_SEEK As Byte = &H04 '/* 0x05-0xfe are reserved */
    Private Const BTSDK_AVRCP_PLAYSTATUS_ERROR As Byte = &HFF




    Private Const BTSDK_AVRCP_PACKET_HEAD As Byte = &H01
    Private Const BTSDK_AVRCP_SUBPACKET As Byte = &H02

    '/* The type of subpacket in BtSdkGetFolderItemRsp struct */
    Private Const BTSDK_AVRCP_PACKET_BROWSABLE_ITEM As Byte = &H01
    Private Const BTSDK_AVRCP_PACKET_MEDIA_ATTR As Byte = &H02

    Private Const BTSDK_AVRCP_CHARACTERSETID_UTF8 As UInt16 = &H006A


    '/* Player application settings PDUs */
    '/* List of defined Player Application Settings And Values. */
    Private Const BTSDK_AVRCP_PASA_ILLEGAL As Byte = &H00
    Private Const BTSDK_AVRCP_PASA_EQUALIZER_ONOFF_STATUS As Byte = &H01
    Private Const BTSDK_AVRCP_PASA_REPEAT_MODE_STATUS As Byte = &H02
    Private Const BTSDK_AVRCP_PASA_SHUFFLE_ONOFF_STATUS As Byte = &H03
    Private Const BTSDK_AVRCP_PASA_SCAN_ONOFF_STATUS As Byte = &H04

    '/* as byte = &h01 Equalizer ON/OFF status */
    Private Const BTSDK_AVRCP_EQUALIZER_OFF As Byte = &H01
    Private Const BTSDK_AVRCP_EQUALIZER_ON As Byte = &H02

    '/* as byte = &h02 Repeat Mode status */
    Private Const BTSDK_AVRCP_REPEAT_MODE_OFF As Byte = &H01
    Private Const BTSDK_AVRCP_REPEAT_MODE_SINGLE_TRACK_REPEAT As Byte = &H02
    Private Const BTSDK_AVRCP_REPEAT_MODE_ALL_TRACK_REPEAT As Byte = &H03
    Private Const BTSDK_AVRCP_REPEAT_MODE_GROUP_REPEAT As Byte = &H04

    '/* as byte = &h03 Shuffle ON/OFF status */
    Private Const BTSDK_AVRCP_SHUFFLE_OFF As Byte = &H01
    Private Const BTSDK_AVRCP_SHUFFLE_ALL_TRACKS_SHUFFLE As Byte = &H02
    Private Const BTSDK_AVRCP_SHUFFLE_GROUP_SHUFFLE As Byte = &H03

    '/* as byte = &h04 Scan ON/OFF status */
    Private Const BTSDK_AVRCP_SCAN_OFF As Byte = &H01
    Private Const BTSDK_AVRCP_SCAN_ALL_TRACKS_SCAN As Byte = &H02
    Private Const BTSDK_AVRCP_SCAN_GROUP_SCAN As Byte = &H03






    '/* Media Content Navigation */
    '/* There are four scopes in which media content navigation may take place. 
    'Scopes summarizes them And they are described In more detail In the following sections.

    Private Const BTSDK_AVRCP_SCOPE_MEDIAPLAYER_LIST As Byte = &H00 '/* Media Player Item, Contains all available media players */
    Private Const BTSDK_AVRCP_SCOPE_MEDIAPLAYER_VIRTUAL_FILESYSTEM As Byte = &H01 '/* Folder Item And Media Element Item, The virtual filesystem containing the media content Of the browsed player */
    Private Const BTSDK_AVRCP_SCOPE_MEDIAPLAYER_SEARCH As Byte = &H02 '/* Media Element Item, The results Of a search operation On the browsed player */
    Private Const BTSDK_AVRCP_SCOPE_MEDIAPLAYER_NOWPLAYING As Byte = &H03 '/* Media Element Item, The Now Playing list (Or queue) Of the addressed player */

    '/* Item Type - 1 Octet */
    Private Const BTSDK_AVRCP_ITEMTYPE_MEDIAPLAYER_ITEM As Byte = &H01
    Private Const BTSDK_AVRCP_ITEMTYPE_FOLDER_ITEM As Byte = &H02
    Private Const BTSDK_AVRCP_ITEMTYPE_MEDIAELEMENT_ITEM As Byte = &H03

    '/* List of Media Attributes. */
    Private Const BTSDK_AVRCP_MA_ILLEGAL As Byte = &H00 '/* should Not be used */
    Private Const BTSDK_AVRCP_MA_TITLEOF_MEDIA As Byte = &H01 '/* Any text encoded In specified character Set */
    Private Const BTSDK_AVRCP_MA_NAMEOF_ARTIST As Byte = &H02 '/* Any text encoded In specified character Set */
    Private Const BTSDK_AVRCP_MA_NAMEOF_ALBUM As Byte = &H03 '/* Any text encoded In specified character Set */
    Private Const BTSDK_AVRCP_MA_NUMBEROF_MEDIA As Byte = &H04 '/* Numeric ASCII text With zero suppresses, ex. Track number Of the CD */
    Private Const BTSDK_AVRCP_MA_TOTALNUMBEROF_MEDIA As Byte = &H05 '/* Numeric ASCII text With zero suppresses */
    Private Const BTSDK_AVRCP_MA_GENRE As Byte = &H06 '/* Any text encoded In specified character Set */
    Private Const BTSDK_AVRCP_MA_PLAYING_TIME As Byte = &H07 '/* Playing time In millisecond, 2min30sec->150000, 08-as byte = &hFFFFFFFF reserved For future use */

    '/* Major Player Type - 1 Octet */
    Private Const BTSDK_AVRCP_MAJORPLAYERTYPE_AUDIO As Byte = &H01
    Private Const BTSDK_AVRCP_MAJORPLAYERTYPE_VIDEO As Byte = &H02
    Private Const BTSDK_AVRCP_MAJORPLAYERTYPE_BROADCASTING_AUDIO As Byte = &H04
    Private Const BTSDK_AVRCP_MAJORPLAYERTYPE_BROADCASTING_VIDEO As Byte = &H08

    '/* Player Sub Type - 4 Octets */
    Private Const BTSDK_AVRCP_PLAYERSUBTYPE_NONE As UInt32 = &H00000000
    Private Const BTSDK_AVRCP_PLAYERSUBTYPE_AUDIOBOOK As UInt32 = &H00000001
    Private Const BTSDK_AVRCP_PLAYERSUBTYPE_PODCAST As UInt32 = &H00000002

    '/* Feature Bit Mask - 16 Octets */
    Private Const BTSDK_AVRCP_FBM_OCTET_ALL As Byte = &HFF

    '/* Octet 0 */
    Private Const BTSDK_AVRCP_FBM_SELECT As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_UP As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_DOWN As Byte = &H04
    Private Const BTSDK_AVRCP_FBM_LEFT As Byte = &H08
    Private Const BTSDK_AVRCP_FBM_RIGHT As Byte = &H10
    Private Const BTSDK_AVRCP_FBM_RIGHTUP As Byte = &H20
    Private Const BTSDK_AVRCP_FBM_RIGHTDOWN As Byte = &H40
    Private Const BTSDK_AVRCP_FBM_LEFTUP As Byte = &H80

    '/* Octet 1 */
    Private Const BTSDK_AVRCP_FBM_LEFTDOWN As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_ROOTMENU As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_SETUPMENU As Byte = &H04
    Private Const BTSDK_AVRCP_FBM_CONTENTSMENU As Byte = &H08
    Private Const BTSDK_AVRCP_FBM_FAVORITEMENU As Byte = &H10
    Private Const BTSDK_AVRCP_FBM_EXIT As Byte = &H20
    Private Const BTSDK_AVRCP_FBM_0 As Byte = &H40
    Private Const BTSDK_AVRCP_FBM_1 As Byte = &H80

    '/* Octet 2 */
    Private Const BTSDK_AVRCP_FBM_2 As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_3 As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_4 As Byte = &H04
    Private Const BTSDK_AVRCP_FBM_5 As Byte = &H08
    Private Const BTSDK_AVRCP_FBM_6 As Byte = &H10
    Private Const BTSDK_AVRCP_FBM_7 As Byte = &H20
    Private Const BTSDK_AVRCP_FBM_8 As Byte = &H40
    Private Const BTSDK_AVRCP_FBM_9 As Byte = &H80

    '/* Octet 3 */
    Private Const BTSDK_AVRCP_FBM_DOT As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_ENTER As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_CLEAR As Byte = &H04
    Private Const BTSDK_AVRCP_FBM_CHANNELUP As Byte = &H08
    Private Const BTSDK_AVRCP_FBM_CHANNELDOWN As Byte = &H10
    Private Const BTSDK_AVRCP_FBM_PREVIOUSCHANNEL As Byte = &H20
    Private Const BTSDK_AVRCP_FBM_SOUNDSELECT As Byte = &H40
    Private Const BTSDK_AVRCP_FBM_INPUTSELECT As Byte = &H80

    '/* Octet 4 */
    Private Const BTSDK_AVRCP_FBM_DISPLAY_INFORMATION As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_HELP As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_PAGEUP As Byte = &H04
    Private Const BTSDK_AVRCP_FBM_PAGEDOWN As Byte = &H08
    Private Const BTSDK_AVRCP_FBM_POWER As Byte = &H10
    Private Const BTSDK_AVRCP_FBM_VOLUMEUP As Byte = &H20
    Private Const BTSDK_AVRCP_FBM_VOLUMEDOWN As Byte = &H40
    Private Const BTSDK_AVRCP_FBM_MUTE As Byte = &H80

    '/* Octet 5 */
    Private Const BTSDK_AVRCP_FBM_PLAY As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_STOP As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_PAUSE As Byte = &H04
    Private Const BTSDK_AVRCP_FBM_RECORD As Byte = &H08
    Private Const BTSDK_AVRCP_FBM_REWIND As Byte = &H10
    Private Const BTSDK_AVRCP_FBM_FASTFORWARD As Byte = &H20
    Private Const BTSDK_AVRCP_FBM_EJECT As Byte = &H40
    Private Const BTSDK_AVRCP_FBM_FORWARD As Byte = &H80

    '/* Octet 6 */
    Private Const BTSDK_AVRCP_FBM_BACKWARD As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_ANGLE As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_SUBPICTURE As Byte = &H04
    Private Const BTSDK_AVRCP_FBM_F1 As Byte = &H08
    Private Const BTSDK_AVRCP_FBM_F2 As Byte = &H10
    Private Const BTSDK_AVRCP_FBM_F3 As Byte = &H20
    Private Const BTSDK_AVRCP_FBM_F4 As Byte = &H40
    Private Const BTSDK_AVRCP_FBM_F5 As Byte = &H80

    '/* Octet 7 */
    Private Const BTSDK_AVRCP_FBM_VENDOR_UNIQUE As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_BASIC_GROUP_NAVIGATION As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_ADVANCED_CONTROL_PLAYER As Byte = &H04
    Private Const BTSDK_AVRCP_FBM_BROWSING As Byte = &H08
    Private Const BTSDK_AVRCP_FBM_SEARCHING As Byte = &H10
    Private Const BTSDK_AVRCP_FBM_ADDTO_NOWPLAYING As Byte = &H20
    Private Const BTSDK_AVRCP_FBM_UIDS_UNIQUE_INPLAYERBROWSE_TREE As Byte = &H40
    Private Const BTSDK_AVRCP_FBM_ONLY_BROWSABLE_WHEN_ADDRESSED As Byte = &H80

    '/* Octet 8 */
    Private Const BTSDK_AVRCP_FBM_ONLY_SEARCHABLE_WHEN_ADDRESSED As Byte = &H01
    Private Const BTSDK_AVRCP_FBM_NOWPLAYING As Byte = &H02
    Private Const BTSDK_AVRCP_FBM_UIDPERSISTENCY As Byte = &H04

    '/* Folder Item*/
    '/* Folder Type - 1 Octet */
    Private Const BTSDK_AVRCP_FOLDERTYPE_MIXED As Byte = &H00
    Private Const BTSDK_AVRCP_FOLDERTYPE_TITLES As Byte = &H01
    Private Const BTSDK_AVRCP_FOLDERTYPE_ALBUMS As Byte = &H02
    Private Const BTSDK_AVRCP_FOLDERTYPE_ARTISTS As Byte = &H03
    Private Const BTSDK_AVRCP_FOLDERTYPE_GENRES As Byte = &H04
    Private Const BTSDK_AVRCP_FOLDERTYPE_PLAYLISTS As Byte = &H05
    Private Const BTSDK_AVRCP_FOLDERTYPE_YEARS As Byte = &H06

    '/* Is Playable - 1 Octet */
    Private Const BTSDK_AVRCP_ISPLAYABLE_CANNOT As Byte = &H00
    Private Const BTSDK_AVRCP_ISPLAYABLE_CAN As Byte = &H01

    '/* Media Type - 1 Octet */
    Private Const BTSDK_AVRCP_MEDIATYPE_AUDIO As Byte = &H00
    Private Const BTSDK_AVRCP_MEDIATYPE_VIDEO As Byte = &H01

    '/* Browsing Commands */
    Private Const BTSDK_AVRCP_DIRECTION_FOLDER_UP As Byte = &H00
    Private Const BTSDK_AVRCP_DIRECTION_FOLDER_DOWN As Byte = &H01

    '/* Volume Handling */
    Private Const BTSDK_AVRCP_ABSOLUTE_VOLUME_MIN As Byte = &H00
    Private Const BTSDK_AVRCP_ABSOLUTE_VOLUME_MAX As Byte = &H7F

    '/* Basic Group Navigation	*/
    Private Const BTSDK_AVRCP_BGN_NEXTGROUP As UInt16 = &H0000
    Private Const BTSDK_AVRCP_BGN_PREVIOUSGROUP As UInt16 = &H0001



    '/* List of Error Status Code */
    Private Const BTSDK_AVRCP_ERROR_INVALID_COMMAND As Byte = &H00 '/* All */
    Private Const BTSDK_AVRCP_ERROR_INVALID_PARAMETER As Byte = &H01 '/* All */
    Private Const BTSDK_AVRCP_ERROR_SPECIFIED_PARAMETER_NOTFOUND As Byte = &H02 '/* All */
    Private Const BTSDK_AVRCP_ERROR_INTERNAL_ERROR As Byte = &H03 '/* All */
    Private Const BTSDK_AVRCP_ERROR_SUCCESSFUL As Byte = &H04 '/* All except where the response CType Is AV/C REJECTED */
    Private Const BTSDK_AVRCP_ERROR_UID_CHANGED As Byte = &H05 '/* All */
    Private Const BTSDK_AVRCP_ERROR_RESERVED As Byte = &H06 '/* All, ??? */
    Private Const BTSDK_AVRCP_ERROR_INVALID_DIRECTION As Byte = &H07 '/* The Direction parameter Is invalid, Change Path */
    Private Const BTSDK_AVRCP_ERROR_NOTA_DIRECTORY As Byte = &H08 '/* The UID provided does Not refer To a folder item, Change Path */
    Private Const BTSDK_AVRCP_ERROR_UID_DOESNOT_EXIST As Byte = &H09 '/* The UID provided does Not refer To any currently valid item. Change Path, PlayItem, AddToNowPlaying, GetItemAttributes */
    Private Const BTSDK_AVRCP_ERROR_INVALID_SCOPE As Byte = &H0A '/* The scope parameter Is invalid. GetFolderItems, PlayItem, AddToNowPlayer, GetItemAttributes. */
    Private Const BTSDK_AVRCP_ERROR_RANGE_OUTOF_BOUNDS As Byte = &H0B '/* The start Of range provided Is Not valid, GetFolderItems */
    Private Const BTSDK_AVRCP_ERROR_UID_ISA_DIRECTORY As Byte = &H0C '/* The UID provided refers To a directory, which cannot be handled by this media player, PlayItem, AddToNowPlaying */
    Private Const BTSDK_AVRCP_ERROR_MEDIA_INUSE As Byte = &H0D '/* The media Is Not able To be used For this operation at this time, PlayItem, AddToNowPlaying */
    Private Const BTSDK_AVRCP_ERROR_NOWPLAYING_LISTFULL As Byte = &H0E '/* No more items can be added To the Now Playing List, AddToNowPlaying */
    Private Const BTSDK_AVRCP_ERROR_SEARCH_NOTSUPPORTED As Byte = &H0F '/* The Browsed Media Player does Not support search, Search */
    Private Const BTSDK_AVRCP_ERROR_SEARCH_INPROGRESS As Byte = &H10 '/* A search operation Is already In progress, Search */
    Private Const BTSDK_AVRCP_ERROR_INVALID_PLAYERID As Byte = &H11 '/* The specified Player Id does Not refer To a valid player, SetAddressedPlayer, SetBrowsedPlayer */
    Private Const BTSDK_AVRCP_ERROR_PLAYER_NOT_BROWSABLE As Byte = &H12 '/* The Player Id supplied refers To a Media Player which does Not support browsing. SetBrowsedPlayer */
    Private Const BTSDK_AVRCP_ERROR_PLAYER_NOT_ADDRESSED As Byte = &H13 '/* The Player Id supplied refers To a player which Is Not currently addressed, And the command Is Not able To be performed If the player Is Not Set As addressed.Search, SetBrowsedPlayer */
    Private Const BTSDK_AVRCP_ERROR_NO_VALID_SEARCH_RESULTS As Byte = &H14 '/* The Search result list does Not contain valid entries, e.g. after being invalidated due To change Of browsed player. GetFolderItems */
    Private Const BTSDK_AVRCP_ERROR_NO_AVAILABLE_PLAYERS As Byte = &H15 '/* All */
    Private Const BTSDK_AVRCP_ERROR_ADDRESSED_PLAYER_CHANGED As Byte = &H16 '/* Register Notification. as byte = &h17-as byte = &hff Reserved all */
    Private Const BTSDK_AVRCP_ERROR_TIMEOUT As Byte = &H88 '/* Monitor timer expired. Private Error code. */
    Private Const BTSDK_AVRCP_ERROR_NOT_IMPLEMENTED As Byte = &H89 '/* Not Implemented response Is recived. Private Error code. */




    Private Const BTSDK_AVRCP_BATTERYSTATUS_NORMAL As Byte = &H0
    Private Const BTSDK_AVRCP_BATTERYSTATUS_WARNING As Byte = &H1
    Private Const BTSDK_AVRCP_BATTERYSTATUS_CRITICAL As Byte = &H2
    Private Const BTSDK_AVRCP_BATTERYSTATUS_EXTERNAL As Byte = &H3
    Private Const BTSDK_AVRCP_BATTERYSTATUS_FULL_CHARGE As Byte = &H4


    '/* AV/C Response Code, 4 Bits */
    Private Const BTSDK_AVRCP_RSP_NOT_IMPLEMENTED As Byte = &H8
    Private Const BTSDK_AVRCP_RSP_ACCEPTED As Byte = &H9
    Private Const BTSDK_AVRCP_RSP_REJECTED As Byte = &HA
    Private Const BTSDK_AVRCP_RSP_STABLE As Byte = &HC              '/* Implemented */
    Private Const BTSDK_AVRCP_RSP_CHANGED As Byte = &HD             '/* For notification */
    Private Const BTSDK_AVRCP_RSP_INTERIM As Byte = &HF

    '/* Capabilities ID */
    '/* Used by Btsdk_AVRCP_GetCapabilities Command to specific capability reuqested */
    Private Const BTSDK_AVRCP_CAPABILITYID_COMPANY_ID As Byte = &H2
    Private Const BTSDK_AVRCP_CAPABILITYID_EVENTS_SUPPORTED As Byte = &H3









    Public Event BlueSoleil_Event_AVRCP_PlayStatus(ByVal trackLenSec As Double, ByVal trackPosSec As Double, ByVal isPlaying As Boolean)
    Public Event BlueSoleil_Event_AVRCP_TrackAlbum(ByVal trackAlbum As String)
    Public Event BlueSoleil_Event_AVRCP_TrackArtist(ByVal trackArtist As String)
    Public Event BlueSoleil_Event_AVRCP_TrackTitle(ByVal trackTitle As String)
    Public Event BlueSoleil_Event_AVRCP_AbsoluteVolume(ByVal newVolPct As Double)

    Public Event BlueSoleil_Event_AVRCP_GetSupportedEvents(ByVal supportsPlaybackStatusChanged As Boolean, ByVal supportsTrackChanged As Boolean, ByVal supportsTrackEnded As Boolean, ByVal supportsTrackStarted As Boolean, ByVal supportsTrackPosChanged As Boolean, ByVal supportsBattStatusChanged As Boolean, ByVal supportsSystemStatusChanged As Boolean, ByVal supportsPlayerSettingChanged As Boolean, ByVal supportsNowPlayingContentChanged As Boolean, ByVal supportsNumPlayersChanged As Boolean, ByVal supportsCurrPlayerChanged As Boolean, ByVal supportsUIDsChanged As Boolean, ByVal supportsVolChanged As Boolean)

    Public Event BlueSoleil_Event_AVRCP_BatteryStatusChanged(ByVal isCritical As Boolean, ByVal isLow As Boolean, ByVal isNormal As Boolean, ByVal isCharging As Boolean, ByVal isFullyCharged As Boolean)


    Public Event BlueSoleil_Event_AVRCP_TrackChanged()
    Public Event BlueSoleil_Event_AVRCP_PlayerSetting(ByVal repeatOne As Boolean, ByVal shuffleOn As Boolean)


    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function avrcpCallback_CTResponse(ByVal hdl As UInt32, ByVal e As UInt16, ByVal param As IntPtr) As Byte
    Public avrcpEvent_CTResponse As avrcpCallback_CTResponse = AddressOf BlueSoleil_AVRCP_Callback_CTResponse_Func

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Function avrcpCallback_TGCommand(ByVal dvcHandle As UInt32, ByVal tl As Byte, ByVal cmd_type As UInt16, ByVal param As IntPtr) As Byte
    Public avrcpEvent_TGCommand As avrcpCallback_TGCommand = AddressOf BlueSoleil_AVRCP_Callback_TGCommand_Func

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub avrcpCallback_ConnectionEvent(ByVal the_event As UInt16, ByVal param As IntPtr)
    Public avrcpEvent_ConnectionEvent As avrcpCallback_ConnectionEvent = AddressOf BlueSoleil_AVRCP_Callback_ConnectionEvent_Func

    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> Public Delegate Sub avrcpCallback_PassThruCmd(ByVal op_id As Byte, ByVal state_flag As Byte)
    Public avrcpEvent_PassThruCmd As avrcpCallback_PassThruCmd = AddressOf BlueSoleil_AVRCP_Callback_PassThruCmd_Func




    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_GetCapabilitiesReq(ByVal dvcHandle As UInt32, ByRef ptrGetCapabilitiesReqStru As Byte) As UInt32
    End Function


    'Btsdk_AVRCP_InformCharSetReq
    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_InformCharSetReq(ByRef ptrInformCharSetReqStruc As Byte) As UInt32
    End Function


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_AVRCP_RegIndCbk4ThirdParty(ByVal ptrFunc_EventIndicator As UInt32)
    End Sub

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_AVRCP_CTRegResponseCbk(ByVal ptrFunc_CTresponseInfo As UInt32)
    End Sub


    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_AVRCP_TGRegCommandCbk(ByVal ptrFunc_TGCommand As UInt32)
    End Sub

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Sub Btsdk_AVRCP_RegPassThrCmdCbk4ThirdParty(ByVal ptrFunc_PassThruCmd As UInt32)
    End Sub




    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_RegNotifReq(ByVal dev_hdl As UInt32, ByRef ptrNotifReqStruc As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_PassThroughReq(ByRef passthruStru As Byte) As UInt32
    End Function


    ' <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    ' Private Function Btsdk_AVRCP_PassThroughReqEx(ByVal dev_hdl As UInt32, ByRef passthruStru As Byte) As UInt32
    ' End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_GetPlayStatusReq(ByVal dvcHandle As UInt32) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_GetItemAttrReq(ByVal dvcHandle As UInt32, ByRef struGetItemAttribReq As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_GetElementAttrReq(ByVal dvcHandle As UInt32, ByRef struGetElementAttribReq As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_GetCurPlayerAppSetValReq(ByVal dvcHandle As UInt32, ByRef struGetCurPlayerAppSetValReq As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_SetCurPlayerAppSetValReq(ByVal dvcHandle As UInt32, ByRef struSetCurPlayerAppSetValReq As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_ChangePathReq(ByVal dvcHandle As UInt32, ByRef struChangePathReq As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_GetFolderItemsReq(ByVal dvcHandle As UInt32, ByRef struGetFolderItemReq As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_SetAbsoluteVolReq(ByVal dvcHandle As UInt32, ByRef struSetAbsoluteVolReq As Byte) As UInt32
    End Function

    <DllImport("BsSDK.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.Cdecl)>
    Private Function Btsdk_AVRCP_SetBrowsedPlayerReq(ByVal dvcHandle As UInt32, ByRef struSetBrowsedPlayerReq As Byte) As UInt32
    End Function



    Public Function BlueSoleil_AVRCP_EnableBrowsing(ByVal dvcHandle As UInt32) As Boolean

        Dim struBytes_SetBrowsedPlayerReq(0 To 0) As Byte

        BlueSoleil_AVRCP_InitStruBytes_SetBrowsedPlayerReq(struBytes_SetBrowsedPlayerReq)

        Dim retUInt As UInt32
        retUInt = Btsdk_AVRCP_SetBrowsedPlayerReq(dvcHandle, struBytes_SetBrowsedPlayerReq(0))

        Return (retUInt = BTSDK_OK)

    End Function

    Public Sub BlueSoleil_AVRCP_SendCmd_Play(ByVal dvcHandle As UInteger)

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)
        Dim retInt As UInteger

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_PLAY
        tempBytes(6) = 0
        tempBytes(7) = 0
        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_PLAY
        tempBytes(6) = 0
        tempBytes(7) = 0
        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub



    Public Sub BlueSoleil_AVRCP_SendCmd_Pause(ByVal dvcHandle As UInteger)

        Dim retInt As UInteger

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_PAUSE
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_PAUSE
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub


    Public Sub BlueSoleil_AVRCP_SendCmd_Next(ByVal dvcHandle As UInteger)

        If dvcHandle = 0 Then Exit Sub

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_FORWARD
        tempBytes(6) = 0
        tempBytes(7) = 0

        Dim retInt As UInteger

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_FORWARD
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub

    Public Function BlueSoleil_AVRCP_SendReq_GetPlayStatus(ByVal dvcHandle As UInt32) As Boolean

        If dvcHandle = 0 Then Return False

        Dim retInt As UInt32
        retInt = Btsdk_AVRCP_GetPlayStatusReq(dvcHandle)

        Return (retInt = BTSDK_OK)

    End Function



    Private Sub BlueSoleil_AVRCP_InitStruBytes_GetCapabilitiesReqStru_GetEvents(ByRef inpByteArray() As Byte)


        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 1     'size, req type

        ReDim inpByteArray(0 To struSize - 1)

        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, 0, 4)

        inpByteArray(4) = BTSDK_AVRCP_CAPABILITYID_EVENTS_SUPPORTED     'request type

    End Sub

    Private Sub BlueSoleil_AVRCP_InitStruBytes_GetCapabilitiesReqStru_GetCompanyIDs(ByRef inpByteArray() As Byte)


        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 1     'size, req type

        ReDim inpByteArray(0 To struSize - 1)

        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, 0, 4)

        inpByteArray(4) = BTSDK_AVRCP_CAPABILITYID_COMPANY_ID       'request type

    End Sub

    Public Function BlueSoleil_AVRCP_SendReq_GetCapabilities_SupportedEvents(ByVal dvcHandle As UInt32) As Boolean

        If dvcHandle = 0 Then Return False


        Dim bytesGetCapabilitiesReq(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_GetCapabilitiesReqStru_GetEvents(bytesGetCapabilitiesReq)

        Dim retInt As UInt32
        retInt = Btsdk_AVRCP_GetCapabilitiesReq(dvcHandle, bytesGetCapabilitiesReq(0))

        Return (retInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_AVRCP_SendReq_GetPlayerSettings(ByVal dvcHandle As UInt32) As Boolean

        If dvcHandle = 0 Then Return False

        'this request queries the Repeat and Shuffle values.

        Dim bytesGetCurPlayerAppSetValReq(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_GetCurPlayerAppSetValReqStru(bytesGetCurPlayerAppSetValReq, True, True)

        Dim retInt As UInt32
        retInt = Btsdk_AVRCP_GetCurPlayerAppSetValReq(dvcHandle, bytesGetCurPlayerAppSetValReq(0))

        Return (retInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_AVRCP_SendReq_SetPlayerSettings(ByVal dvcHandle As UInt32, ByVal setRepeatOne As Boolean, ByVal setShuffle As Boolean) As Boolean

        If dvcHandle = 0 Then Return False

        'this request sets the Repeat and Shuffle values.

        Dim bytesGetCurPlayerAppSetValReq(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_SetCurPlayerAppSetValReqStru(bytesGetCurPlayerAppSetValReq, setRepeatOne, setShuffle)

        Dim retInt As UInt32
        retInt = Btsdk_AVRCP_SetCurPlayerAppSetValReq(dvcHandle, bytesGetCurPlayerAppSetValReq(0))

        Return (retInt = BTSDK_OK)

    End Function

    Public Function BlueSoleil_AVRCP_SendReq_SetAbsoluteVolumePct(ByVal dvcHandle As UInt32, ByVal volPct As Double) As Boolean

        If dvcHandle = 0 Then Return False

        Dim bytesSetAbsVolReq(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_SetAbsoluteVolReqStru(bytesSetAbsVolReq, volPct)

        Dim retInt As UInt32
        retInt = Btsdk_AVRCP_SetAbsoluteVolReq(dvcHandle, bytesSetAbsVolReq(0))

        Return (retInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_AVRCP_SendReq_GetItemInfo(ByVal dvcHandle As UInt32) As Boolean

        If dvcHandle = 0 Then Return False

        Dim bytesGetItemAttribs(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_GetItemAttribs(bytesGetItemAttribs)

        Dim retInt As UInt32
        retInt = Btsdk_AVRCP_GetItemAttrReq(dvcHandle, bytesGetItemAttribs(0))

        Return (retInt = BTSDK_OK)

    End Function

    Public Function BlueSoleil_AVRCP_SendReq_GetFolderItems_NowPlayingList(ByVal dvcHandle As UInt32) As Boolean

        Dim bytesGetFolderItemsReqStru(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_GetFolderItems(bytesGetFolderItemsReqStru, BTSDK_AVRCP_SCOPE_MEDIAPLAYER_NOWPLAYING)    ' BTSDK_AVRCP_SCOPE_MEDIAPLAYER_NOWPLAYING)

        Dim retUInt As UInt32 = Btsdk_AVRCP_GetFolderItemsReq(dvcHandle, bytesGetFolderItemsReqStru(0))

        Return (retUInt = BTSDK_OK)

    End Function


    Public Function BlueSoleil_AVRCP_SendReq_GetFolderItems_FileSystem(ByVal dvcHandle As UInt32) As Boolean

        Dim bytesGetFolderItemsReqStru(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_GetFolderItems(bytesGetFolderItemsReqStru, BTSDK_AVRCP_SCOPE_MEDIAPLAYER_VIRTUAL_FILESYSTEM)    ' BTSDK_AVRCP_SCOPE_MEDIAPLAYER_NOWPLAYING)

        Dim retUInt As UInt32 = Btsdk_AVRCP_GetFolderItemsReq(dvcHandle, bytesGetFolderItemsReqStru(0))

        Return (retUInt = BTSDK_OK)

    End Function

    Public Function BlueSoleil_AVRCP_SendReq_GetElementInfo(ByVal dvcHandle As UInt32) As Boolean

        If dvcHandle = 0 Then Return False

        Dim bytesGetItemAttribs(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_GetElementAttribs(bytesGetItemAttribs)

        Dim retInt As UInt32
        retInt = Btsdk_AVRCP_GetElementAttrReq(dvcHandle, bytesGetItemAttribs(0))

        If retInt <> BTSDK_OK Then

        End If

        Return (retInt = BTSDK_OK)

    End Function

    Public Sub BlueSoleil_AVRCP_SendCmd_Prev(ByVal dvcHandle As UInteger)

        If dvcHandle = 0 Then Exit Sub

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_BACKWARD
        tempBytes(6) = 0
        tempBytes(7) = 0

        Dim retInt As UInteger

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_BACKWARD
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub



    Public Sub BlueSoleil_AVRCP_SendCmd_Mute(ByVal dvcHandle As UInteger)

        If dvcHandle = 0 Then Exit Sub

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_MUTE
        tempBytes(6) = 0
        tempBytes(7) = 0

        Dim retInt As UInteger

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_MUTE
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub


    Public Sub BlueSoleil_AVRCP_SendCmd_Stop(ByVal dvcHandle As UInteger)

        If dvcHandle = 0 Then Exit Sub

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_STOP
        tempBytes(6) = 0
        tempBytes(7) = 0

        Dim retInt As UInteger

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_STOP
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub


    Public Sub BlueSoleil_AVRCP_SendCmd_VolumeUp(ByVal dvcHandle As UInteger)

        If dvcHandle = 0 Then Exit Sub

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_VOLUME_UP
        tempBytes(6) = 0
        tempBytes(7) = 0

        Dim retInt As UInteger

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_VOLUME_UP
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub


    Public Sub BlueSoleil_AVRCP_SendCmd_VolumeDown(ByVal dvcHandle As UInteger)

        If dvcHandle = 0 Then Exit Sub

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_VOLUME_DOWN
        tempBytes(6) = 0
        tempBytes(7) = 0

        Dim retInt As UInteger

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_VOLUME_DOWN
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub

    Public Sub BlueSoleil_AVRCP_SendCmd_Power(ByVal dvcHandle As UInteger)

        If dvcHandle = 0 Then Exit Sub

        Dim tempBytes(0 To 7) As Byte
        tempBytes = BitConverter.GetBytes(dvcHandle)
        ReDim Preserve tempBytes(0 To 7)

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_PRESSED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_POWER
        tempBytes(6) = 0
        tempBytes(7) = 0

        Dim retInt As UInteger

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

        Application.DoEvents()

        tempBytes(4) = BTSDK_AVRCP_BUTTON_STATE_RELEASED
        tempBytes(5) = BTSDK_AVRCP_OPID_AVC_PANEL_POWER
        tempBytes(6) = 0
        tempBytes(7) = 0

        retInt = Btsdk_AVRCP_PassThroughReq(tempBytes(0))
        'retInt = Btsdk_AVRCP_PassThroughReqEx(dvcHandle, tempBytes(0))

    End Sub

    Private Sub BlueSoleil_AVRCP_InitStruBytes_GetItemAttribs(ByRef inpByteArray() As Byte)

        Dim numAttribs As Byte = 0


        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 1 + 8 + 2 + 1 + (4 * numAttribs)

        Dim currByteIdx As Integer = 0

        'structure size
        ReDim inpByteArray(0 To struSize - 1)
        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, currByteIdx, tempIntBytes.Length)
        currByteIdx = currByteIdx + tempIntBytes.Length

        'scope
        inpByteArray(currByteIdx) = BTSDK_AVRCP_SCOPE_MEDIAPLAYER_NOWPLAYING
        currByteIdx = currByteIdx + 1

        'uid
        Dim tempUID As String = "0x0"
        tempIntBytes = System.Text.Encoding.UTF8.GetBytes(tempUID & Chr(0))
        'Array.Copy(tempIntBytes, 0, inpByteArray, currByteIdx, tempIntBytes.Length)
        currByteIdx = currByteIdx + 8

        'uid_counter
        inpByteArray(currByteIdx) = 0
        inpByteArray(currByteIdx + 1) = 0
        currByteIdx = currByteIdx + 2

        'num attributes.  zero for all.
        inpByteArray(currByteIdx) = numAttribs
        currByteIdx = currByteIdx + 1


        'attrib list.  don't need it since we request all.
        '
        ''



    End Sub

    Private Sub BlueSoleil_AVRCP_InitStruBytes_GetFolderItems(ByRef inpByteArray() As Byte, ByVal avrcpScope As Byte)

        Dim numAttribs As Byte = 1

        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 1 + 4 + 4 + 1 + (4 * numAttribs)       'struSize, scope, firstIdx, lastIdx, attrlistNumItems, attributeidArray

        numAttribs = 0

        Dim currByteIdx As Integer = 0

        'structure size
        ReDim inpByteArray(0 To struSize - 1)
        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, currByteIdx, tempIntBytes.Length)
        currByteIdx = currByteIdx + tempIntBytes.Length

        'scope.  should be BTSDK_AVRCP_SCOPE_MEDIAPLAYER_NOWPLAYING or similar.
        inpByteArray(currByteIdx) = avrcpScope
        currByteIdx = currByteIdx + 1

        'leave start index at zero
        tempIntBytes = BitConverter.GetBytes(CUInt(0))
        Array.Copy(tempIntBytes, 0, inpByteArray, currByteIdx, tempIntBytes.Length)
        currByteIdx = currByteIdx + 4 'tempIntBytes.Length

        'set end index to max.
        tempIntBytes = BitConverter.GetBytes(CUInt(0))
        Array.Copy(tempIntBytes, 0, inpByteArray, currByteIdx, tempIntBytes.Length)
        currByteIdx = currByteIdx + 4 'tempIntBytes.Length

        'leave attrCount at zero.
        inpByteArray(currByteIdx) = 0
        currByteIdx = currByteIdx + 1


    End Sub

    Private Sub BlueSoleil_AVRCP_InitStruBytes_GetElementAttribs(ByRef inpByteArray() As Byte)

        Dim numAttribs As Byte = 0


        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 8 + 1 + (4 * numAttribs)

        Dim currByteIdx As Integer = 0

        'structure size
        ReDim inpByteArray(0 To struSize - 1)
        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, currByteIdx, tempIntBytes.Length)
        currByteIdx = currByteIdx + tempIntBytes.Length

        'uid
        Dim tempUID As String = "0x0"
        tempIntBytes = System.Text.Encoding.UTF8.GetBytes(tempUID & Chr(0))
        'Array.Copy(tempIntBytes, 0, inpByteArray, currByteIdx, tempIntBytes.Length)
        currByteIdx = currByteIdx + 8

        'num attributes.  zero for all.
        inpByteArray(currByteIdx) = numAttribs
        currByteIdx = currByteIdx + 1


        'attrib list.  don't need it since we request all.
        '
        ''



    End Sub


    Private Sub BlueSoleil_AVRCP_InitStruBytes_RegisterNotifiReqStru(ByRef inpByteArray() As Byte, ByVal evtID As Byte, Optional ByVal playbackUpdateInterval As UInt32 = 100)

        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 1 + 4     'size, eventID, interval (if eventID is poschange)

        ReDim inpByteArray(0 To struSize - 1)

        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, 0, 4)

        inpByteArray(4) = evtID

        If evtID = BTSDK_AVRCP_EVENT_PLAYBACK_POS_CHANGED Then
            tempIntBytes = BitConverter.GetBytes(playbackUpdateInterval)
            Array.Copy(tempIntBytes, 0, inpByteArray, 5, 4)

        End If



    End Sub


    Private Sub BlueSoleil_AVRCP_InitStruBytes_SetAbsoluteVolReqStru(ByRef inpByteArray() As Byte, ByVal newVolPct As Double)

        Dim struSize As Integer = 4 + 1     'structure size, vol (0-127)

        ReDim inpByteArray(0 To struSize - 1)

        Dim tempIntBytes(0 To 3) As Byte
        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, 0, 4)

        Dim newVolInt As Integer = CInt(127 * (newVolPct / 100))
        If newVolInt > 127 Then newVolInt = 127
        If newVolInt < 0 Then newVolInt = 0

        inpByteArray(4) = CByte(newVolInt)

    End Sub

    Private Sub BlueSoleil_AVRCP_InitStruBytes_SetBrowsedPlayerReq(ByRef inpByteArray() As Byte)

        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 2     '4 byte size, 2 byte player id

        ReDim inpByteArray(0 To struSize - 1)

        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, 0, tempIntBytes.Length)

        inpByteArray(4) = 0
        inpByteArray(5) = 0

    End Sub


    Private Sub BlueSoleil_AVRCP_InitStruBytes_InformCharSetStru(ByRef inpByteArray() As Byte)

        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 1 + 2     '4 byte size, 1 byte num of charsets, 2 byte charset id

        ReDim inpByteArray(0 To struSize - 1)

        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, 0, tempIntBytes.Length)

        inpByteArray(4) = 1

        Dim tempShortBytes(0 To 1) As Byte
        tempShortBytes = BitConverter.GetBytes(BTSDK_AVRCP_CHARACTERSETID_UTF8)
        Array.Copy(tempShortBytes, 0, inpByteArray, 5, tempShortBytes.Length)

    End Sub

    Private Sub BlueSoleil_AVRCP_InitStruBytes_GetCurPlayerAppSetValReqStru(ByRef inpByteArray() As Byte, ByVal getRepeat As Boolean, ByVal getShuffle As Boolean)

        'this is used for querying the value that is set in the player.

        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 1 + (2 * 1)     'size, num of values, 2x 1-byte values

        ReDim inpByteArray(0 To struSize - 1)

        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, 0, 4)

        inpByteArray(4) = 2 'number of values we are setting.

        inpByteArray(5) = BTSDK_AVRCP_PASA_REPEAT_MODE_STATUS
        inpByteArray(6) = BTSDK_AVRCP_PASA_SHUFFLE_ONOFF_STATUS


    End Sub



    Private Sub BlueSoleil_AVRCP_InitStruBytes_SetCurPlayerAppSetValReqStru(ByRef inpByteArray() As Byte, ByVal setRepeatOne As Boolean, ByVal setShuffle As Boolean)

        'this is used for querying the value that is set in the player.

        Dim tempIntBytes(0 To 3) As Byte

        Dim struSize As Integer = 4 + 1 + (2 * 2)     'size, num of values, 2x 1-byte values

        ReDim inpByteArray(0 To struSize - 1)

        tempIntBytes = BitConverter.GetBytes(struSize)
        Array.Copy(tempIntBytes, 0, inpByteArray, 0, 4)

        inpByteArray(4) = 2 'number of values we are setting.

        inpByteArray(5) = BTSDK_AVRCP_PASA_REPEAT_MODE_STATUS
        If setRepeatOne = True Then
            inpByteArray(6) = BTSDK_AVRCP_REPEAT_MODE_SINGLE_TRACK_REPEAT
        Else
            inpByteArray(6) = BTSDK_AVRCP_REPEAT_MODE_ALL_TRACK_REPEAT
        End If

        inpByteArray(7) = BTSDK_AVRCP_PASA_SHUFFLE_ONOFF_STATUS
        If setShuffle = True Then
            inpByteArray(8) = BTSDK_AVRCP_SHUFFLE_ALL_TRACKS_SHUFFLE
        Else
            inpByteArray(8) = BTSDK_AVRCP_SHUFFLE_OFF
        End If


    End Sub



    Private Function BlueSoleil_AVRCP_Callback_RegisterNotificationType(ByVal dvcHandle As UInt32, ByVal notType As Byte) As UInt32

        Dim avEvtBytes(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_RegisterNotifiReqStru(avEvtBytes, notType)

        Dim retInt As UInteger = Btsdk_AVRCP_RegNotifReq(dvcHandle, avEvtBytes(0))

        Return retInt

    End Function


    Private Sub BlueSoleil_AVRCP_Callback_ConnectionEvent_Func(ByVal the_event As UInt16, ByVal param As IntPtr)


        Select Case the_event
            ' Case 

        End Select

    End Sub

    Private Sub BlueSoleil_AVRCP_Callback_PassThruCmd_Func(ByVal op_id As Byte, ByVal state_flag As Byte)

        Debug.Print("BlueSoleil_AVRCP_Callback_PassThruCmd_Func")
        '

        If state_flag <> BTSDK_AVRCP_BUTTON_STATE_PRESSED Then
            Exit Sub
        End If

        Select Case op_id
            Case BTSDK_AVRCP_OPID_AVC_PANEL_PLAY
                Debug.Print("BlueSoleil_AVRCP_Callback_PassThruCmd BTSDK_AVRCP_OPID_AVC_PANEL_PLAY")
                'user pushed PLAY on phone?


        End Select


    End Sub

    Private Sub BlueSoleil_AVRCP_Callback_PassThruCmd_Remove()

        BlueSoleil_AVRCP_Callback_PassThruCmd_Add(IntPtr.Zero)


    End Sub


    Private Sub BlueSoleil_AVRCP_Callback_ConnectionEvent_Remove()


        BlueSoleil_AVRCP_Callback_ConnectionEvent_Add(IntPtr.Zero)


    End Sub


    Private Function BlueSoleil_AVRCP_Callback_PassThruCmd_Add(ByVal ptrFunc_AVRCPPassthruCallback As IntPtr) As Boolean

        Dim retBool As Boolean = False

        'maybe try Btsdk_AVRCP_RegNotifReq 

        If ptrFunc_AVRCPPassthruCallback = IntPtr.Zero Then

            Btsdk_AVRCP_CTRegResponseCbk(CInt(0))
        Else
            Btsdk_AVRCP_CTRegResponseCbk(CType(ptrFunc_AVRCPPassthruCallback, UInt32))
        End If


        Return True

    End Function


    Private Function BlueSoleil_AVRCP_Callback_ConnectionEvent_Add(ByVal ptrFunc_AVRCPPassthruCallback As IntPtr) As Boolean

        Dim retBool As Boolean = False

        'maybe try Btsdk_AVRCP_RegNotifReq 

        If ptrFunc_AVRCPPassthruCallback = IntPtr.Zero Then

            Btsdk_AVRCP_RegIndCbk4ThirdParty(CInt(0))
        Else
            Btsdk_AVRCP_RegIndCbk4ThirdParty(CType(ptrFunc_AVRCPPassthruCallback, UInt32))
        End If


        Return True

    End Function



    Private Function BlueSoleil_AVRCP_Callback_CTResponse_ParseGetCapabilitiesEventsRspStru_EventsSupported(ByVal ptrEvtData As IntPtr, ByRef retPlaybackStatusChanged As Boolean, ByRef retTrackChanged As Boolean, ByRef retTrackEnded As Boolean, ByRef retTrackStarted As Boolean, ByRef retTrackPosChanged As Boolean, ByRef retBattStatusChanged As Boolean, ByRef retSystemStatusChanged As Boolean, ByRef retPlayerSettingChanged As Boolean, ByRef retNowPlayingContentChanged As Boolean, ByRef retNumPlayersChanged As Boolean, ByRef retCurrPlayerChanged As Boolean, ByRef retUIDsChanged As Boolean, ByRef retVolChanged As Boolean) As Boolean

        retPlaybackStatusChanged = False
        retTrackChanged = False
        retTrackEnded = False
        retTrackStarted = False
        retTrackPosChanged = False
        retBattStatusChanged = False
        retSystemStatusChanged = False
        retPlayerSettingChanged = False
        retNowPlayingContentChanged = False
        retNumPlayersChanged = False
        retCurrPlayerChanged = False
        retUIDsChanged = False
        retVolChanged = False



        'get the first 4 bytes to figure out the real size of the structure.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim strucSize As UInt32
        strucSize = BitConverter.ToUInt32(evtData, 0)

        ReDim evtData(0 To CInt(strucSize - 1))
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)

        Dim currByteIdx As Integer = 4

        Dim capabilityID As Byte = evtData(currByteIdx)
        currByteIdx = currByteIdx + 1

        Dim numVals As Integer = evtData(currByteIdx)
        currByteIdx = currByteIdx + 1

        Dim currRepeatVal As Integer = -1
        Dim currShuffleVal As Integer = -1

        Dim valSize As Integer = 1
        If capabilityID <> BTSDK_AVRCP_CAPABILITYID_EVENTS_SUPPORTED Then Return False



        Do
            If currByteIdx + valSize > strucSize - 1 Then Exit Do

            Dim attrValue1Byte As Byte = evtData(currByteIdx)

            Select Case attrValue1Byte
                Case BTSDK_AVRCP_EVENT_PLAYBACK_STATUS_CHANGED
                    retPlaybackStatusChanged = True

                Case BTSDK_AVRCP_EVENT_TRACK_CHANGED
                    retTrackChanged = True

                Case BTSDK_AVRCP_EVENT_TRACK_REACHED_END
                    retTrackEnded = True

                Case BTSDK_AVRCP_EVENT_TRACK_REACHED_START
                    retTrackStarted = True

                Case BTSDK_AVRCP_EVENT_PLAYBACK_POS_CHANGED
                    retTrackPosChanged = True

                Case BTSDK_AVRCP_EVENT_BATT_STATUS_CHANGED
                    retBattStatusChanged = True

                Case BTSDK_AVRCP_EVENT_SYSTEM_STATUS_CHANGED
                    retSystemStatusChanged = True

                Case BTSDK_AVRCP_EVENT_PLAYER_APPLICATION_SETTING_CHANGED
                    retPlayerSettingChanged = True

                Case BTSDK_AVRCP_EVENT_NOW_PLAYING_CONTENT_CHANGED
                    retNowPlayingContentChanged = True

                Case BTSDK_AVRCP_EVENT_AVAILABLE_PLAYERS_CHANGED
                    retNumPlayersChanged = True

                Case BTSDK_AVRCP_EVENT_ADDRESSED_PLAYER_CHANGED
                    retCurrPlayerChanged = True

                Case BTSDK_AVRCP_EVENT_UIDS_CHANGED
                    retUIDsChanged = True

                Case BTSDK_AVRCP_EVENT_VOLUME_CHANGED
                    retVolChanged = True

                Case Else


            End Select

            currByteIdx = currByteIdx + valSize
        Loop



        Return True

    End Function







    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseGetPlayerSetting(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'get the first 4 bytes to figure out the real size of the structure.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim strucSize As UInt32
        strucSize = BitConverter.ToUInt32(evtData, 0)

        ReDim evtData(0 To CInt(strucSize - 1))
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)

        Dim currByteIdx As Integer = 4

        Dim numVals As Integer = evtData(currByteIdx)
        currByteIdx = currByteIdx + 1

        Dim currRepeatVal As Integer = -1
        Dim currShuffleVal As Integer = -1

        Do
            If currByteIdx + 2 > strucSize Then Exit Do

            Dim attrType As Byte = evtData(currByteIdx)
            Dim attrValue As Byte = evtData(currByteIdx + 1)

            Select Case attrType

                Case BTSDK_AVRCP_PASA_REPEAT_MODE_STATUS
                    currRepeatVal = attrValue

                Case BTSDK_AVRCP_PASA_SHUFFLE_ONOFF_STATUS
                    currShuffleVal = attrValue

            End Select

            currByteIdx = currByteIdx + 2
        Loop


        Dim boolRepeatOne As Boolean = (currRepeatVal = BTSDK_AVRCP_REPEAT_MODE_SINGLE_TRACK_REPEAT)
        Dim boolShuffle As Boolean = (currShuffleVal = BTSDK_AVRCP_SHUFFLE_ALL_TRACKS_SHUFFLE)

        'RaiseEvent BlueSoleil_Event_AVRCP_PlayerSetting(boolRepeatOne, boolShuffle)
        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_PlayerSetting(boolRepeatOne, boolShuffle))
        t.Start()

    End Sub




    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseBattStatusChangedNotif(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'get the first 4 bytes to figure out the real size of the structure.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim strucSize As UInt32
        strucSize = BitConverter.ToUInt32(evtData, 0)


        If strucSize < 6 Then Exit Sub

        ReDim evtData(0 To CInt(strucSize - 1))
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)

        Dim currByteIdx As Integer = 4


        Dim rspCode As Byte = evtData(currByteIdx)
        currByteIdx = currByteIdx + 1

        Dim statusID As Byte = evtData(currByteIdx)


        Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - BatteryChanged - rspCode = " & rspCode & "   statusID = " & statusID)


        If rspCode = BTSDK_AVRCP_RSP_CHANGED Then

            Select Case statusID

                Case BTSDK_AVRCP_BATTERYSTATUS_CRITICAL
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_BatteryStatusChanged(True, False, False, False, False))
                    t.Start()

                Case BTSDK_AVRCP_BATTERYSTATUS_WARNING
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_BatteryStatusChanged(False, True, False, False, False))
                    t.Start()

                Case BTSDK_AVRCP_BATTERYSTATUS_NORMAL
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_BatteryStatusChanged(False, False, True, False, False))
                    t.Start()

                Case BTSDK_AVRCP_BATTERYSTATUS_EXTERNAL
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_BatteryStatusChanged(False, False, False, True, False))
                    t.Start()

                Case BTSDK_AVRCP_BATTERYSTATUS_FULL_CHARGE
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_BatteryStatusChanged(False, False, False, False, True))
                    t.Start()

                Case Else
                    Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_BatteryStatusChanged(False, False, False, False, False))
                    t.Start()

            End Select

        End If



    End Sub

    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseGetStatusResponse(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'expecting 13 bytes.  4byte size, 4byte tracklen, 4byte trackpos, 1byte playstatus

        'first, read the length from structure bytes.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim struLen As Integer = BitConverter.ToInt32(evtData, 0)

        'if length is not what we're expecting, bail out.
        If struLen < 13 Then Exit Sub

        'expand our byte buffer to hold the entire structure.
        ReDim evtData(0 To struLen - 1)

        'and now copy the whole thing from the pointer to our array.
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)



        'parse structure.  
        Dim tempInt As UInt32

        Dim trackLen_Sec As Double
        tempInt = BitConverter.ToUInt32(evtData, 4)
        trackLen_Sec = tempInt / 1000

        Dim trackPos_Sec As Double
        tempInt = BitConverter.ToUInt32(evtData, 8)
        trackPos_Sec = tempInt / 1000

        Dim trackPlayStatus As Byte = evtData(12)

        'RaiseEvent BlueSoleil_Event_AVRCP_PlayStatus(trackLen_Sec, trackPos_Sec, (trackPlayStatus = BTSDK_AVRCP_PLAYSTATUS_PLAYING))
        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_PlayStatus(trackLen_Sec, trackPos_Sec, (trackPlayStatus = BTSDK_AVRCP_PLAYSTATUS_PLAYING)))
        t.Start()




    End Sub


    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseSetAbsoluteVolumeResponse(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'expecting 5 bytes.  4byte size, 1byte volume

        'first, read the length from structure bytes.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim struLen As Integer = BitConverter.ToInt32(evtData, 0)

        'if length is not what we're expecting, bail out.
        If struLen < 5 Then Exit Sub

        'expand our byte buffer to hold the entire structure.
        ReDim evtData(0 To struLen - 1)

        'and now copy the whole thing from the pointer to our array.
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)

        'parse structure.  
        Dim playerVol As Byte = evtData(4)

        Dim playerVolPct As Double = playerVol * 100 / 128
        If playerVolPct < 0 Then playerVolPct = 0
        If playerVolPct > 100 Then playerVolPct = 100

        'RaiseEvent BlueSoleil_Event_AVRCP_AbsoluteVolume(playerVolPct)
        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_AbsoluteVolume(playerVolPct))
        t.Start()

    End Sub

    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseVolumeChangedNotif(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'expecting 6 bytes.  4byte size, 1 byte rspCode, 1byte volume

        'first, read the length from structure bytes.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim struLen As Integer = BitConverter.ToInt32(evtData, 0)

        'if length is not what we're expecting, bail out.
        If struLen < 6 Then Exit Sub

        'expand our byte buffer to hold the entire structure.
        ReDim evtData(0 To struLen - 1)

        'and now copy the whole thing from the pointer to our array.
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)

        'parse structure.  
        Dim playerVol As Byte = evtData(5)

        Dim playerVolPct As Double = playerVol * 100 / 128
        If playerVolPct < 0 Then playerVolPct = 0
        If playerVolPct > 100 Then playerVolPct = 100

        'RaiseEvent BlueSoleil_Event_AVRCP_AbsoluteVolume(playerVolPct)
        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_AbsoluteVolume(playerVolPct))
        t.Start()

    End Sub



    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParsePlayStatusChanged(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'expecting 6 bytes.  4byte size, 1byte code, 1byte status code.

        'first, read the length from structure bytes.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim struLen As Integer = BitConverter.ToInt32(evtData, 0)

        'if length is not what we're expecting, bail out.
        If struLen < 6 Then
            Exit Sub
        End If

        'expand our byte buffer to hold the entire structure.
        ReDim evtData(0 To struLen - 1)

        'and now copy the whole thing from the pointer to our array.
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)

        'parse structure.  
        Dim trackPlayStatus As Byte = evtData(5)


        'RaiseEvent BlueSoleil_Event_AVRCP_PlayStatus(-1, -1, (trackPlayStatus = BTSDK_AVRCP_PLAYSTATUS_PLAYING))
        Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_PlayStatus(-1, -1, (trackPlayStatus = BTSDK_AVRCP_PLAYSTATUS_PLAYING)))
        t.Start()

    End Sub



    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseGetItemAttribsResponse(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'expecting 8+? bytes.  4 byte size, 4 byte strucType, strucData (either BtSdkGetItemAttrRspHeadStru or BtSdk4IDStringStru)


        'get the first 4 bytes to figure out the real size of the structure.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim strucSize As UInt32
        strucSize = BitConverter.ToUInt32(evtData, 0)

        Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseItemAttribs - Len = " & strucSize)

        If strucSize < 8 Then Exit Sub

        ReDim evtData(0 To CInt(strucSize - 1))
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)


        Dim currByteIdx As Integer = 4

        'parse structure...

        Dim headStatus As Byte = 0
        Dim headNumAttribs As Byte = 0

        Dim entryAttrID As UInt32
        Dim entryAttrCharSet As UInt16
        Dim entryAttrLen As UInt16
        Dim entryAttrVal As String = ""
        Dim tempVal(0 To 0) As Byte
        'get substruc-type
        Dim subType As UInt32 = 0



        Do

            If currByteIdx + 8 >= evtData.Length - 1 Then Exit Do

            subType = BitConverter.ToUInt32(evtData, currByteIdx)
            currByteIdx = currByteIdx + 4

            Select Case subType
                Case BTSDK_AVRCP_PACKET_HEAD
                    headStatus = evtData(currByteIdx)
                    currByteIdx = currByteIdx + 1
                    headNumAttribs = evtData(currByteIdx)
                    currByteIdx = currByteIdx + 1


                Case BTSDK_AVRCP_SUBPACKET
                    entryAttrID = BitConverter.ToUInt32(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 4
                    entryAttrCharSet = BitConverter.ToUInt16(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 2
                    entryAttrLen = BitConverter.ToUInt16(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 2

                    If entryAttrLen + currByteIdx <= evtData.Length Then
                        ReDim tempVal(0 To entryAttrLen - 1)
                        Array.Copy(evtData, currByteIdx, tempVal, 0, tempVal.Length)        'might be len-1 to remove null.

                        entryAttrVal = System.Text.Encoding.UTF8.GetString(tempVal)
                        entryAttrVal = Replace(entryAttrVal, Chr(0), "")

                        Select Case entryAttrID
                            Case BTSDK_AVRCP_MA_NAMEOF_ARTIST
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseItemAttribs - Artist = " & entryAttrVal)

                            Case BTSDK_AVRCP_MA_NAMEOF_ALBUM
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseItemAttribs - Album = " & entryAttrVal)

                            Case BTSDK_AVRCP_MA_TITLEOF_MEDIA
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseItemAttribs - Title = " & entryAttrVal)

                            Case BTSDK_AVRCP_MA_PLAYING_TIME
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseItemAttribs - PlayingTime = " & entryAttrVal)

                            Case Else
                                'dont care.

                        End Select

                    End If
                    currByteIdx = currByteIdx + entryAttrLen

                Case Else
                    'aww shit..  but this should never happen.

            End Select

        Loop


    End Sub





    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseGetElementAttribsResponse(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'expecting 8+? bytes.  4 byte size, 4 byte strucType, strucData (either BtSdkGetItemAttrRspHeadStru or BtSdk4IDStringStru)

        'get the first 4 bytes to figure out the real size of the structure.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim strucSize As UInt32
        strucSize = BitConverter.ToUInt32(evtData, 0)

        Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseElementAttribs - Len = " & strucSize)

        If strucSize < 8 Then Exit Sub

        ReDim evtData(0 To CInt(strucSize - 1))
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)


        Dim currByteIdx As Integer = 4

        'parse structure...

        Dim headStatus As Byte = 0
        Dim headNumAttribs As Byte = 0

        Dim entryAttrID As UInt32
        Dim entryAttrCharSet As UInt16
        Dim entryAttrLen As UInt16
        Dim entryAttrVal As String = ""
        Dim tempVal(0 To 0) As Byte
        'get substruc-type
        Dim subType As UInt32 = 0

        Do

            If currByteIdx + 8 >= evtData.Length - 1 Then Exit Do

            subType = BitConverter.ToUInt32(evtData, currByteIdx)
            currByteIdx = currByteIdx + 4

            Select Case subType
                Case BTSDK_AVRCP_PACKET_HEAD
                    headStatus = evtData(currByteIdx)
                    currByteIdx = currByteIdx + 1
                    headNumAttribs = evtData(currByteIdx)
                    currByteIdx = currByteIdx + 1

                Case BTSDK_AVRCP_SUBPACKET
                    entryAttrID = BitConverter.ToUInt32(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 4
                    entryAttrCharSet = BitConverter.ToUInt16(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 2
                    entryAttrLen = BitConverter.ToUInt16(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 2

                    If entryAttrLen + currByteIdx <= evtData.Length Then
                        ReDim tempVal(0 To entryAttrLen - 1)
                        Array.Copy(evtData, currByteIdx, tempVal, 0, tempVal.Length)        'might be len-1 to remove null.

                        entryAttrVal = System.Text.Encoding.UTF8.GetString(tempVal)
                        entryAttrVal = Replace(entryAttrVal, Chr(0), "")

                        Select Case entryAttrID
                            Case BTSDK_AVRCP_MA_NAMEOF_ARTIST
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseElementAttribs - Artist = " & entryAttrVal)
                                'RaiseEvent BlueSoleil_Event_AVRCP_TrackArtist(entryAttrVal)
                                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_TrackArtist(entryAttrVal))
                                t.Start()

                            Case BTSDK_AVRCP_MA_NAMEOF_ALBUM
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseElementAttribs - Album = " & entryAttrVal)
                                'RaiseEvent BlueSoleil_Event_AVRCP_TrackAlbum(entryAttrVal)
                                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_TrackAlbum(entryAttrVal))
                                t.Start()

                            Case BTSDK_AVRCP_MA_TITLEOF_MEDIA
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseElementAttribs - Title = " & entryAttrVal)
                                'RaiseEvent BlueSoleil_Event_AVRCP_TrackTitle(entryAttrVal)
                                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_TrackTitle(entryAttrVal))
                                t.Start()

                            Case BTSDK_AVRCP_MA_PLAYING_TIME
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseElementAttribs - PlayingTime = " & entryAttrVal)

                            Case Else
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseElementAttribs - Other = " & entryAttrID & " " & entryAttrVal)
                                'dont care.

                        End Select

                    End If
                    currByteIdx = currByteIdx + entryAttrLen

                Case Else
                    'aww shit..  but this should never happen.

            End Select

        Loop


    End Sub


    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseGetFolderItemsResponse(ByVal hdl As UInt32, ByVal ptrEvtData As IntPtr)

        'expecting 8+? bytes.  4 byte size, 4 byte strucType, strucData (either BtSdkBrowsableItemStru or BtSdk4IDStringStru)

        'get the first 4 bytes to figure out the real size of the structure.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim strucSize As UInt32
        strucSize = BitConverter.ToUInt32(evtData, 0)

        Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseGetFolderItems - Len = " & strucSize)


        If strucSize < 8 Then Exit Sub

        ReDim evtData(0 To CInt(strucSize - 1))
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)


        Dim currByteIdx As Integer = 4

        'parse structure...

        Dim headStatus As Byte = 0
        Dim headNumAttribs As Byte = 0

        Dim entryItemsNum As UInt16
        Dim entryItemUIDcounter As UInt16
        Dim entryItemLen As UInt16
        Dim entryItemType As Byte
        Dim entryStatus As Byte
        Dim entryAttrVal As String = ""
        Dim tempVal(0 To 0) As Byte
        'get substruc-type
        Dim subType As UInt32 = 0


        Dim folder_Type As Byte
        Dim folder_IsPlayable As Boolean = False
        Dim folder_UID As UInt64
        Dim folder_Name As String = ""

        Dim elementItem_UID As UInt64
        Dim elementItem_IsVideo As Boolean = False
        Dim elementItem_Name As String = ""


        Do

            If currByteIdx + 8 >= evtData.Length - 1 Then Exit Do

            subType = BitConverter.ToUInt32(evtData, currByteIdx)
            currByteIdx = currByteIdx + 4

            Select Case subType
                Case BTSDK_AVRCP_PACKET_HEAD, BTSDK_AVRCP_PACKET_BROWSABLE_ITEM
                    headStatus = evtData(currByteIdx)
                    currByteIdx = currByteIdx + 1
                    headNumAttribs = evtData(currByteIdx)
                    currByteIdx = currByteIdx + 1


                Case BTSDK_AVRCP_SUBPACKET, BTSDK_AVRCP_PACKET_MEDIA_ATTR

                    'parseBrowsableItemStru  let's cheat and just walk through it.  

                    entryItemsNum = BitConverter.ToUInt16(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 2
                    entryItemUIDcounter = BitConverter.ToUInt16(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 2
                    entryItemLen = BitConverter.ToUInt16(evtData, currByteIdx)
                    currByteIdx = currByteIdx + 2
                    entryItemType = evtData(currByteIdx)
                    currByteIdx = currByteIdx + 1
                    entryStatus = evtData(currByteIdx)
                    currByteIdx = currByteIdx + 1

                    If entryItemLen + currByteIdx <= evtData.Length Then

                        Select Case entryItemType
                            Case BTSDK_AVRCP_ITEMTYPE_MEDIAPLAYER_ITEM
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func parseBrowsableItemStru - Type = BTSDK_AVRCP_ITEMTYPE_MEDIAPLAYER_ITEM")
                            'parse BtSdkMediaPlayerItemStru at currByteIdx


                            Case BTSDK_AVRCP_ITEMTYPE_FOLDER_ITEM
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func parseBrowsableItemStru - Type = BTSDK_AVRCP_ITEMTYPE_FOLDER_ITEM")
                                'parse BtSdkFolderItemStru at currByteIdx
                                BlueSoleil_AVRCP_Callback_CTResponse_ParseFolderItemStru(evtData, currByteIdx, folder_UID, folder_Type, folder_IsPlayable, folder_Name)
                                Debug.Print("FolderUID = " & folder_UID & "  IsPlayable = " & folder_IsPlayable & "  FolderType = " & folder_Type & "  Name = " & folder_Name)


                            Case BTSDK_AVRCP_ITEMTYPE_MEDIAELEMENT_ITEM
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func parseBrowsableItemStru - Type = BTSDK_AVRCP_ITEMTYPE_MEDIAELEMENT_ITEM")
                                'parse BtSdkMediaElementItemStru at currByteIdx
                                BlueSoleil_AVRCP_Callback_CTResponse_ParseMediaElementItemStru(evtData, currByteIdx, elementItem_UID, elementItem_IsVideo, elementItem_Name)
                                Debug.Print("ElementUID = " & elementItem_UID & "  IsVideo = " & elementItem_IsVideo & "  Name = " & elementItem_Name)


                            Case Else
                                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func parseBrowsableItemStru - Type = Unknown = " & entryItemType)


                        End Select

                    End If
                    currByteIdx = currByteIdx + entryItemLen

                Case Else
                    'aww shit..  but this should never happen.

            End Select

        Loop


    End Sub

    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseFolderItemStru(ByRef struByteArray() As Byte, ByVal struStartIdx As Integer, ByRef retUID As UInt64, ByRef retFolderType As Byte, ByRef retFolderIsPlayable As Boolean, ByRef retFolderName As String)

        'expecting 14+? bytes.  8 byte UID, 1 byte folderType, 1 byte isPlayable, 2 byte charsetID, 2 byte nameLen, name


        'folder type could be:
        'BTSDK_AVRCP_FOLDERTYPE_MIXED
        'BTSDK_AVRCP_FOLDERTYPE_TITLES
        'BTSDK_AVRCP_FOLDERTYPE_ALBUMS
        'BTSDK_AVRCP_FOLDERTYPE_ARTISTIS
        'BTSDK_AVRCP_FOLDERTYPE_GENRES
        'BTSDK_AVRCP_FOLDERTYPE_PLAYLISTS
        'BTSDK_AVRCP_FOLDERTYPE_YEARS


        Dim strucSize As Integer = 14

        Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseFolderItemStru")


        Dim currByteIdx As Integer = struStartIdx

        Dim folderUID As UInt64 = BitConverter.ToUInt64(struByteArray, currByteIdx)
        currByteIdx = currByteIdx + 8

        Dim folderType As Byte = struByteArray(currByteIdx)
        currByteIdx = currByteIdx + 1

        Dim folderIsPlayable As Byte = struByteArray(currByteIdx)
        currByteIdx = currByteIdx + 1

        Dim charsetID As UInt16 = BitConverter.ToUInt16(struByteArray, currByteIdx)
        currByteIdx = currByteIdx + 2

        Dim nameLen As UInt16 = BitConverter.ToUInt16(struByteArray, currByteIdx)
        currByteIdx = currByteIdx + 2

        Dim folderName As String = System.Text.Encoding.UTF8.GetString(struByteArray, currByteIdx, nameLen)

        retUID = folderUID
        retFolderIsPlayable = (folderIsPlayable = BTSDK_AVRCP_ISPLAYABLE_CAN)
        retFolderName = folderName

    End Sub

    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseMediaElementItemStru(ByRef struByteArray() As Byte, ByVal struStartIdx As Integer, ByRef retUID As UInt64, ByRef retMediaIsVideo As Boolean, ByRef retElementItemName As String)

        'expecting 14+? bytes.  8 byte UID, 1 byte mediaType, 1 byte attribNum, 2 byte charsetID, 2 byte nameLen, name

        Dim strucSize As Integer = 14

        Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func ParseMediaElementItemStru")


        Dim currByteIdx As Integer = struStartIdx

        Dim elementUID As UInt64 = BitConverter.ToUInt64(struByteArray, currByteIdx)
        currByteIdx = currByteIdx + 8

        Dim mediaType As Byte = struByteArray(currByteIdx)
        currByteIdx = currByteIdx + 1

        Dim attribNum As Byte = struByteArray(currByteIdx)
        currByteIdx = currByteIdx + 1

        Dim charsetID As UInt16 = BitConverter.ToUInt16(struByteArray, currByteIdx)
        currByteIdx = currByteIdx + 2

        Dim nameLen As UInt16 = BitConverter.ToUInt16(struByteArray, currByteIdx)
        currByteIdx = currByteIdx + 2

        Dim itemName As String = System.Text.Encoding.UTF8.GetString(struByteArray, currByteIdx, nameLen)

        retUID = elementUID
        retMediaIsVideo = (mediaType = BTSDK_AVRCP_MEDIATYPE_VIDEO)
        retElementItemName = itemName

    End Sub


    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseBrowsableItemStru(ByRef struByteArray() As Byte, ByVal struStartIdx As Integer)



    End Sub


    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_ParseGeneralRejectRspStru(ByVal ptrEvtData As IntPtr, ByRef retCmdType As UInt32, ByRef retErrorCode As Byte, ByRef retErrorDesc As String)

        'expecting 9 bytes.  4 byte length, 4 byte cmdType, 1 byte errorCode



        'first, read the length from structure bytes.
        'get the first 4 bytes to figure out the real size of the structure.
        Dim evtData(0 To 3) As Byte
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)
        Dim strucSize As UInt32
        strucSize = BitConverter.ToUInt32(evtData, 0)

        Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_ParseGeneralRejectRspStru Len = " & strucSize)

        If strucSize < 8 Then Exit Sub

        If strucSize = 8 Then
            '    strucSize = 9           'stupid.
        End If

        ReDim evtData(0 To CInt(strucSize - 1))
        Marshal.Copy(ptrEvtData, evtData, 0, evtData.Length)

        Dim cmdType As UInt32
        Dim errorCode As Byte

        cmdType = BitConverter.ToUInt32(evtData, 4)

        retCmdType = cmdType
        retErrorCode = errorCode
        retErrorDesc = ""

        Debug.Print("Rejected CmdType = " & cmdType & "    ErrorCode = " & errorCode)


        If strucSize < 9 Then Exit Sub

        errorCode = evtData(8)

        Select Case errorCode
            Case BTSDK_AVRCP_ERROR_INVALID_COMMAND : retErrorDesc = "Invalid command, sent If TG received a PDU that it did Not understand."
            Case BTSDK_AVRCP_ERROR_INVALID_PARAMETER : retErrorDesc = "Invalid parameter, sent If the TG received a PDU With a parameter ID that it did Not understand. Sent If there Is only one parameter ID In the PDU."
            Case BTSDK_AVRCP_ERROR_SPECIFIED_PARAMETER_NOTFOUND : retErrorDesc = "Specified parameter Not found., sent If the parameter ID Is understood, but content Is wrong Or corrupted."
            Case BTSDK_AVRCP_ERROR_INTERNAL_ERROR : retErrorDesc = "Internal Error, sent If there are Error conditions Not covered by a more specific Error code."
            Case BTSDK_AVRCP_ERROR_UID_CHANGED : retErrorDesc = "UID Changed – The UIDs On the device have changed."
            Case BTSDK_AVRCP_ERROR_RESERVED : retErrorDesc = "Reserved."
            Case BTSDK_AVRCP_ERROR_INVALID_DIRECTION : retErrorDesc = "Invalid Direction – The Direction parameter Is invalid."
            Case BTSDK_AVRCP_ERROR_NOTA_DIRECTORY : retErrorDesc = "Not a Directory – The UID provided does Not refer To a folder item"
            Case BTSDK_AVRCP_ERROR_UID_DOESNOT_EXIST : retErrorDesc = "Does Not Exist – The UID provided does Not refer To any currently valid item"
            Case BTSDK_AVRCP_ERROR_INVALID_SCOPE : retErrorDesc = "Invalid Scope – The scope parameter Is invalid"
            Case BTSDK_AVRCP_ERROR_RANGE_OUTOF_BOUNDS : retErrorDesc = "Range Out Of Bounds – The start Of range provided Is Not valid"
            Case BTSDK_AVRCP_ERROR_UID_ISA_DIRECTORY : retErrorDesc = "UID Is a Directory – The UID provided refers To a directory, which cannot be handled by this media player"
            Case BTSDK_AVRCP_ERROR_MEDIA_INUSE : retErrorDesc = "Media In Use – The media Is Not able To be used For this operation at this time"
            Case BTSDK_AVRCP_ERROR_NOWPLAYING_LISTFULL : retErrorDesc = "Now Playing List Full – No more items can be added To the Now Playing List"
            Case BTSDK_AVRCP_ERROR_SEARCH_NOTSUPPORTED : retErrorDesc = "Search Not Supported – The Browsed Media Player does Not support search"
            Case BTSDK_AVRCP_ERROR_SEARCH_INPROGRESS : retErrorDesc = "Search In Progress – A search operation Is already In progress"
            Case BTSDK_AVRCP_ERROR_INVALID_PLAYERID : retErrorDesc = "Invalid Player Id – The specified Player Id does Not refer To a valid player"
            Case BTSDK_AVRCP_ERROR_PLAYER_NOT_BROWSABLE : retErrorDesc = "Player Not Browsable – The Player Id supplied refers To a Media Player which does Not support browsing."
            Case BTSDK_AVRCP_ERROR_PLAYER_NOT_ADDRESSED : retErrorDesc = "Player Not Addressed. The Player Id supplied refers To a player which Is Not currently addressed, And the command Is Not able To be performed If the player Is Not Set As addressed."
            Case BTSDK_AVRCP_ERROR_NO_VALID_SEARCH_RESULTS : retErrorDesc = "No valid Search Results – The Search result list does Not contain valid entries, e.g.after being invalidated due To change Of browsed player."
            Case BTSDK_AVRCP_ERROR_NO_AVAILABLE_PLAYERS : retErrorDesc = "No available players."
            Case BTSDK_AVRCP_ERROR_ADDRESSED_PLAYER_CHANGED : retErrorDesc = "Addressed Player Changed."

            Case Else
                retErrorDesc = "Unknown Error"

        End Select

        Debug.Print("Rejected CmdDesc = " & retErrorDesc)

    End Sub

    Private Function BlueSoleil_AVRCP_Callback_CTResponse_Func(ByVal hdl As UInteger, ByVal avrcpEvent As UInt16, ByVal evtData As IntPtr) As Byte

        Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func  Event = " & avrcpEvent)

        Dim strData As String = ""

        Select Case avrcpEvent


            Case BTSDK_APP_EV_AVRCP_GET_CAPABILITIES_RSP
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - Response To GetCapabilities")

                Dim retPlaybackStatusChanged, retTrackChanged, retTrackEnded, retTrackStarted, retTrackPosChanged, retBattStatusChanged, retSystemStatusChanged, retPlayerSettingChanged, retNowPlayingContentChanged, retNumPlayersChanged, retCurrPlayerChanged, retUIDsChanged, retVolChanged As Boolean
                BlueSoleil_AVRCP_Callback_CTResponse_ParseGetCapabilitiesEventsRspStru_EventsSupported(evtData, retPlaybackStatusChanged, retTrackChanged, retTrackEnded, retTrackStarted, retTrackPosChanged, retBattStatusChanged, retSystemStatusChanged, retPlayerSettingChanged, retNowPlayingContentChanged, retNumPlayersChanged, retCurrPlayerChanged, retUIDsChanged, retVolChanged)
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_GetSupportedEvents(retPlaybackStatusChanged, retTrackChanged, retTrackEnded, retTrackStarted, retTrackPosChanged, retBattStatusChanged, retSystemStatusChanged, retPlayerSettingChanged, retNowPlayingContentChanged, retNumPlayersChanged, retCurrPlayerChanged, retUIDsChanged, retVolChanged))
                t.Start()

            Case BTSDK_APP_EV_AVRCP_GET_PLAY_STATUS_RSP
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - Response To GetStatus")
                BlueSoleil_AVRCP_Callback_CTResponse_ParseGetStatusResponse(hdl, evtData)


            Case BTSDK_APP_EV_AVRCP_GET_ITEM_ATTRIBUTES_RSP
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - Response To GetItemAttribs")
                BlueSoleil_AVRCP_Callback_CTResponse_ParseGetItemAttribsResponse(hdl, evtData)


            Case BTSDK_APP_EV_AVRCP_GET_ELEMENT_ATTR_RSP
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - Response To GetElementAttribs")
                BlueSoleil_AVRCP_Callback_CTResponse_ParseGetElementAttribsResponse(hdl, evtData)

            Case BTSDK_APP_EV_AVRCP_SET_ABSOLUTE_VOLUME_RSP
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - Response To SetAbsoluteVolume")
                BlueSoleil_AVRCP_Callback_CTResponse_ParseSetAbsoluteVolumeResponse(hdl, evtData)

            Case BTSDK_APP_EV_AVRCP_PLAYBACK_STATUS_CHANGED_NOTIF
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - StatusChanged")
                BlueSoleil_AVRCP_Callback_CTResponse_ParsePlayStatusChanged(hdl, evtData)



            Case BTSDK_APP_EV_AVRCP_GET_CURRENTPLAYER_SETTING_VALUE_RSP
                '   Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - Response To GetPlayerSetting")
                BlueSoleil_AVRCP_Callback_CTResponse_ParseGetPlayerSetting(hdl, evtData)

            Case BTSDK_APP_EV_AVRCP_SET_CURRENTPLAYER_SETTING_VALUE_RSP
                '  Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - Response To SetPlayerSetting")
                '  BlueSoleil_AVRCP_Callback_CTResponse_ParseGetPlayerSetting(hdl, evtData)



            Case BTSDK_APP_EV_AVRCP_VOLUME_CHANGED_NOTIF
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - VolumeChanged")


            Case BTSDK_APP_EV_AVRCP_BATT_STATUS_CHANGED_NOTIF
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - BatteryStatusChanged")
                BlueSoleil_AVRCP_Callback_CTResponse_ParseBattStatusChangedNotif(hdl, evtData)



            Case BTSDK_APP_EV_AVRCP_PLAYBACK_POS_CHANGED_NOTIF
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - PosChanged")



            Case BTSDK_AVRCP_EVENT_TRACK_REACHED_START
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - TrackReachedStart")

            Case BTSDK_AVRCP_EVENT_TRACK_REACHED_END
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - TrackReachedEnd")


            Case BTSDK_APP_EV_AVRCP_TRACK_CHANGED_NOTIF
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - TrackChanged")
                'RaiseEvent BlueSoleil_Event_AVRCP_TrackChanged()
                Dim t As New Threading.Thread(Sub() RaiseEvent BlueSoleil_Event_AVRCP_TrackChanged())
                t.Start()


            Case BTSDK_APP_EV_AVRCP_NOW_PLAYING_CONTENT_CHANGED_NOTIF
                'now-playing list changed on device.
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - NowPlayingContentChanged")



            Case BTSDK_APP_EV_AVRCP_GET_FOLDER_ITEMS_RSP
                'response to request for folder items.
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - Response to GetFolderItems")


            Case BTSDK_APP_EV_AVRCP_GENERAL_REJECT_RSP
                Debug.Print("BlueSoleil_AVRCP_Callback_CTResponse_Func - BTSDK_APP_EV_AVRCP_GENERAL_REJECT_RSP")
                BlueSoleil_AVRCP_Callback_CTResponse_ParseGeneralRejectRspStru(evtData, 0, 0, "")

            Case Else


        End Select


        Return BTSDK_TRUE

    End Function



    Private Sub BlueSoleil_AVRCP_Callback_CTResponse_Remove()

        BlueSoleil_AVRCP_Callback_CTResponse_Add(IntPtr.Zero)


    End Sub


    Private Function BlueSoleil_AVRCP_Callback_CTResponse_Add(ByVal ptrFunc_AVRCPEventCallback As IntPtr) As Boolean

        Dim retBool As Boolean = False

        'maybe try Btsdk_AVRCP_RegNotifReq 

        If ptrFunc_AVRCPEventCallback = IntPtr.Zero Then

            Btsdk_AVRCP_CTRegResponseCbk(CInt(0))
        Else
            Btsdk_AVRCP_CTRegResponseCbk(CType(ptrFunc_AVRCPEventCallback, UInt32))
        End If

        Return retBool

    End Function


    Private Function BlueSoleil_AVRCP_Callback_TGCommand_Func(ByVal dvcHandle As UInt32, ByVal tl As Byte, ByVal cmd_type As UInt16, ByVal param As IntPtr) As Byte

        Debug.Print("BlueSoleil_AVRCP_Callback_TGCommand_Func")


        Select Case cmd_type
            '   Case BTSDK_APP_EV_AVRCP_GET_PLAY_STATUS_IND

            '   Debug.Print("BlueSoleil_AVRCP_Callback_TGCommand_Func BTSDK_APP_EV_AVRCP_GET_PLAY_STATUS_IND")



        End Select


        Return BTSDK_FALSE  '?

    End Function


    Private Function BlueSoleil_AVRCP_Callback_TGCommand_Add(ByVal ptrFunc_AVRCPEventCallback As IntPtr) As Boolean

        Dim retBool As Boolean = False

        'maybe try Btsdk_AVRCP_RegNotifReq 

        If ptrFunc_AVRCPEventCallback = IntPtr.Zero Then

            Btsdk_AVRCP_TGRegCommandCbk(CInt(0))
        Else
            Btsdk_AVRCP_TGRegCommandCbk(CType(ptrFunc_AVRCPEventCallback, UInt32))
        End If

        Return True

    End Function

    Private Sub BlueSoleil_AVRCP_Callback_TGCommand_Remove()

        BlueSoleil_AVRCP_Callback_TGCommand_Add(IntPtr.Zero)

    End Sub

    Private Sub BlueSoleil_AVRCP_RegisterNotificationTypes(ByVal dvcHandle As UInt32)

        '  BlueSoleil_AVRCP_Callback_RegisterNotificationType(dvcHandle, BTSDK_AVRCP_EVENT_BATT_STATUS_CHANGED)
        'BTSDK_AVRCP_EVENT_BATT_STATUS_CHANGED

        BlueSoleil_AVRCP_Callback_RegisterNotificationType(dvcHandle, BTSDK_AVRCP_EVENT_PLAYBACK_STATUS_CHANGED)
        BlueSoleil_AVRCP_Callback_RegisterNotificationType(dvcHandle, BTSDK_AVRCP_EVENT_TRACK_CHANGED)

        'BlueSoleil_AVRCP_Callback_RegisterNotificationType(dvcHandle, BTSDK_AVRCP_EVENT_TRACK_REACHED_START)


        ' BlueSoleil_AVRCP_Callback_RegisterNotificationType(dvcHandle, BTSDK_AVRCP_EVENT_PLAYBACK_POS_CHANGED)


    End Sub


    Public Sub BlueSoleil_AVRCP_RegisterCallbacks(ByVal dvcHandle As UInt32)


        Dim bytesInformCharSetStru(0 To 0) As Byte
        BlueSoleil_AVRCP_InitStruBytes_InformCharSetStru(bytesInformCharSetStru)
        Btsdk_AVRCP_InformCharSetReq(bytesInformCharSetStru(0))


        ' BlueSoleil_AVRCP_Callback_PassThruCmd_Add(IntPtr.Zero)
        ' BlueSoleil_AVRCP_Callback_ConnectionEvent_Add(IntPtr.Zero)
        ' BlueSoleil_AVRCP_Callback_TGCommand_Add(IntPtr.Zero)

        '  BlueSoleil_AVRCP_Callback_CTResponse_Add(IntPtr.Zero)


        ' Dim avrcpEventPtr_PassThruCmd As IntPtr
        ' avrcpEventPtr_PassThruCmd = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(avrcpEvent_PassThruCmd)
        ' BlueSoleil_AVRCP_Callback_PassThruCmd_Add(avrcpEventPtr_PassThruCmd)

        ' Dim avrcpEventPtr_ConnEvent As IntPtr
        ' avrcpEventPtr_ConnEvent = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(avrcpEvent_ConnectionEvent)
        ' BlueSoleil_AVRCP_Callback_ConnectionEvent_Add(avrcpEventPtr_ConnEvent)

        ' Dim avrcpEventPtr_TGCommand As IntPtr
        ' avrcpEventPtr_TGCommand = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(avrcpEvent_TGCommand)
        ' BlueSoleil_AVRCP_Callback_TGCommand_Add(avrcpEventPtr_TGCommand)

        BlueSoleil_AVRCP_RegisterNotificationTypes(dvcHandle)

        Dim functPtr_CTResponse As IntPtr = System.Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(avrcpEvent_CTResponse)
        BlueSoleil_AVRCP_Callback_CTResponse_Add(functPtr_CTResponse)



    End Sub

    Public Sub BlueSoleil_AVRCP_UnregisterCallbacks(ByVal dvcHandle As UInt32)

        BlueSoleil_AVRCP_Callback_CTResponse_Add(IntPtr.Zero)

    End Sub

End Module
