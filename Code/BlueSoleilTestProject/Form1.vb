
Option Explicit On
Option Strict On

Imports System.ComponentModel

Public Class Form1

    Public handlersDone As Boolean = False

    Public dvcNameArray(0 To 0) As String
    Public dvcHandleArray(0 To 0) As UInt32
    Public dvcArrayCount As Integer = 0
    Public dvcCurrHandle As UInt32 = 0
    Public dvcCurrAddress As String = ""

    Public panHandleDvc As UInt32 = 0
    Public panHandleConn As UInt32 = 0
    Public panHandleSvc As UInt32 = 0
    Public panIPaddress As String = "0.0.0.0"

    Public avrcpHandleDvc As UInt32 = 0
    Public avrcpHandleConn As UInt32 = 0
    Public avrcpHandleSvc As UInt32 = 0
    Public avrcpTrackTitle As String = ""
    Public avrcpTrackArtist As String = ""
    Public avrcpTrackAlbum As String = ""
    Public avrcpTrackPos As Double = 0
    Public avrcpTrackLen As Double = -2
    Public avrcpAbsVolPct As Double = 100
    Public avrcpIsPlaying As Boolean = False
    Public avrcpBatteryStatus As String = "Unknown"

    Public avrcpTrackPosUpdateTicks As Long = 0         'we are gonna use this to calculate the track position in between notifications.
    Public avrcpNeedsTrackInfo As Boolean = True        'these are used for knowing when to request track/status, instead of querying it every few seconds..
    Public avrcpNeedsPlayStatus As Boolean = True       ''
    Public avrcpNeedsSupportedEvents As Boolean = True


    Public avrcpEventSupported_PlaybackStatusChanged As Boolean = False
    Public avrcpEventSupported_TrackChanged As Boolean = False
    Public avrcpEventSupported_TrackEnded As Boolean = False
    Public avrcpEventSupported_TrackStarted As Boolean = False
    Public avrcpEventSupported_TrackPosChanged As Boolean = False
    Public avrcpEventSupported_BatteryStatusChanged As Boolean = False
    Public avrcpEventSupported_SystemStatusChanged As Boolean = False
    Public avrcpEventSupported_PlayerSettingChanged As Boolean = False
    Public avrcpEventSupported_NowPlayingContentChanged As Boolean = False
    Public avrcpEventSupported_NumPlayersChanged As Boolean = False
    Public avrcpEventSupported_CurrPlayerChanged As Boolean = False
    Public avrcpEventSupported_UIDsChanged As Boolean = False
    Public avrcpEventSupported_VolumeChanged As Boolean = False






    Public hfpHandle_LocalSvc_HFAG As UInt32 = 0
    Public hfpHandle_LocalSvc_HFunit As UInt32 = 0
    Public hfpHandle_LocalSvc_HSAG As UInt32 = 0
    Public hfpHandle_LocalSvc_HSunit As UInt32 = 0

    Public hfpHandleConnHSAG As UInt32 = 0

    Public hfpHandleDvc As UInt32 = 0
    Public hfpHandleConnHFAG As UInt32 = 0
    Public hfpHandleSvc As UInt32 = 0
    Public hfpStatusStr As String = "Not Connected"
    Public hfpSignalPct As Double = 0
    Public hfpBatteryPct As Double = 0
    Public hfpSpeakerVolumePct As Double = 0
    Public hfpMicVolumePct As Double = 0
    Public hfpNetworkAvailable As Boolean = False
    Public hfpNetworkName As String = ""
    Public hfpCallerIDno As String = ""
    Public hfpCallerIDname As String = ""
    Public hfpSubscriberNo As String = ""
    Public hfpSubscriberName As String = ""
    Public hfpModelName As String = ""
    Public hfpManufacturerName As String = ""
    Public hfpVoiceCmdStateEnabled As Boolean = False
    Public hfpIsRoaming As Boolean = False

    Public hfpRequestCounter As Integer = 0


    Public pbapHandleDvc As UInt32 = 0
    Public pbapHandleConn As UInt32 = 0
    Public pbapHandleSvc As UInt32 = 0

    Public mapHandleDvc As UInt32 = 0
    Public mapHandleConn As UInt32 = 0
    Public mapHandleSvc As UInt32 = 0
    Public mapMsgReceived As Boolean = False
    Public mapHandleMNSsvc As UInt32 = 0

    Public sppHandleDvc As UInt32 = 0
    Public sppHandleConn As UInt32 = 0
    Public sppHandleSvc As UInt32 = 0
    Public sppCOMMportNum As Integer = 0

    Public ftpHandleDvc As UInt32 = 0
    Public ftpHandleConn As UInt32 = 0
    Public ftpHandleSvc As UInt32 = 0
    Public ftpRemotePath As String = ""

    Public oppHandleDvc As UInt32 = 0
    Public oppHandleConn As UInt32 = 0
    Public oppHandleSvc As UInt32 = 0

    Private Delegate Sub DelegateFTPfoundFolder(ByVal foundFolderName As String)    'these are used for handling some FTP events on the UI thread.
    Private Delegate Sub DelegateFTPfoundFile(ByVal foundFileName As String, ByVal foundFileSize As UInt64)

    Public a2dpHandleDvc As UInt32 = 0
    Public a2dpHandleConn As UInt32 = 0
    Public a2dpHandleSvc As UInt32 = 0
    Public a2dpHandleSNKsvc As UInt32 = 0
    Public a2dpHandleSRCsvc As UInt32 = 0

    Private Sub btnInitBlueSoleil_Click(sender As Object, e As EventArgs) Handles btnInitBlueSoleil.Click



        Dim TorF As Boolean = False

        Do

            BlueSoleil_Init()

            Dim startTime As DateTime = Now
            Do
                My.Application.DoEvents()
                Threading.Thread.Sleep(100)

                If BlueSoleil_IsSDKinitialized() = True Then Exit Do

                If Now.Subtract(startTime).TotalSeconds > 5 Then

                    Exit Do

                End If

            Loop


            If BlueSoleil_IsSDKinitialized() = False Then
                'failed to init?
                My.Application.DoEvents()
            End If


            TorF = BlueSoleil_Status_RegisterCallbacks()

        Loop Until TorF = True



        BlueSoleil_StopBlueTooth()

        BlueSoleil_StartBlueTooth()



        BlueSoleil_SetLocalDeviceServiceClass(True, True, True)


        MsgBox("Done.")

    End Sub

    Private Sub btnDeInitBlueSoleil_Click(sender As Object, e As EventArgs) Handles btnDeInitBlueSoleil.Click

        BlueSoleil_Status_UnregisterCallbacks()

        BlueSoleil_Done()

        MsgBox("Done.")

    End Sub

    Private Sub btnGetPhonebook_Click(sender As Object, e As EventArgs) Handles btnGetPhonebook.Click

        Dim writeToPath As String = My.Application.Info.DirectoryPath
        If Strings.Right(writeToPath, 1) <> "\" Then writeToPath = writeToPath & "\"
        writeToPath = writeToPath & "Contacts"

        If IO.Directory.Exists(writeToPath) = False Then
            Try
                IO.Directory.CreateDirectory(writeToPath)
            Catch ex As Exception

            End Try
        End If

        Select Case LCase(cboPhoneBooks.Text)
            Case ("telecom/pb.vcf")

            Case ("telecom/ich.vcf")
                writeToPath = writeToPath & "\History_Incoming"
            Case ("telecom/och.vcf")
                writeToPath = writeToPath & "\History_Outgoing"
            Case ("telecom/mch.vcf")
                writeToPath = writeToPath & "\History_Missed"
            Case ("telecom/cch.vcf")
                writeToPath = writeToPath & "\History_Combined"

            Case Else


        End Select

        Try
            If IO.Directory.Exists(writeToPath) = False Then
                IO.Directory.CreateDirectory(writeToPath)
            End If
        Catch ex As Exception

        End Try


        Dim phoneDvcName As String = cboPhoneList.Text

        BlueSoleil_ConnectService_ByName("PBAP", phoneDvcName, pbapHandleDvc, pbapHandleConn, pbapHandleSvc)

        Dim pbFN As String = Mid(cboPhoneBooks.Text, InStrRev(cboPhoneBooks.Text, "/") + 1)

        Dim pullTorF As Boolean = False
        pullTorF = BlueSoleil_PBAP_PullPhoneBook_ByPath(pbapHandleConn, writeToPath & "\" & pbFN, cboPhoneBooks.Text)

        BlueSoleil_DisconnectServiceConn(pbapHandleConn)

        tvwPhoneBook.Nodes.Clear()

        Dim i As Integer = 0, j As Integer = 0
        Dim tempItem As TreeNode = Nothing
        Dim phoneItem As TreeNode = Nothing

        Dim contactName As String = "", contactNumbers(0 To 0) As String, contactNumberLabels(0 To 0) As String, contactNumberCount As Integer, contactImage As Bitmap = Nothing, contactOrganization As String = ""
        Dim contactEMail As String = "", contactLastCallDateTime As DateTime = Nothing, contactLastCallType As String = "", contactAddresses(0 To 0) As String, contactAddressLabels(0 To 0) As String, contactAddressCount As Integer, contactBirthday As String = "", contactGeoPos As String = "", contactNotes As String = ""

        Dim contactOffsets(0 To 0) As Long
        If pullTorF = True Then
            VCard_GetContactOffsets(writeToPath & "\" & pbFN, contactOffsets, "")


            For i = 0 To contactOffsets.Length - 1
                VCard_GetContactInfo(writeToPath & "\" & pbFN, contactOffsets(i), contactName, contactEMail, contactNumbers, contactNumberLabels, contactNumberCount, contactImage, contactLastCallDateTime, contactLastCallType, contactAddresses, contactAddressLabels, contactAddressCount, contactBirthday, contactGeoPos, contactNotes, contactOrganization)

                tempItem = tvwPhoneBook.Nodes.Add(contactName)

                If IsNothing(contactImage) = False Then
                    'contactImage.Dispose()
                    tempItem.Tag = contactImage
                End If

                tempItem.Tag = contactImage

                For j = 0 To contactNumberCount - 1
                    If contactNumberLabels(j) <> "" Then
                        phoneItem = tempItem.Nodes.Add(contactNumberLabels(j) & " Phone: " & contactNumbers(j))
                    Else
                        phoneItem = tempItem.Nodes.Add("Phone: " & contactNumbers(j))
                    End If

                Next j

                For j = 0 To contactAddressCount - 1
                    If contactAddressLabels(j) <> "" Then
                        phoneItem = tempItem.Nodes.Add(contactAddressLabels(j) & " Addr: " & Replace(contactAddresses(j), vbNewLine, " \ "))
                    Else
                        phoneItem = tempItem.Nodes.Add("Addr: " & Replace(contactAddresses(j), vbNewLine, " \ "))
                    End If

                Next j

                If contactEMail <> "" Then
                    phoneItem = tempItem.Nodes.Add("EMail: " & contactEMail)
                End If

                If contactBirthday <> "" Then
                    phoneItem = tempItem.Nodes.Add("Birthday: " & contactBirthday)
                End If

                If IsNothing(contactLastCallDateTime) = False AndAlso contactLastCallDateTime.Year <> 1 Then
                    phoneItem = tempItem.Nodes.Add("Last Call: " & contactLastCallType & " " & contactLastCallDateTime.ToShortDateString & " " & contactLastCallDateTime.ToShortTimeString)
                End If

                If contactGeoPos <> "" Then
                    phoneItem = tempItem.Nodes.Add("GeoPos: " & contactGeoPos)
                End If

                If contactNotes <> "" Then
                    phoneItem = tempItem.Nodes.Add("Notes: " & contactNotes)
                End If



            Next i



        End If

        tvwPhoneBook.ExpandAll()

        MsgBox("Done.  Connect = " & (pbapHandleConn <> 0) & "  Pull = " & pullTorF & ".  # of contacts = " & contactOffsets.Length)



        pbapHandleConn = 0

    End Sub

    Private Sub btnGetMessages_Click(sender As Object, e As EventArgs) Handles btnGetMessages.Click

        Dim writeToPath As String = My.Application.Info.DirectoryPath
        If Strings.Right(writeToPath, 1) <> "\" Then writeToPath = writeToPath & "\"
        writeToPath = writeToPath & "Inbox"

        If IO.Directory.Exists(writeToPath) = False Then
            Try
                IO.Directory.CreateDirectory(writeToPath)
            Catch ex As Exception

            End Try
        End If


        Dim flTorF As Boolean = False, arrayFolders(0 To 0) As String, folderCount As Integer = 0
        flTorF = BlueSoleil_MAP_PullFolderList(mapHandleConn, writeToPath & "\FolderList.xml")
        If flTorF = True Then
            folderCount = BlueSoleil_MAP_XML_GetFolderListInfo(writeToPath & "\FolderList.xml", arrayFolders)
        End If

        If chkGetMessagesUnread.Checked = False Then
            MsgBox("The program will be unresponsive for a minute or more.  Click OK to begin.")
        End If

        Dim mlTorF As Boolean = False
        mlTorF = BlueSoleil_MAP_PullMessageList(mapHandleConn, "inbox", writeToPath & "\MessageList.xml", chkGetMessagesUnread.Checked)

        Dim msgHandles(0 To 0) As String, msgSubjects(0 To 0) As String, msgDateTimes(0 To 0) As DateTime, msgSenderNames(0 To 0) As String, msgSenderAddresses(0 To 0) As String, msgRecipAddresses(0 To 0) As String, msgTypes(0 To 0) As String, msgSizes(0 To 0) As Integer, msgAttachmentSizes(0 To 0) As Integer, msgReadStates(0 To 0) As Boolean
        Dim msgCount As Integer = 0
        Dim firstMsgInfo As String = ""
        Dim msgTorF As Boolean = False
        Dim msgFileName As String = ""
        Dim msgText As String = "", msgFromName As String = "", msgFromNo As String = "", msgType As String = "", msgStatus As String = "", msgFolder As String = ""
        Dim msgAttachBytes(0 To 0) As Byte, msgAttachType As String = "", msgAttachSize As Long = 0

        Dim i As Integer
        Dim tempItem As ListViewItem = Nothing

        If mlTorF = True Then
            msgCount = BlueSoleil_MAP_XML_GetMessageListInfo(writeToPath & "\MessageList.xml", msgHandles, msgSubjects, msgDateTimes, msgSenderNames, msgSenderAddresses, msgRecipAddresses, msgTypes, msgSizes, msgAttachmentSizes, msgReadStates)

            If msgCount > 0 Then

                'fill the listview columns with the arrays from the MessageList.
                For i = 0 To msgCount - 1
                    tempItem = lvwMessages.Items.Add(msgSenderAddresses(i))
                    tempItem.SubItems.Add(msgSenderNames(i))
                    tempItem.SubItems.Add(msgRecipAddresses(i))
                    tempItem.SubItems.Add(msgSubjects(i))
                    tempItem.SubItems.Add(msgTypes(i))
                    tempItem.SubItems.Add(msgReadStates(i).ToString)
                    tempItem.SubItems.Add(msgAttachmentSizes(i).ToString)
                    tempItem.SubItems.Add(msgHandles(i))
                Next i


                'get some info about the first message, 
                msgFileName = writeToPath & "\" & msgHandles(0) & ".BMSG"
                Try
                    If IO.File.Exists(msgFileName) = True Then
                        IO.File.Delete(msgFileName)
                    End If
                Catch ex As Exception

                End Try

                msgTorF = BlueSoleil_MAP_PullMessage(mapHandleConn, msgHandles(0), msgFileName, False)
                If msgTorF = True Then
                    BMSG_GetMessageInfo(msgFileName, msgText, msgFromName, msgFromNo, msgType, msgStatus, msgFolder, msgAttachBytes, msgAttachType)
                    msgAttachSize = msgAttachBytes.Length
                    If msgAttachType = "" Then msgAttachSize = 0
                    firstMsgInfo = "First Message =" & vbNewLine & "From: " & msgFromNo & "  " & msgFromName & vbNewLine & "Msg Subj: " & msgSubjects(0) & vbNewLine & "Msg Text: " & msgText & vbNewLine & "Type: " & msgType & vbNewLine & "Status: " & msgStatus & vbNewLine & "Folder: " & msgFolder & vbNewLine & "Attachment Size: " & msgAttachSize & " bytes"
                End If


            End If



        End If

        MsgBox("Done." & vbNewLine & "GetFolderList = " & flTorF & " (" & folderCount & " folders)" & vbNewLine & "GetMessageList = " & mlTorF & ".  Pulled " & msgCount & " messages.  " & vbNewLine & "GetMessage = " & msgTorF & vbNewLine & firstMsgInfo)

    End Sub

    Private Sub btnSendText_Click(sender As Object, e As EventArgs) Handles btnSendText.Click

        Dim writeToPath As String = My.Application.Info.DirectoryPath
        If Strings.Right(writeToPath, 1) <> "\" Then writeToPath = writeToPath & "\"
        writeToPath = writeToPath & "Outbox"

        If IO.Directory.Exists(writeToPath) = False Then
            Try
                IO.Directory.CreateDirectory(writeToPath)
            Catch ex As Exception

            End Try
        End If

        If IO.File.Exists(writeToPath & "\ToSend.BMSG") = True Then
            Try
                IO.File.Delete(writeToPath & "\ToSend.BMSG")
            Catch ex As Exception

            End Try
        End If

        Dim phoneDvcName As String = cboPhoneList.Text

        BMSG_WriteMessageFile_Text(writeToPath & "\ToSend.BMSG", tbSendMessagePhoneNo.Text, tbSendMessageText.Text)

        Dim sendTorF As Boolean = False
        sendTorF = BlueSoleil_MAP_PushMessage_BMSG(mapHandleConn, writeToPath & "\ToSend.BMSG")


        MsgBox("Done.  Send = " & sendTorF)

    End Sub

    Private Sub btnTether_Click(sender As Object, e As EventArgs) Handles btnTether.Click

        Dim phoneDvcName As String = cboPhoneList.Text

        BlueSoleil_PAN_RegisterCallbackForIPaddress()

        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("PAN", phoneDvcName, panHandleDvc, panHandleConn, panHandleSvc)

        If connTorF = False Then
            BlueSoleil_PAN_UnregisterCallbackForIPaddress()
        End If

        MsgBox("Done.  Return = " & (panHandleConn <> 0))

    End Sub

    Private Sub btnUntether_Click(sender As Object, e As EventArgs) Handles btnUntether.Click

        BlueSoleil_DisconnectServiceConn(panHandleConn)

        BlueSoleil_PAN_UnregisterCallbackForIPaddress()

        MsgBox("Done.")

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnMediaConnect.Click

        Dim phoneDvcName As String = cboPhoneList.Text

        BlueSoleil_ConnectService_ByName("AVRCP", phoneDvcName, avrcpHandleDvc, avrcpHandleConn, avrcpHandleSvc)


        If avrcpHandleConn <> 0 Then
            BlueSoleil_AVRCP_RegisterCallbacks(dvcCurrHandle)
            avrcpNeedsTrackInfo = True
            avrcpNeedsPlayStatus = True

            MsgBox("Done.  Return = " & (avrcpHandleDvc <> 0) & ".  Anything playing on your phone (including Navigation) should play through the PC now.")
        Else


            MsgBox("Done.  Return = " & (avrcpHandleDvc <> 0))
        End If

    End Sub


    Private Sub bsHandler_Status_TurnOn()

    End Sub

    Private Sub bsHandler_Status_TurnOff()

    End Sub

    Private Sub bsHandler_Status_Plugged()

    End Sub

    Private Sub bsHandler_Status_Unplugged()

    End Sub

    Private Sub bsHandler_Status_DevicePaired()

    End Sub

    Private Sub bsHandler_Status_DeviceUnpaired()

    End Sub

    Private Sub bsHandler_Status_DeviceDeleted()

    End Sub

    Private Sub bsHandler_Status_DeviceFound(ByVal dvcHandle As UInt32)

        Debug.Print("DEVICE FOUND!!")

    End Sub

    Private Sub bsHandler_Status_ServiceConnectedInbound(ByVal dvcHandle As UInt32, ByVal svcHandle As UInt32, ByVal svcClass As UInt16)


    End Sub

    Private Sub bsHandler_Status_ServiceDisconnectedInbound(ByVal dvcHandle As UInt32, ByVal svcHandle As UInt32, ByVal svcClass As UInt16)


    End Sub

    Private Sub bsHandler_Status_ServiceConnectedOutbound(ByVal dvcHandle As UInt32, ByVal svcHandle As UInt32, ByVal svcClass As UInt16)


    End Sub

    Private Sub bsHandler_Status_ServiceDisconnectedOutbound(ByVal dvcHandle As UInt32, ByVal svcHandle As UInt32, ByVal svcClass As UInt16)


    End Sub



    Private Sub bsHandler_MAP_MsgNotification()
        Debug.Print("New Msg Received!")

        mapMsgReceived = True



    End Sub




    Private Sub bsHandler_FTP_FoundFolder(ByVal foundFolder As String)

        'The FTP_FoundFolder and FTP_FoundFile event-handlers are the only event-handlers that directly modify the UI (adding the file/folder to the listview).
        'since these events can happen on a non-UI thread, we need to check for that (InvokeRequired) and use BeginInvoke to execute this function on the UI thread.
        '
        'All of the other events simply update strings which are inherrently thread-safe, or update numeric values, and rely on the timer to refresh the UI.
        '

        If lvwFTPbrowser.InvokeRequired = True Then
            'thread-safe UI update.
            Dim args(0 To 0) As Object
            args(0) = foundFolder
            lvwFTPbrowser.BeginInvoke(New DelegateFTPfoundFolder(AddressOf bsHandler_FTP_FoundFolder), args)
        Else
            SyncLock lvwFTPbrowser
                Dim tempItem As ListViewItem = lvwFTPbrowser.Items.Add("[" & foundFolder & "]")
            End SyncLock
        End If



    End Sub

    Private Sub bsHandler_FTP_FoundFile(ByVal foundFile As String, ByVal fileSize As UInt64)

        'The FTP_FoundFolder and FTP_FoundFile event-handlers are the only event-handlers that directly modify the UI (adding the file/folder to the listview).
        'since these events can happen on a non-UI thread, we need to check for that (InvokeRequired) and use BeginInvoke to execute this function on the UI thread.
        '
        'All of the other events simply update strings which are inherrently thread-safe, or update numeric values, and rely on the timer to refresh the UI.
        '

        If lvwFTPbrowser.InvokeRequired = True Then
            'thread-safe UI update.
            Dim args(0 To 1) As Object
            args(0) = foundFile
            args(1) = fileSize
            lvwFTPbrowser.BeginInvoke(New DelegateFTPfoundFile(AddressOf bsHandler_FTP_FoundFile), args)
        Else
            SyncLock lvwFTPbrowser
                Dim tempItem As ListViewItem = lvwFTPbrowser.Items.Add(foundFile)
                tempItem.SubItems.Add(Format(fileSize, "###,###,###,##0"))

                lvwFTPbrowser.Sort()
            End SyncLock
        End If

    End Sub

    Private Sub bsHandler_PAN_IPchange(ByVal newIP As String)
        panIPaddress = newIP

    End Sub

    Private Sub bsHandler_AVRCP_TrackAlbum(ByVal newAlbum As String)

        avrcpTrackAlbum = newAlbum

    End Sub

    Private Sub bsHandler_AVRCP_TrackArtist(ByVal newArtist As String)

        avrcpTrackArtist = newArtist

    End Sub

    Private Sub bsHandler_AVRCP_TrackTitle(ByVal newTitle As String)

        avrcpTrackTitle = newTitle

    End Sub

    Private Sub bsHandler_AVRCP_TrackChanged()

        avrcpTrackPosUpdateTicks = 0
        avrcpNeedsTrackInfo = True
        avrcpNeedsPlayStatus = True

    End Sub

    Private Sub bsHandler_AVRCP_BatteryStatus(ByVal isCritical As Boolean, ByVal isLow As Boolean, ByVal isNormal As Boolean, ByVal isCharging As Boolean, ByVal isFullyCharged As Boolean)

        Dim retStr As String = "Unknown"
        If isCritical = True Then retStr = "Critical"
        If isLow = True Then retStr = "Low"
        If isNormal = True Then retStr = "Normal"
        If isCharging = True Then retStr = "Charging"
        If isFullyCharged = True Then retStr = "Fully Charged"

        avrcpBatteryStatus = retStr


    End Sub


    Private Sub bsHandler_AVRCP_PlayStatusInfo(ByVal trackLenSec As Double, ByVal trackPosSec As Double, ByVal isPlaying As Boolean)

        Dim prevTrackLen As Double = avrcpTrackLen
        Dim prevPlaying As Boolean = avrcpIsPlaying

        If trackLenSec <> -1 And trackPosSec <> -1 Then
            avrcpTrackLen = trackLenSec
            avrcpTrackPos = trackPosSec
            avrcpTrackPosUpdateTicks = (DateTime.UtcNow.Ticks \ TimeSpan.TicksPerMillisecond)
        End If

        avrcpIsPlaying = isPlaying

        If avrcpTrackLen <> prevTrackLen Or avrcpIsPlaying <> prevPlaying Then
            avrcpNeedsPlayStatus = True
            avrcpNeedsTrackInfo = True
        End If

    End Sub

    Private Sub bsHandler_AVRCP_AbsoluteVolume(ByVal volPct As Double)

        avrcpAbsVolPct = volPct


        Dim tempVolInt As Integer = CInt(TrackBar2.Maximum * avrcpAbsVolPct / 100)
        If tempVolInt < 0 Then tempVolInt = 0
        If tempVolInt > TrackBar2.Maximum Then tempVolInt = TrackBar2.Maximum
        TrackBar2.Value = tempVolInt



    End Sub

    Private Sub bsHandler_HFP_Ringing()

        hfpStatusStr = "Ringing"

    End Sub

    Private Sub bsHandler_HFP_OngoingCall()

        hfpStatusStr = "Ongoing Call"

    End Sub

    Private Sub bsHandler_HFP_OutgoingCall()

        hfpStatusStr = "Outgoing Call"

    End Sub

    Private Sub bsHandler_HFP_StandBy()

        hfpStatusStr = "StandBy"

    End Sub

    Private Sub bsHandler_HFP_SignalQuality(ByVal signalPct As Double)

        hfpSignalPct = signalPct

    End Sub

    Private Sub bsHandler_HFP_SpeakerVolume(ByVal volumePct As Double)

        hfpSpeakerVolumePct = volumePct

        Dim tempVolInt As Integer = CInt(TrackBar1.Maximum * hfpSpeakerVolumePct / 100)
        If tempVolInt < 0 Then tempVolInt = 0
        If tempVolInt > TrackBar1.Maximum Then tempVolInt = TrackBar1.Maximum
        TrackBar1.Value = tempVolInt

    End Sub


    Private Sub bsHandler_HFP_ExtCmdInd(ByVal atCmdResult As String)



    End Sub


    Private Sub bsHandler_HFP_MicVolume(ByVal volumePct As Double)

        hfpMicVolumePct = volumePct

    End Sub

    Private Sub bsHandler_HFP_BatteryCharge(ByVal batteryPct As Double)

        hfpBatteryPct = batteryPct

    End Sub

    Private Sub bsHandler_HFP_StartRoaming()

        hfpIsRoaming = True

    End Sub

    Private Sub bsHandler_HFP_EndRoaming()

        hfpIsRoaming = False

    End Sub

    Private Sub bsHandler_HFP_NetworkAvailable()

        hfpNetworkAvailable = True

    End Sub


    Private Sub bsHandler_HFP_NetworkUnavailable()

        hfpNetworkAvailable = False

    End Sub

    Private Sub bsHandler_HFP_NetworkName(ByVal networkName As String)

        Dim p1 As Integer = InStr(1, networkName, Chr(0))
        If p1 > 0 Then Exit Sub


        hfpNetworkName = networkName

    End Sub

    Private Sub bsHandler_HFP_ModelName(ByVal modelName As String)

        hfpModelName = modelName

    End Sub

    Private Sub bsHandler_HFP_ManufacturerName(ByVal manuName As String)

        hfpManufacturerName = manuName

    End Sub

    Private Sub bsHandler_HFP_ConnectionReleased()

        '!

    End Sub

    Private Sub bsHandler_HFP_CallerID(ByVal phoneNo As String, ByVal phoneName As String)

        hfpCallerIDno = phoneNo
        hfpCallerIDname = phoneName

    End Sub

    Private Sub bsHandler_HFP_SubscriberNo(ByVal phoneNo As String, ByVal phoneName As String)

        hfpSubscriberNo = phoneNo
        hfpSubscriberName = phoneName

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnMediaDisconnect.Click

        BlueSoleil_AVRCP_UnregisterCallbacks(avrcpHandleDvc)

        BlueSoleil_DisconnectServiceConn(avrcpHandleConn)

        avrcpHandleConn = 0
        avrcpHandleDvc = 0

        MsgBox("Done.")

    End Sub

    Private Sub btnMediaPlay_Click(sender As Object, e As EventArgs) Handles btnMediaPlay.Click

        BlueSoleil_AVRCP_SendCmd_Play(avrcpHandleDvc)
        avrcpNeedsTrackInfo = True
        avrcpNeedsPlayStatus = True

    End Sub

    Private Sub btnMediaPause_Click(sender As Object, e As EventArgs) Handles btnMediaPause.Click

        BlueSoleil_AVRCP_SendCmd_Pause(avrcpHandleDvc)
        avrcpIsPlaying = False
        avrcpNeedsPlayStatus = True

    End Sub

    Private Sub btnMediaStop_Click(sender As Object, e As EventArgs) Handles btnMediaStop.Click

        BlueSoleil_AVRCP_SendCmd_Stop(avrcpHandleDvc)
        avrcpNeedsTrackInfo = True
        avrcpNeedsPlayStatus = True

    End Sub

    Private Sub btnMediaPrev_Click(sender As Object, e As EventArgs) Handles btnMediaPrev.Click

        BlueSoleil_AVRCP_SendCmd_Prev(avrcpHandleDvc)
        avrcpNeedsTrackInfo = True
        avrcpNeedsPlayStatus = True

    End Sub

    Private Sub btnMediaNext_Click(sender As Object, e As EventArgs) Handles btnMediaNext.Click

        BlueSoleil_AVRCP_SendCmd_Next(avrcpHandleDvc)
        avrcpNeedsTrackInfo = True
        avrcpNeedsPlayStatus = True

    End Sub

    Private Sub btnMediaPower_Click(sender As Object, e As EventArgs)

        BlueSoleil_AVRCP_SendCmd_Power(avrcpHandleDvc)
        avrcpNeedsTrackInfo = True
        avrcpNeedsPlayStatus = True

    End Sub

    Private Sub btnMediaMute_Click(sender As Object, e As EventArgs)

        BlueSoleil_AVRCP_SendCmd_Mute(avrcpHandleDvc)
        avrcpNeedsTrackInfo = True
        avrcpNeedsPlayStatus = True

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'this timer is pretty much a request queue for all of the profiles.
        'it's linear and ensures no overlapping.

        If hfpVoiceCmdStateEnabled = True Then Exit Sub

        If mapMsgReceived = True Then
            mapMsgReceived = False
            MsgBox("New Message Received!")
            mapMsgReceived = False
        End If
        If mapHandleConn <> 0 Then
            If BlueSoleil_GetConnectionProperties(mapHandleConn) = False Then
                'lost connection?
                Debug.Print("MAP connection lost?")
                BlueSoleil_DisconnectServiceConn(mapHandleConn)
                mapHandleConn = 0
            End If
        End If


        If dvcCurrHandle <> 0 Then
            Dim currDvcLinkPct As Double = BlueSoleil_GetRemoteLinkQualityPct(dvcCurrHandle)
            lblDeviceLink.Text = "Link:  " & Format(currDvcLinkPct, "#0") & "%"

            Dim currDvcRSSIdb As Double = BlueSoleil_GetRemoteRSSI_Decibles(dvcCurrHandle)
            lblDeviceRSSI.Text = "RSSI:  " & currDvcRSSIdb & " db"

            'Dim svcClassArray(0 To 0) As UInt32, svcClassCount As Integer
            'Dim tempAdd As Boolean
            '  tempAdd = BlueSoleil_GetRemoteDeviceServiceHandles_Refresh(dvcCurrHandle, svcClassArray, svcClassCount)


            'using ZERO for the device handle here to get the statistics of the local device, since the remote device stats were unreliable.
            Dim currDvcBytesRcvd As UInt32, currDvcBytesSent As UInt32
            BlueSoleil_GetRemoteLinkDataStatistics(0, currDvcBytesRcvd, currDvcBytesSent)
            lblDeviceRcvd.Text = "Rcvd:  " & Format(currDvcBytesRcvd, "###,###,###,##0")
            lblDeviceSent.Text = "Sent:  " & Format(currDvcBytesSent, "###,###,###,##0")

        End If


        lblSerialCOMMport.Text = "Port:  COM " & sppCOMMportNum

        lblIPaddress.Text = "IP Address:  " & panIPaddress


        If avrcpHandleDvc <> 0 Then

            lblMediaAlbum.Text = "Album:  " & avrcpTrackAlbum
            lblMediaArtist.Text = "Artist:  " & avrcpTrackArtist
            lblMediaTitle.Text = "Title:  " & avrcpTrackTitle
            lblMediaBattery.Text = "Battery:  " & avrcpBatteryStatus

            If avrcpTrackPos >= 0 And avrcpTrackLen >= 0 Then

                'calculate the current position of the track, based on last known trackpos, and time since update.
                Dim currTicks As Long = (DateTime.UtcNow.Ticks \ TimeSpan.TicksPerMillisecond)
                Dim secondsSincePosUpdate As Double = (currTicks - avrcpTrackPosUpdateTicks) / 1000
                If avrcpTrackPosUpdateTicks = 0 Or avrcpIsPlaying = False Then secondsSincePosUpdate = 0

                Dim posSpan As New TimeSpan(0, 0, CInt(avrcpTrackPos + secondsSincePosUpdate))
                Dim lenSpan As New TimeSpan(0, 0, CInt(avrcpTrackLen))

                'maybe only update if playing.
                lblMediaPos.Text = "Pos: " & posSpan.Minutes & ":" & Format(posSpan.Seconds, "00") & " / " & lenSpan.Minutes & ":" & Format(lenSpan.Seconds, "00")
            End If

            If avrcpNeedsTrackInfo = True Then
                avrcpNeedsTrackInfo = False
                BlueSoleil_AVRCP_SendReq_GetElementInfo(avrcpHandleDvc)
            End If


            If avrcpNeedsPlayStatus = True Then
                avrcpNeedsPlayStatus = False
                BlueSoleil_AVRCP_SendReq_GetPlayStatus(avrcpHandleDvc)
                Application.DoEvents()
                Exit Sub
            End If


            If avrcpNeedsSupportedEvents = True Then
                avrcpNeedsSupportedEvents = False
                BlueSoleil_AVRCP_SendReq_GetCapabilities_SupportedEvents(avrcpHandleDvc)
                Application.DoEvents()
                Exit Sub
            End If


            'BlueSoleil_AVRCP_SendReq_GetPlayerSettings(avrcpHandleDvc)

        End If

        If hfpHandleConnHFAG <> 0 And hfpVoiceCmdStateEnabled = False Then

            Dim tempModelName As String = hfpModelName
            If tempModelName = "" Then tempModelName = cboPhoneList.Text

            lblHandsFreeStatus.Text = "Status:  " & hfpStatusStr
            lblHandsFreeBattery.Text = "Battery:  " & Format(hfpBatteryPct, "#0") & "%"
            lblHandsFreeSignal.Text = "Signal:  " & Format(hfpSignalPct, "#0") & "%"

            lblHandsFreeIncomingNo.Text = "Caller ID:  " & hfpCallerIDno & " " & hfpCallerIDname
            lblHandsFreeYourNo.Text = "Your No:  " & hfpSubscriberNo & " " & hfpSubscriberName
            lblHandsFreeNetwork.Text = "Network:  " & hfpNetworkName
            lblHandsFreePhoneType.Text = "Phone:  " & hfpManufacturerName & " " & tempModelName
            lblHandsFreeRoaming.Text = "Roaming:  " & hfpIsRoaming

            'only one request per timer tick.
            Select Case hfpRequestCounter Mod 10

                Case 0
                    BlueSoleil_HFP_SendATcmd(hfpHandleConnHFAG, "AT+CSQ", 1000)     'probably not required
                    Application.DoEvents()

                Case 1
                    BlueSoleil_HFP_SendATcmd(hfpHandleConnHFAG, "AT+CBC", 1000)     'probably not required
                    Application.DoEvents()


                Case 2
                    If hfpSubscriberNo = "" Then
                        BlueSoleil_HFP_SendRequest_GetSubscriberNumber(hfpHandleConnHFAG)
                        Application.DoEvents()
                    End If

                Case 3
                    If hfpNetworkName = "" Then
                        BlueSoleil_HFP_SendRequest_GetNetworkOperator(hfpHandleConnHFAG)
                        Application.DoEvents()
                    End If

                Case 4
                    If hfpManufacturerName = "" Then
                        ' BlueSoleil_HFP_SendATcmd(hfpHandleConnHFAG, "AT+CMER", 1000)
                        ' Application.DoEvents()

                        BlueSoleil_HFP_GetManufacturer(hfpHandleConnHFAG, hfpManufacturerName)
                        Application.DoEvents()
                    End If

                Case 5
                    If hfpModelName = "" Then
                        ' BlueSoleil_HFP_SendATcmd(hfpHandleConnHFAG, "AT+CGMM", 1000)

                        BlueSoleil_HFP_GetModel(hfpHandleConnHFAG, hfpManufacturerName)
                        Application.DoEvents()
                    End If

                Case 6
                    If hfpManufacturerName = "" Then
                        BlueSoleil_HFP_SendATcmd(hfpHandleConnHFAG, "AT+CGMI", 1000)
                    End If
                    hfpRequestCounter = -1

            End Select

            hfpRequestCounter = hfpRequestCounter + 1
        End If


        Dim isInit As Boolean = BlueSoleil_IsSDKinitialized()
        Dim isServerConn As Boolean = BlueSoleil_IsServerConnected()
        Dim isReady As Boolean = BlueSoleil_IsBluetoothReady()

        isReady = isReady

    End Sub

    Private Sub btnRefreshDevices_Click(sender As Object, e As EventArgs) Handles btnRefreshPairedDevices.Click


        BlueSoleil_GetPairedDevices_NamesAndHandles(dvcNameArray, dvcHandleArray, dvcArrayCount)

        cboPhoneList.Items.Clear()

        Dim i As Integer
        For i = 0 To dvcArrayCount - 1
            cboPhoneList.Items.Add(dvcNameArray(i))
        Next i
        If dvcArrayCount > 0 Then
            cboPhoneList.SelectedIndex = 0
        End If

        MsgBox("Done")

    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        RemoveHandler BlueSoleil_Status_TurnOn, AddressOf bsHandler_Status_TurnOn
        RemoveHandler BlueSoleil_Status_TurnOff, AddressOf bsHandler_Status_TurnOff
        RemoveHandler BlueSoleil_Status_Plugged, AddressOf bsHandler_Status_Plugged
        RemoveHandler BlueSoleil_Status_Unplugged, AddressOf bsHandler_Status_Unplugged
        RemoveHandler BlueSoleil_Status_DevicePaired, AddressOf bsHandler_Status_DevicePaired
        RemoveHandler BlueSoleil_Status_DeviceUnpaired, AddressOf bsHandler_Status_DeviceUnpaired
        RemoveHandler BlueSoleil_Status_DeviceDeleted, AddressOf bsHandler_Status_DeviceDeleted
        RemoveHandler BlueSoleil_Status_DeviceFound, AddressOf bsHandler_Status_DeviceFound

        RemoveHandler BlueSoleil_Status_ServiceConnectedInbound, AddressOf bsHandler_Status_ServiceConnectedInbound
        RemoveHandler BlueSoleil_Status_ServiceDisconnectedInbound, AddressOf bsHandler_Status_ServiceDisconnectedInbound
        RemoveHandler BlueSoleil_Status_ServiceConnectedOutbound, AddressOf bsHandler_Status_ServiceConnectedOutbound
        RemoveHandler BlueSoleil_Status_ServiceDisconnectedOutbound, AddressOf bsHandler_Status_ServiceDisconnectedOutbound

        RemoveHandler BlueSoleil_Event_FTP_FoundFolder, AddressOf bsHandler_FTP_FoundFolder
        RemoveHandler BlueSoleil_Event_FTP_FoundFile, AddressOf bsHandler_FTP_FoundFile

        RemoveHandler BlueSoleil_Event_MAP_MsgNotification, AddressOf bsHandler_MAP_MsgNotification

        RemoveHandler BlueSoleil_Event_PAN_IPchanged, AddressOf bsHandler_PAN_IPchange

        RemoveHandler BlueSoleil_Event_AVRCP_PlayStatus, AddressOf bsHandler_AVRCP_PlayStatusInfo
        RemoveHandler BlueSoleil_Event_AVRCP_TrackAlbum, AddressOf bsHandler_AVRCP_TrackAlbum
        RemoveHandler BlueSoleil_Event_AVRCP_TrackArtist, AddressOf bsHandler_AVRCP_TrackArtist
        RemoveHandler BlueSoleil_Event_AVRCP_TrackTitle, AddressOf bsHandler_AVRCP_TrackTitle
        RemoveHandler BlueSoleil_Event_AVRCP_TrackChanged, AddressOf bsHandler_AVRCP_TrackChanged
        RemoveHandler BlueSoleil_Event_AVRCP_AbsoluteVolume, AddressOf bsHandler_AVRCP_AbsoluteVolume
        RemoveHandler BlueSoleil_Event_AVRCP_BatteryStatusChanged, AddressOf bsHandler_AVRCP_BatteryStatus

        RemoveHandler BlueSoleil_Event_HFP_Ringing, AddressOf bsHandler_HFP_Ringing
        RemoveHandler BlueSoleil_Event_HFP_OngoingCall, AddressOf bsHandler_HFP_OngoingCall
        RemoveHandler BlueSoleil_Event_HFP_OutgoingCall, AddressOf bsHandler_HFP_OutgoingCall
        RemoveHandler BlueSoleil_Event_HFP_Standby, AddressOf bsHandler_HFP_StandBy
        RemoveHandler BlueSoleil_Event_HFP_SignalQuality, AddressOf bsHandler_HFP_SignalQuality
        RemoveHandler BlueSoleil_Event_HFP_BatteryCharge, AddressOf bsHandler_HFP_BatteryCharge
        RemoveHandler BlueSoleil_Event_HFP_SpeakerVolume, AddressOf bsHandler_HFP_SpeakerVolume
        RemoveHandler BlueSoleil_Event_HFP_MicVolume, AddressOf bsHandler_HFP_MicVolume
        RemoveHandler BlueSoleil_Event_HFP_ExtCmdInd, AddressOf bsHandler_HFP_ExtCmdInd

        RemoveHandler BlueSoleil_Event_HFP_NetworkAvailable, AddressOf bsHandler_HFP_NetworkAvailable
        RemoveHandler BlueSoleil_Event_HFP_NetworkUnavailable, AddressOf bsHandler_HFP_NetworkUnavailable
        RemoveHandler BlueSoleil_Event_HFP_NetworkOperatorName, AddressOf bsHandler_HFP_NetworkName
        RemoveHandler BlueSoleil_Event_HFP_ConnectionReleased, AddressOf bsHandler_HFP_ConnectionReleased
        RemoveHandler BlueSoleil_Event_HFP_CallerID, AddressOf bsHandler_HFP_CallerID
        RemoveHandler BlueSoleil_Event_HFP_SubscriberPhoneNo, AddressOf bsHandler_HFP_SubscriberNo
        RemoveHandler BlueSoleil_Event_HFP_ModelName, AddressOf bsHandler_HFP_ModelName
        RemoveHandler BlueSoleil_Event_HFP_ManufacturerName, AddressOf bsHandler_HFP_ManufacturerName

        RemoveHandler BlueSoleil_Event_HFP_StartRoaming, AddressOf bsHandler_HFP_StartRoaming
        RemoveHandler BlueSoleil_Event_HFP_StopRoaming, AddressOf bsHandler_HFP_EndRoaming



    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click

        BlueSoleil_StopBlueTooth()

        MsgBox("Done.")

    End Sub

    Private Sub chkGetMessagesUnread_CheckedChanged(sender As Object, e As EventArgs) Handles chkGetMessagesUnread.CheckedChanged

        If chkGetMessagesUnread.Checked = False Then
            MsgBox("Downloading the list of ALL read and unread messages will usually take up to a minute (or more).")
        End If

    End Sub

    Private Sub btnHandsFreeConnect_Click(sender As Object, e As EventArgs) Handles btnHandsFreeConnect.Click

        Dim phoneDvcName As String = cboPhoneList.Text



        'BlueSoleil_ConnectService_ByName("A2DP", phoneDvcName, a2dpHandleDvc, a2dpHandleConn, a2dpHandleSvc)


        hfpHandle_LocalSvc_HFAG = BlueSoleil_HFP_RegisterService_HandsFreeAudioGateway("TestBS_HFAG")
        hfpHandle_LocalSvc_HFunit = BlueSoleil_HFP_RegisterService_HandsFreeUnit("TestBS_HFunit")


        ' hfpHandle_LocalSvc_HSunit = BlueSoleil_HFP_RegisterService_HeadSetUnit("TestBS_HSunit")
        ' hfpHandle_LocalSvc_HSAG = BlueSoleil_HFP_RegisterService_HeadSetAudioGateway("TestBS_HSAG")



        BlueSoleil_HFP_RegisterCallbacks()

        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("HFP", phoneDvcName, hfpHandleDvc, hfpHandleConnHFAG, hfpHandleSvc)

        If hfpHandleConnHFAG <> 0 Then


            My.Application.DoEvents()
            Threading.Thread.Sleep(150)



            BlueSoleil_HFP_SendRequest_GetSubscriberNumber(hfpHandleConnHFAG)
            My.Application.DoEvents()
            Threading.Thread.Sleep(150)

            BlueSoleil_HFP_SendRequest_GetNetworkOperator(hfpHandleConnHFAG)
            My.Application.DoEvents()
            '  Threading.Thread.Sleep(150)

            Dim volPct As Double = 100 * TrackBar1.Value / 15
            BlueSoleil_HFP_SetSpeakerVol(hfpHandleConnHFAG, volPct)
            My.Application.DoEvents()
            '  Threading.Thread.Sleep(150)

            BlueSoleil_HFP_GetManufacturer(hfpHandleConnHFAG, hfpManufacturerName)
            My.Application.DoEvents()
            '     Threading.Thread.Sleep(150)

            BlueSoleil_HFP_GetModel(hfpHandleConnHFAG, hfpModelName)
            My.Application.DoEvents()
            '    Threading.Thread.Sleep(150)

            If hfpModelName = "" Then
                hfpModelName = BlueSoleil_GetRemoteDeviceName(hfpHandleDvc)
            End If
        Else

            BlueSoleil_HFP_UnregisterCallbacks()
        End If

        MsgBox("Done.  Return = " & (hfpHandleConnHFAG <> 0))



    End Sub

    Private Sub btnHandsFreeDisconnect_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDisconnect.Click


        BlueSoleil_DisconnectServiceConn(hfpHandleConnHFAG)
        BlueSoleil_DisconnectServiceConn(hfpHandleConnHSAG)
        hfpHandleConnHFAG = 0
        hfpHandleConnHSAG = 0

        hfpHandleDvc = 0
        hfpStatusStr = "Not Connected"

        BlueSoleil_HFP_UnregisterService(hfpHandle_LocalSvc_HFunit)
        BlueSoleil_HFP_UnregisterService(hfpHandle_LocalSvc_HFAG)

        BlueSoleil_HFP_UnregisterService(hfpHandle_LocalSvc_HSunit)
        BlueSoleil_HFP_UnregisterService(hfpHandle_LocalSvc_HSAG)

        hfpHandle_LocalSvc_HFAG = 0
        hfpHandle_LocalSvc_HSAG = 0
        hfpHandle_LocalSvc_HFunit = 0
        hfpHandle_LocalSvc_HSunit = 0

        BlueSoleil_HFP_UnregisterCallbacks()



        MsgBox("Done.")

    End Sub

    Private Sub btnHandsFreeDial_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDial.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_Dial(hfpHandleConnHFAG, tbxHandsFreePhoneNo.Text)


    End Sub

    Private Sub btnHandsFreeHangUp_Click(sender As Object, e As EventArgs) Handles btnHandsFreeHangUp.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_HangUp(hfpHandleConnHFAG)

    End Sub

    Private Sub btnHandsFreeAnswer_Click(sender As Object, e As EventArgs) Handles btnHandsFreeAnswer.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_AnswerCall(hfpHandleConnHFAG)

    End Sub

    Private Sub btnHandsFreeDTMF_1_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_1.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "1")

    End Sub

    Private Sub btnHandsFreeDTMF_2_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_2.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "2")

    End Sub

    Private Sub btnHandsFreeDTMF_3_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_3.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "3")

    End Sub

    Private Sub btnHandsFreeDTMF_4_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_4.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "4")

    End Sub

    Private Sub btnHandsFreeDTMF_5_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_5.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "5")

    End Sub

    Private Sub btnHandsFreeDTMF_6_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_6.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "6")

    End Sub

    Private Sub btnHandsFreeDTMF_7_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_7.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "7")

    End Sub

    Private Sub btnHandsFreeDTMF_8_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_8.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "8")

    End Sub

    Private Sub btnHandsFreeDTMF_9_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_9.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "9")

    End Sub

    Private Sub btnHandsFreeDTMF_Star_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_Star.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "*")

    End Sub

    Private Sub btnHandsFreeDTMF_0_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_0.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "0")

    End Sub

    Private Sub btnHandsFreeDTMF_Hash_Click(sender As Object, e As EventArgs) Handles btnHandsFreeDTMF_Hash.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_SendDTMF(hfpHandleConnHFAG, "#")

    End Sub

    Private Sub Form1_GiveFeedback(sender As Object, e As GiveFeedbackEventArgs) Handles Me.GiveFeedback

    End Sub

    Private Sub btnMsgConnect_Click(sender As Object, e As EventArgs) Handles btnMsgConnect.Click

        Dim phoneDvcName As String = cboPhoneList.Text

        'Dim masBTpath As String = My.Application.Info.DirectoryPath
        'If Strings.Right(masBTpath, 1) <> "\" Then masBTpath = masBTpath & "\"
        'masBTpath = masBTpath & "MASserver"
        'BlueSoleil_MAP_RegisterServers(mapHandleMNSsvc, mapHandleMASsvc, masBTpath)

        mapHandleMNSSvc = BlueSoleil_MAP_RegisterNotificationService()

        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("MAP", phoneDvcName, mapHandleDvc, mapHandleConn, mapHandleSvc)


        If mapHandleConn <> 0 Then

            BlueSoleil_MAP_EnableNotifications(mapHandleConn, True)
        End If

        If mapHandleConn <> 0 Then
            MsgBox("Done.  Return = " & (mapHandleConn <> 0) & ".  If you receive a text message, a notification *should* pop up.")

        Else

            BlueSoleil_MAP_UnregisterServers(mapHandleMNSsvc, 0)

            MsgBox("Done.  Return = " & (mapHandleConn <> 0))

        End If


    End Sub

    Private Sub btnMsgDisconnect_Click(sender As Object, e As EventArgs) Handles btnMsgDisconnect.Click

        BlueSoleil_MAP_EnableNotifications(mapHandleConn, False)

        BlueSoleil_DisconnectServiceConn(mapHandleConn)
        mapHandleConn = 0

        BlueSoleil_MAP_UnregisterServers(mapHandleMNSsvc, 0)


        MsgBox("Done.")

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        If handlersDone = True Then Exit Sub
        handlersDone = True

        AddHandler BlueSoleil_Status_TurnOn, AddressOf bsHandler_Status_TurnOn
        AddHandler BlueSoleil_Status_TurnOff, AddressOf bsHandler_Status_TurnOff
        AddHandler BlueSoleil_Status_Plugged, AddressOf bsHandler_Status_Plugged
        AddHandler BlueSoleil_Status_Unplugged, AddressOf bsHandler_Status_Unplugged
        AddHandler BlueSoleil_Status_DevicePaired, AddressOf bsHandler_Status_DevicePaired
        AddHandler BlueSoleil_Status_DeviceUnpaired, AddressOf bsHandler_Status_DeviceUnpaired
        AddHandler BlueSoleil_Status_DeviceDeleted, AddressOf bsHandler_Status_DeviceDeleted
        AddHandler BlueSoleil_Status_DeviceFound, AddressOf bsHandler_Status_DeviceFound

        AddHandler BlueSoleil_Status_ServiceConnectedInbound, AddressOf bsHandler_Status_ServiceConnectedInbound
        AddHandler BlueSoleil_Status_ServiceDisconnectedInbound, AddressOf bsHandler_Status_ServiceDisconnectedInbound
        AddHandler BlueSoleil_Status_ServiceConnectedOutbound, AddressOf bsHandler_Status_ServiceConnectedOutbound
        AddHandler BlueSoleil_Status_ServiceDisconnectedOutbound, AddressOf bsHandler_Status_ServiceDisconnectedOutbound


        AddHandler BlueSoleil_Event_FTP_FoundFolder, AddressOf bsHandler_FTP_FoundFolder
        AddHandler BlueSoleil_Event_FTP_FoundFile, AddressOf bsHandler_FTP_FoundFile

        AddHandler BlueSoleil_Event_MAP_MsgNotification, AddressOf bsHandler_MAP_MsgNotification

        AddHandler BlueSoleil_Event_PAN_IPchanged, AddressOf bsHandler_PAN_IPchange

        AddHandler BlueSoleil_Event_AVRCP_PlayStatus, AddressOf bsHandler_AVRCP_PlayStatusInfo
        AddHandler BlueSoleil_Event_AVRCP_TrackAlbum, AddressOf bsHandler_AVRCP_TrackAlbum
        AddHandler BlueSoleil_Event_AVRCP_TrackArtist, AddressOf bsHandler_AVRCP_TrackArtist
        AddHandler BlueSoleil_Event_AVRCP_TrackTitle, AddressOf bsHandler_AVRCP_TrackTitle
        AddHandler BlueSoleil_Event_AVRCP_TrackChanged, AddressOf bsHandler_AVRCP_TrackChanged
        AddHandler BlueSoleil_Event_AVRCP_AbsoluteVolume, AddressOf bsHandler_AVRCP_AbsoluteVolume
        AddHandler BlueSoleil_Event_AVRCP_BatteryStatusChanged, AddressOf bsHandler_AVRCP_BatteryStatus

        AddHandler BlueSoleil_Event_HFP_Ringing, AddressOf bsHandler_HFP_Ringing
        AddHandler BlueSoleil_Event_HFP_OngoingCall, AddressOf bsHandler_HFP_OngoingCall
        AddHandler BlueSoleil_Event_HFP_OutgoingCall, AddressOf bsHandler_HFP_OutgoingCall
        AddHandler BlueSoleil_Event_HFP_Standby, AddressOf bsHandler_HFP_StandBy
        AddHandler BlueSoleil_Event_HFP_SignalQuality, AddressOf bsHandler_HFP_SignalQuality
        AddHandler BlueSoleil_Event_HFP_BatteryCharge, AddressOf bsHandler_HFP_BatteryCharge
        AddHandler BlueSoleil_Event_HFP_SpeakerVolume, AddressOf bsHandler_HFP_SpeakerVolume
        AddHandler BlueSoleil_Event_HFP_MicVolume, AddressOf bsHandler_HFP_MicVolume
        AddHandler BlueSoleil_Event_HFP_ExtCmdInd, AddressOf bsHandler_HFP_ExtCmdInd



        AddHandler BlueSoleil_Event_HFP_NetworkAvailable, AddressOf bsHandler_HFP_NetworkAvailable
        AddHandler BlueSoleil_Event_HFP_NetworkUnavailable, AddressOf bsHandler_HFP_NetworkUnavailable
        AddHandler BlueSoleil_Event_HFP_NetworkOperatorName, AddressOf bsHandler_HFP_NetworkName
        AddHandler BlueSoleil_Event_HFP_ConnectionReleased, AddressOf bsHandler_HFP_ConnectionReleased
        AddHandler BlueSoleil_Event_HFP_CallerID, AddressOf bsHandler_HFP_CallerID
        AddHandler BlueSoleil_Event_HFP_SubscriberPhoneNo, AddressOf bsHandler_HFP_SubscriberNo
        AddHandler BlueSoleil_Event_HFP_ModelName, AddressOf bsHandler_HFP_ModelName
        AddHandler BlueSoleil_Event_HFP_ManufacturerName, AddressOf bsHandler_HFP_ManufacturerName

        AddHandler BlueSoleil_Event_HFP_StartRoaming, AddressOf bsHandler_HFP_StartRoaming
        AddHandler BlueSoleil_Event_HFP_StopRoaming, AddressOf bsHandler_HFP_EndRoaming


        If cboPhoneBooks.Items.Count < 1 Then
            cboPhoneBooks.Items.Add("telecom/pb.vcf")
            cboPhoneBooks.Items.Add("telecom/ich.vcf")
            cboPhoneBooks.Items.Add("telecom/och.vcf")
            cboPhoneBooks.Items.Add("telecom/mch.vcf")
            cboPhoneBooks.Items.Add("telecom/cch.vcf")
            cboPhoneBooks.SelectedIndex = 0
        End If




    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        BlueSoleil_AVRCP_SendCmd_Power(avrcpHandleDvc)


    End Sub

    Private Sub Form1_InputLanguageChanged(sender As Object, e As InputLanguageChangedEventArgs) Handles Me.InputLanguageChanged

    End Sub

    Private Sub Form1_RightToLeftLayoutChanged(sender As Object, e As EventArgs) Handles Me.RightToLeftLayoutChanged

    End Sub

    Private Sub btnMediaMute_Click_1(sender As Object, e As EventArgs) Handles btnMediaMute.Click

        BlueSoleil_AVRCP_SendCmd_Mute(avrcpHandleDvc)

    End Sub

    Private Sub btnMediaVolUp_Click(sender As Object, e As EventArgs) Handles btnMediaVolUp.Click

        BlueSoleil_AVRCP_SendCmd_VolumeUp(avrcpHandleDvc)

    End Sub

    Private Sub btnMediaVolDown_Click(sender As Object, e As EventArgs) Handles btnMediaVolDown.Click

        BlueSoleil_AVRCP_SendCmd_VolumeDown(avrcpHandleDvc)

    End Sub

    Private Sub btnHandsFreeTransfer_Click(sender As Object, e As EventArgs) Handles btnHandsFreeTransfer.Click

        BlueSoleil_HFP_TransferAudioConnection(hfpHandleConnHFAG)

    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll

    End Sub

    Private Sub TrackBar1_ValueChanged(sender As Object, e As EventArgs) Handles TrackBar1.ValueChanged



    End Sub

    Private Sub TrackBar1_ParentChanged(sender As Object, e As EventArgs) Handles TrackBar1.ParentChanged

    End Sub

    Private Sub TrackBar1_MouseUp(sender As Object, e As MouseEventArgs) Handles TrackBar1.MouseUp

        Dim volPct As Double = 100 * TrackBar1.Value / 15
        BlueSoleil_HFP_SetSpeakerVol(hfpHandleConnHFAG, volPct)

    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll

    End Sub

    Private Sub TrackBar2_MouseUp(sender As Object, e As MouseEventArgs) Handles TrackBar2.MouseUp

        Dim volPct As Double = 100 * TrackBar1.Value / 15
        BlueSoleil_AVRCP_SendReq_SetAbsoluteVolumePct(avrcpHandleDvc, volPct)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub Button4_MouseDown(sender As Object, e As MouseEventArgs) Handles Button4.MouseDown

        hfpVoiceCmdStateEnabled = True
        Dim retBool As Boolean = BlueSoleil_HFP_SetVoiceRecognitionState(hfpHandleConnHFAG, hfpVoiceCmdStateEnabled)

    End Sub

    Private Sub Button4_MouseUp(sender As Object, e As MouseEventArgs) Handles Button4.MouseUp

        hfpVoiceCmdStateEnabled = False
        Dim retBool As Boolean = BlueSoleil_HFP_SetVoiceRecognitionState(hfpHandleConnHFAG, hfpVoiceCmdStateEnabled)

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub cboPhoneList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPhoneList.SelectedIndexChanged

        'get device handle, store it.

        Dim i As Integer
        For i = 0 To dvcArrayCount - 1
            If dvcNameArray(i) = cboPhoneList.Text Then
                dvcCurrHandle = dvcHandleArray(i)

                Dim dvcAddress As String = BlueSoleil_GetRemoteDeviceAddress(dvcCurrHandle)
                lblDeviceAddress.Text = "Device Address:  " & dvcAddress
                dvcCurrAddress = dvcAddress

                Exit For
            End If
        Next i

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles lblDeviceSent.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click




        If dvcCurrHandle = 0 Then Exit Sub



        'get service handles for device.
        Dim svcHandleArray(0 To 0) As UInt32, svcHandleCount As Integer = 0

        Dim TorF As Boolean
        TorF = BlueSoleil_GetRemoteDeviceServiceHandles_Refresh(dvcCurrHandle, svcHandleArray, svcHandleCount)

        Dim svcListStr As String = ""

        lvwRemoteServices.Items.Clear()

        Dim tempItem As ListViewItem = Nothing

        Dim i As Integer, svcName As String = "", svcClass As UShort = 0
        If TorF = True Then
            For i = 0 To svcHandleCount - 1
                TorF = BlueSoleil_GetRemoteServiceAttributes(svcHandleArray(i), svcName, svcClass)

                If svcClass <> 0 And TorF = True Then
                    svcListStr = svcListStr & "0x" & Strings.Right("00000000" & Hex(svcHandleArray(i)), 8) & "    0x" & Strings.Right("0000" & Hex(svcClass), 4) & "    " & svcName & "  " & vbNewLine

                    tempItem = lvwRemoteServices.Items.Add("0x" & Strings.Right("00000000" & Hex(svcHandleArray(i)), 8))
                    tempItem.SubItems.Add("0x" & Strings.Right("0000" & Hex(svcClass), 4))
                    tempItem.SubItems.Add(svcName)

                End If

            Next i

            MsgBox("Found " & svcHandleCount & " remote services." & vbNewLine & vbNewLine & svcListStr)

        Else

            MsgBox("Unable to retrieve service list from remote device.")
        End If

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub lvwMessages_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvwMessages.SelectedIndexChanged

    End Sub

    Private Sub lvwMessages_DoubleClick(sender As Object, e As EventArgs) Handles lvwMessages.DoubleClick

        If lvwMessages.SelectedItems.Count < 1 Then Exit Sub

        Dim YorN As MsgBoxResult = MsgBox("Download message from phone?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Pull Message?")
        If YorN <> MsgBoxResult.Yes Then Exit Sub


        Dim writeToPath As String = My.Application.Info.DirectoryPath
        If Strings.Right(writeToPath, 1) <> "\" Then writeToPath = writeToPath & "\"
        writeToPath = writeToPath & "Inbox"

        If IO.Directory.Exists(writeToPath) = False Then
            Try
                IO.Directory.CreateDirectory(writeToPath)
            Catch ex As Exception

            End Try
        End If


        'get the selected message's handle.
        Dim msgHandleToDownload As String = lvwMessages.SelectedItems(0).SubItems(7).Text


        Dim msgInfo As String = ""
        Dim msgTorF As Boolean = False
        Dim msgFileName As String = writeToPath
        If Strings.Right(msgFileName, 1) <> "\" Then msgFileName = msgFileName & "\"
        msgFileName = msgFileName & msgHandleToDownload & ".BMSG"

        Dim msgText As String = "", msgFromName As String = "", msgFromNo As String = "", msgType As String = "", msgStatus As String = "", msgFolder As String = ""
        Dim msgAttachBytes(0 To 0) As Byte, msgAttachType As String = "", msgAttachSize As Long = 0



        msgTorF = BlueSoleil_MAP_PullMessage(mapHandleConn, msgHandleToDownload, msgFileName, True)
        If msgTorF = True Then
            BMSG_GetMessageInfo(msgFileName, msgText, msgFromName, msgFromNo, msgType, msgStatus, msgFolder, msgAttachBytes, msgAttachType)
            msgAttachSize = msgAttachBytes.Length
            If msgAttachType = "" Then msgAttachSize = 0
            msgInfo = "Message =" & vbNewLine & "From: " & msgFromNo & "  " & msgFromName & vbNewLine & "Msg Text: " & msgText & vbNewLine & "Type: " & msgType & vbNewLine & "Status: " & msgStatus & vbNewLine & "Folder: " & msgFolder & vbNewLine & "Attachment Size: " & msgAttachSize & " bytes"
        End If

        MsgBox("GetMessage = " & msgTorF & vbNewLine & msgInfo)

    End Sub

    Private Sub btnHandsFreeATcmd_Click(sender As Object, e As EventArgs) Handles btnHandsFreeATcmd.Click

        If hfpHandleConnHFAG = 0 Then Exit Sub

        Dim retBool As Boolean = BlueSoleil_HFP_SendATcmd(hfpHandleConnHFAG, tbxHandsFreeATcmd.Text, 1000)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim phoneDvcName As String = cboPhoneList.Text

        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("SPP", phoneDvcName, sppHandleDvc, sppHandleConn, sppHandleSvc)

        sppCOMMportNum = BlueSoleil_SPP_GetCOMMportNum(sppHandleConn)

        MsgBox("Done.  Return = " & (sppHandleConn <> 0) & vbNewLine & vbNewLine & "COMM Port = " & sppCOMMportNum)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        BlueSoleil_DisconnectServiceConn(sppHandleConn)

        MsgBox("Done.")

    End Sub

    Private Sub btnMediaGetLists_Click(sender As Object, e As EventArgs) Handles btnMediaGetNowPlaying.Click

        BlueSoleil_AVRCP_SendReq_GetFolderItems_NowPlayingList(avrcpHandleDvc)

    End Sub

    Private Sub tvwPhoneBook_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvwPhoneBook.AfterSelect

        Dim currNode As TreeNode = e.Node
        Dim parNode As TreeNode = currNode.Parent

        'get top-most node parent.
        Do
            If IsNothing(parNode) = False Then
                currNode = parNode
                parNode = currNode.Parent
            Else
                Exit Do
            End If
        Loop


        ' pbxPhonebookPic.BackgroundImage.Dispose()
        ' pbxPhonebookPic.BackgroundImage = Nothing

        pbxPhonebookPic.Visible = False
        If IsNothing(currNode.Tag) = False Then
            If TypeOf (currNode.Tag) Is Bitmap Then
                pbxPhonebookPic.BackgroundImageLayout = ImageLayout.Center
                pbxPhonebookPic.BackgroundImage = CType(currNode.Tag, Bitmap)
                pbxPhonebookPic.Refresh()
                pbxPhonebookPic.Visible = True
            End If
        End If


    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles btnA2DPconnect.Click

        Dim phoneDvcName As String = cboPhoneList.Text

        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("A2DP", phoneDvcName, a2dpHandleDvc, a2dpHandleConn, a2dpHandleSvc)


        MsgBox("Done.  Return = " & (a2dpHandleConn <> 0))


    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles btnA2DPdisconnect.Click

        BlueSoleil_DisconnectServiceConn(a2dpHandleConn)

        MsgBox("Done.")

    End Sub

    Private Sub TabPage9_Click(sender As Object, e As EventArgs) Handles TabPage9.Click

    End Sub

    Private Sub btnFTPconnect_Click(sender As Object, e As EventArgs) Handles btnFTPconnect.Click

        Dim phoneDvcName As String = cboPhoneList.Text

        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("FTP", phoneDvcName, ftpHandleDvc, ftpHandleConn, ftpHandleSvc)

        If connTorF = True Then
            BlueSoleil_FTP_RegisterStatusCallback(ftpHandleConn)

            BlueSoleil_FTP_GetRemotePath(ftpHandleConn, ftpRemotePath)

            BlueSoleil_FTP_BrowseFolder(ftpHandleConn, ftpRemotePath)
        End If

        MsgBox("Done.  Return = " & (ftpHandleConn <> 0))

    End Sub

    Private Sub btnFTPdisconnect_Click(sender As Object, e As EventArgs) Handles btnFTPdisconnect.Click


        BlueSoleil_FTP_UnregisterStatusCallback(ftpHandleConn)

        BlueSoleil_DisconnectServiceConn(ftpHandleConn)

        MsgBox("Done.")

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles btnFTPdelete.Click

        Dim selItem As ListViewItem = Nothing

        If lvwFTPbrowser.SelectedItems.Count < 1 Then Exit Sub

        selItem = lvwFTPbrowser.SelectedItems(0)

        If selItem.Text = "[..]" Then Exit Sub


        'confirm delete.
        Dim YorN As MsgBoxResult = MsgBox("Are you sure you want to delete " & selItem.Text & " ?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, "Confirm delete")
        If YorN <> MsgBoxResult.Yes Then Exit Sub


        Dim retBool As Boolean = False
        If selItem.SubItems.Count < 2 Or Strings.Left(selItem.Text, 1) = "[" Then
            'folder.  remove brackets.
            Dim delDir As String = Replace(selItem.Text, "[", "")
            delDir = Replace(delDir, "]", "")

            retBool = BlueSoleil_FTP_DeleteDirectory(ftpHandleConn, delDir)

        Else
            'file
            retBool = BlueSoleil_FTP_DeleteFile(ftpHandleConn, selItem.Text)
        End If



        'refresh current folder.
        Dim ftpPath As String = ""
        BlueSoleil_FTP_GetRemotePath(ftpHandleConn, ftpPath)
        ftpPath = Replace(ftpPath, Chr(0), "")
        If ftpPath <> "\" And ftpPath <> "" Then
            lvwFTPbrowser.Items.Add("[..]")
        End If
        BlueSoleil_FTP_BrowseFolder(ftpHandleConn, Strings.Mid(ftpPath, 2))


        MsgBox("Return = " & retBool)

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvwFTPbrowser.SelectedIndexChanged

    End Sub

    Private Sub lvwFTPbrowser_DoubleClick(sender As Object, e As EventArgs) Handles lvwFTPbrowser.DoubleClick

        Dim selItem As ListViewItem = Nothing

        If lvwFTPbrowser.SelectedItems.Count < 1 Then Exit Sub

        selItem = lvwFTPbrowser.SelectedItems(0)

        Dim ftpPath As String = ""

        Dim cdTorF As Boolean = False

        'up one folder.
        If selItem.Text = "[..]" Then

            cdTorF = BlueSoleil_FTP_UpOneFolder(ftpHandleConn)

            If cdTorF = False Then
                'failed to change dir
                MsgBox("Failed to move up one level.")
                Exit Sub
            End If

            lvwFTPbrowser.Items.Clear()
            BlueSoleil_FTP_GetRemotePath(ftpHandleConn, ftpPath)

            If ftpPath <> "\" And ftpPath <> "" Then
                lvwFTPbrowser.Items.Add("[..]")
            End If

            If ftpPath = "\" Then ftpPath = "\\"

            BlueSoleil_FTP_BrowseFolder(ftpHandleConn, Strings.Mid(ftpPath, 2))
            Exit Sub
        End If

        'go to folder.
        If selItem.SubItems.Count < 2 Or Strings.Left(selItem.Text, 1) = "[" Then

            Dim newDir As String = Replace(selItem.Text, "[", "")
            newDir = Replace(newDir, "]", "")

            Dim dirAccessDeniedError As Boolean = False

            cdTorF = BlueSoleil_FTP_SetRemotePath(ftpHandleConn, "\" & newDir, dirAccessDeniedError)
            If cdTorF = False Then
                'failed to change dir

                If dirAccessDeniedError = True Then
                    MsgBox("Failed to change to directory " & newDir & " - Access Denied.")
                Else
                    MsgBox("Failed to change to directory " & newDir)
                End If


                Exit Sub
            End If


            lvwFTPbrowser.Items.Clear()
            If newDir <> "\" And newDir <> "" Then
                lvwFTPbrowser.Items.Add("[..]")
            End If

            BlueSoleil_FTP_BrowseFolder(ftpHandleConn, newDir)
            Exit Sub
        End If

        'double-clicked on file.  download?


    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles btnFTPcreateFolder.Click

        Dim fdrName As String = InputBox("Enter name for new folder:")

        If fdrName = "" Then Exit Sub

        Dim retBool As Boolean = BlueSoleil_FTP_CreateDirectory(ftpHandleConn, fdrName)



        'refresh current folder.
        Dim ftpPath As String = ""
        BlueSoleil_FTP_GetRemotePath(ftpHandleConn, ftpPath)
        ftpPath = Replace(ftpPath, Chr(0), "")
        If ftpPath <> "\" And ftpPath <> "" Then
            lvwFTPbrowser.Items.Add("[..]")
        End If
        BlueSoleil_FTP_BrowseFolder(ftpHandleConn, Strings.Mid(ftpPath, 2))


        MsgBox("Return = " & retBool)

    End Sub

    Private Sub btnFTPupload_Click(sender As Object, e As EventArgs) Handles btnFTPupload.Click

        Dim inpFN As String = Browse_ShowOpenFile(Me, "All files (*.*)|*.*", "Select a file to upload...")

        If inpFN = "" Then Exit Sub

        Dim retBool As Boolean
        retBool = BlueSoleil_FTP_PutFile(ftpHandleConn, inpFN, IO.Path.GetFileName(inpFN))


        'refresh current folder.
        Dim ftpPath As String = ""
        BlueSoleil_FTP_GetRemotePath(ftpHandleConn, ftpPath)
        ftpPath = Replace(ftpPath, Chr(0), "")
        If ftpPath <> "\" And ftpPath <> "" Then
            lvwFTPbrowser.Items.Add("[..]")
        End If
        BlueSoleil_FTP_BrowseFolder(ftpHandleConn, Strings.Mid(ftpPath, 2))


        MsgBox("Return = " & retBool)

    End Sub

    Private Sub btnFTPdownload_Click(sender As Object, e As EventArgs) Handles btnFTPdownload.Click

        Dim selItem As ListViewItem = Nothing

        If lvwFTPbrowser.SelectedItems.Count < 1 Then Exit Sub

        selItem = lvwFTPbrowser.SelectedItems(0)

        If selItem.Text = "[..]" Then
            Exit Sub
        End If

        If selItem.SubItems.Count < 2 Or Strings.Left(selItem.Text, 1) = "[" Then
            Dim getDir As String = Replace(selItem.Text, "[", "")
            getDir = Replace(getDir, "]", "")

            Dim saveToDir As String = Browse_ShowOpenFolder(Me, "Select folder to save to...")
            If saveToDir = "" Then Exit Sub
            Dim retBool_GetFolder As Boolean
            retBool_GetFolder = BlueSoleil_FTP_GetFolder(ftpHandleConn, getDir, saveToDir)

            MsgBox("GetFolder Return = " & retBool_GetFolder)

        Else
            Dim saveFN As String = Browse_ShowSaveFile(Me, "All files (*.*)|*.*", "Save As...", selItem.Text)
            If saveFN = "" Then Exit Sub
            Dim retBool_GetFile As Boolean
            retBool_GetFile = BlueSoleil_FTP_GetFile(ftpHandleConn, IO.Path.GetFileName(saveFN), saveFN)

            MsgBox("GetFile Return = " & retBool_GetFile)

        End If




    End Sub

    Private Sub btnFTPcancelXfer_Click(sender As Object, e As EventArgs) Handles btnFTPcancelXfer.Click

        If ftpHandleConn = 0 Then
            Exit Sub
        End If

        Dim retBool As Boolean = BlueSoleil_FTP_CancelTransfer(ftpHandleConn)

        MsgBox("Return = " & retBool)

    End Sub

    Private Sub btnMediaGetFileSystem_Click(sender As Object, e As EventArgs) Handles btnMediaGetFileSystem.Click

        BlueSoleil_AVRCP_SendReq_GetFolderItems_FileSystem(avrcpHandleDvc)



    End Sub

    Private Sub Button8_Click_2(sender As Object, e As EventArgs) Handles Button8.Click

    End Sub

    Private Sub btnMediaEnableBrowsing_Click(sender As Object, e As EventArgs) Handles btnMediaEnableBrowsing.Click

        Dim retBool As Boolean = BlueSoleil_AVRCP_EnableBrowsing(avrcpHandleDvc)

        MsgBox("Done.  Return = " & retBool)

    End Sub

    Private Sub tbxA2DPinfo_TextChanged(sender As Object, e As EventArgs) Handles tbxA2DPinfo.TextChanged

    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles btnOPPpush.Click

        If Trim(tbxOPPfullname.Text) = "" Then
            MsgBox("You must enter the name of the contact.")
            Exit Sub
        End If

        'create path to write temporary vcard file to.
        Dim writeToPath As String = My.Application.Info.DirectoryPath
        If Strings.Right(writeToPath, 1) <> "\" Then writeToPath = writeToPath & "\"
        writeToPath = writeToPath & "Obex"

        If IO.Directory.Exists(writeToPath) = False Then
            Try
                IO.Directory.CreateDirectory(writeToPath)
            Catch ex As Exception

            End Try
        End If

        Dim vcfFileName As String = writeToPath & "\" & "SendMe.VCF"       'arbitrary name.  could use contact fullname.

        'remove old/previous vcard file if it exists.
        Try
            If IO.File.Exists(vcfFileName) = True Then
                IO.File.Delete(vcfFileName)
            End If

        Catch ex As Exception
            MsgBox("Unable to delete old SendMe.VCF file.")
            Exit Sub
        End Try


        'prepare variables.
        Dim vcfFullName As String = Trim(tbxOPPfullname.Text)
        Dim vcfEMail As String = Trim(tbxOPPemail.Text)
        Dim vcfImageFN As String = Trim(tbxOPPimageFN.Text)
        Dim vcfNotes As String = Trim(tbxOPPnotes.Text)
        Dim vcfImageObj As Bitmap = Nothing
        Dim vcfCompany As String = Trim(tbxOPPcompany.Text)
        Dim vcfPhoneCell As String = Trim(tbxOPPphoneCell.Text)
        Dim vcfPhoneWork As String = Trim(tbxOPPphoneWork.Text)
        Dim vcfPhoneHome As String = Trim(tbxOPPphoneHome.Text)
        Dim vcfPhoneNumberCount As Integer = 0
        Dim vcfAddressWork As String = Trim(tbxOPPaddrWork.Text)
        Dim vcfAddressHome As String = Trim(tbxOPPaddrHome.Text)
        Dim vcfAddressCount As Integer = 0

        'prepare array of phone numbers and labels
        Dim vcfPhoneNumberArray(0 To 0) As String, vcfPhoneLabelArray(0 To 0) As String
        If vcfPhoneCell <> "" Then
            ReDim Preserve vcfPhoneNumberArray(0 To vcfPhoneNumberCount)
            ReDim Preserve vcfPhoneLabelArray(0 To vcfPhoneNumberCount)
            vcfPhoneNumberArray(vcfPhoneNumberCount) = vcfPhoneCell
            vcfPhoneLabelArray(vcfPhoneNumberCount) = "CELL"
            vcfPhoneNumberCount = vcfPhoneNumberCount + 1
        End If
        If vcfPhoneWork <> "" Then
            ReDim Preserve vcfPhoneNumberArray(0 To vcfPhoneNumberCount)
            ReDim Preserve vcfPhoneLabelArray(0 To vcfPhoneNumberCount)
            vcfPhoneNumberArray(vcfPhoneNumberCount) = vcfPhoneWork
            vcfPhoneLabelArray(vcfPhoneNumberCount) = "WORK"
            vcfPhoneNumberCount = vcfPhoneNumberCount + 1
        End If
        If vcfPhoneHome <> "" Then
            ReDim Preserve vcfPhoneNumberArray(0 To vcfPhoneNumberCount)
            ReDim Preserve vcfPhoneLabelArray(0 To vcfPhoneNumberCount)
            vcfPhoneNumberArray(vcfPhoneNumberCount) = vcfPhoneHome
            vcfPhoneLabelArray(vcfPhoneNumberCount) = "HOME"
            vcfPhoneNumberCount = vcfPhoneNumberCount + 1
        End If

        'prepare array of addresses and labels.
        Dim vcfAddressArray(0 To 0) As String, vcfAddressLabelArray(0 To 0) As String
        If vcfAddressWork <> "" Then
            ReDim Preserve vcfAddressArray(0 To vcfAddressCount)
            ReDim Preserve vcfAddressLabelArray(0 To vcfAddressCount)
            vcfAddressArray(vcfAddressCount) = vcfAddressWork
            vcfAddressLabelArray(vcfAddressCount) = "WORK"
            vcfAddressCount = vcfAddressCount + 1
        End If
        If vcfAddressHome <> "" Then
            ReDim Preserve vcfAddressArray(0 To vcfAddressCount)
            ReDim Preserve vcfAddressLabelArray(0 To vcfAddressCount)
            vcfAddressArray(vcfAddressCount) = vcfAddressHome
            vcfAddressLabelArray(vcfAddressCount) = "HOME"
            vcfAddressCount = vcfAddressCount + 1
        End If

        'prepare the image object.
        If vcfImageFN <> "" Then
            Try
                If IO.File.Exists(vcfImageFN) = True Then
                    vcfImageObj = CType(Image.FromFile(vcfImageFN), Bitmap)
                End If

            Catch ex As Exception

            End Try
        End If

        'write the vcard file.
        Dim writeTorF As Boolean = VCard_WriteContactInfo_V3(vcfFileName, vcfFullName, vcfPhoneNumberArray, vcfPhoneLabelArray, vcfPhoneNumberCount, vcfAddressArray, vcfAddressLabelArray, vcfAddressCount, vcfImageObj, vcfEMail, vcfCompany, vcfNotes)

        If IsNothing(vcfImageObj) = False Then
            vcfImageObj.Dispose()
        End If




        Dim phoneDvcName As String = cboPhoneList.Text

        'connect to OPP.
        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("OPP", phoneDvcName, oppHandleDvc, oppHandleConn, oppHandleSvc)

        'register OPP status callback?  meh.  don't really need it for simple connect-push-disconnect.

        'push the vcard file.
        Dim pushTorF As Boolean

        If connTorF = True Then

            'the callback doesn't do anything special for us.
            BlueSoleil_OPP_RegisterStatusCallback(oppHandleConn)

            pushTorF = BlueSoleil_OPP_PushVCard(oppHandleConn, vcfFileName)

            'unregister status callback?  meh.
            BlueSoleil_OPP_UnregisterStatusCallback(oppHandleConn)


            'disconnect.
            BlueSoleil_DisconnectServiceConn(oppHandleConn)
        End If


        MsgBox("Connect = " & connTorF & "   PushCard = " & pushTorF)

    End Sub

    Private Sub btnOPPimageBrowse_Click(sender As Object, e As EventArgs) Handles btnOPPimageBrowse.Click

        Dim inpFN As String = Browse_ShowOpenFile(Me, "Image files|*.jpg;*.png;*.bmp;*.gif;*.tif", "Select a file to upload...")

        If inpFN = "" Then Exit Sub

        tbxOPPimageFN.Text = inpFN

    End Sub

    Private Sub Button9_Click_2(sender As Object, e As EventArgs) Handles btnOPPpull.Click

        'create path to write temporary vcard file to.
        Dim writeToPath As String = My.Application.Info.DirectoryPath
        If Strings.Right(writeToPath, 1) <> "\" Then writeToPath = writeToPath & "\"
        writeToPath = writeToPath & "Obex"

        If IO.Directory.Exists(writeToPath) = False Then
            Try
                IO.Directory.CreateDirectory(writeToPath)
            Catch ex As Exception

            End Try
        End If

        Dim vcfFileName As String = writeToPath     'file always gets saved as REMOTE.VCF


        Dim phoneDvcName As String = cboPhoneList.Text

        'connect to OPP.
        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("OPP", phoneDvcName, oppHandleDvc, oppHandleConn, oppHandleSvc)


        'push the vcard file.
        Dim pullTorF As Boolean = False
        Dim pullAccessDenied As Boolean = False
        If connTorF = True Then

            'register OPP status callback?  meh.  don't really need it for simple connect-pull-disconnect.
            BlueSoleil_OPP_RegisterStatusCallback(oppHandleConn)

            pullTorF = BlueSoleil_OPP_PullVCard(oppHandleConn, vcfFileName, pullAccessDenied)

            'unregister status callback?  meh.
            BlueSoleil_OPP_UnregisterStatusCallback(oppHandleConn)

            'disconnect.
            BlueSoleil_DisconnectServiceConn(oppHandleConn)
        End If

        If pullTorF = True Then

            Dim cardName As String = "", cardPhoneNumbers(0 To 0) As String, cardPhoneLabels(0 To 0) As String, cardPhoneCount As Integer = 0
            Dim cardImage As Bitmap = Nothing
            Dim tempImage As Bitmap = Nothing
            Dim cardEMail As String = "", cardLastCallDateTime As DateTime = Nothing, cardLastCallType As String = "", cardAddresses(0 To 0) As String, cardAddressLabels(0 To 0) As String, cardAddressCount As Integer, cardBirthday As String = "", cardGeoPos As String = "", cardNotes As String = "", cardOrganization As String = ""

            'we know the first contact is at offset 0.  However, we could call VCard_GetContactOffsets if we really wanted to, and use the offset from that.
            VCard_GetContactInfo(vcfFileName, 0, cardName, cardEMail, cardPhoneNumbers, cardPhoneLabels, cardPhoneCount, tempImage, cardLastCallDateTime, cardLastCallType, cardAddresses, cardAddressLabels, cardAddressCount, cardBirthday, cardGeoPos, cardNotes, cardOrganization)

            Dim i As Integer
            Dim cardInfoStr As String = "Name:    " & cardName & vbNewLine
            If cardEMail <> "" Then cardInfoStr = cardInfoStr & "E-Mail:   " & cardEMail & vbNewLine
            If cardOrganization <> "" Then cardInfoStr = cardInfoStr & "Company:   " & cardOrganization & vbNewLine
            For i = 0 To cardPhoneCount - 1
                cardInfoStr = cardInfoStr & cardPhoneLabels(i) & " Phone: " & cardPhoneNumbers(i) & vbNewLine
            Next i
            For i = 0 To cardPhoneCount - 1
                cardInfoStr = cardInfoStr & cardAddressLabels(i) & " Addr: " & cardAddresses(i) & vbNewLine
            Next i

            cardInfoStr = cardInfoStr & vbNewLine & "The VCard is saved as " & vcfFileName

            MsgBox("Connect = " & connTorF & "   PullCard = " & pullTorF & vbNewLine & vbNewLine & cardInfoStr)

        Else

            If pullAccessDenied = True Then
                MsgBox("Connect = " & connTorF & "   PullCard = " & pullTorF & "  - Access Denied")
            Else
                MsgBox("Connect = " & connTorF & "   PullCard = " & pullTorF)
            End If

        End If


    End Sub

    Private Sub Button9_Click_3(sender As Object, e As EventArgs) Handles Button9.Click

        If Trim(tbxOPPfullname.Text) = "" Then
            MsgBox("You must enter the name of the contact.")
            Exit Sub
        End If

        'create path to write temporary vcard file to.
        Dim writeToPath As String = My.Application.Info.DirectoryPath
        If Strings.Right(writeToPath, 1) <> "\" Then writeToPath = writeToPath & "\"
        writeToPath = writeToPath & "Obex"

        If IO.Directory.Exists(writeToPath) = False Then
            Try
                IO.Directory.CreateDirectory(writeToPath)
            Catch ex As Exception

            End Try
        End If

        Dim vcfPushFileName As String = writeToPath & "\" & "SendMe.VCF"       'arbitrary name.  could use contact fullname.

        Try
            If IO.File.Exists(vcfPushFileName) = True Then
                IO.File.Delete(vcfPushFileName)
            End If

        Catch ex As Exception
            MsgBox("Unable to delete old SendMe.VCF file.")
            Exit Sub
        End Try


        Dim vcfPullPath As String = writeToPath
        Try
            If IO.File.Exists(vcfPullPath & "\" & "remote.vcf") = True Then
                IO.File.Delete(vcfPullPath & "\" & "remote.vcf")
            End If

        Catch ex As Exception
            MsgBox("Unable to delete old remote.VCF file.")
            Exit Sub
        End Try


        'prepare variables.
        Dim vcfFullName As String = Trim(tbxOPPfullname.Text)
        Dim vcfEMail As String = Trim(tbxOPPemail.Text)
        Dim vcfImageFN As String = Trim(tbxOPPimageFN.Text)
        Dim vcfImageObj As Bitmap = Nothing
        Dim vcfCompany As String = Trim(tbxOPPcompany.Text)
        Dim vcfPhoneCell As String = Trim(tbxOPPphoneCell.Text)
        Dim vcfPhoneWork As String = Trim(tbxOPPphoneWork.Text)
        Dim vcfPhoneHome As String = Trim(tbxOPPphoneHome.Text)
        Dim vcfPhoneNumberCount As Integer = 0
        Dim vcfAddressWork As String = Trim(tbxOPPaddrWork.Text)
        Dim vcfAddressHome As String = Trim(tbxOPPaddrHome.Text)
        Dim vcfAddressCount As Integer = 0

        'prepare array of phone numbers and labels
        Dim vcfPhoneNumberArray(0 To 0) As String, vcfPhoneLabelArray(0 To 0) As String
        If vcfPhoneCell <> "" Then
            ReDim Preserve vcfPhoneNumberArray(0 To vcfPhoneNumberCount)
            ReDim Preserve vcfPhoneLabelArray(0 To vcfPhoneNumberCount)
            vcfPhoneNumberArray(vcfPhoneNumberCount) = vcfPhoneCell
            vcfPhoneLabelArray(vcfPhoneNumberCount) = "CELL"
            vcfPhoneNumberCount = vcfPhoneNumberCount + 1
        End If
        If vcfPhoneWork <> "" Then
            ReDim Preserve vcfPhoneNumberArray(0 To vcfPhoneNumberCount)
            ReDim Preserve vcfPhoneLabelArray(0 To vcfPhoneNumberCount)
            vcfPhoneNumberArray(vcfPhoneNumberCount) = vcfPhoneWork
            vcfPhoneLabelArray(vcfPhoneNumberCount) = "WORK"
            vcfPhoneNumberCount = vcfPhoneNumberCount + 1
        End If
        If vcfPhoneHome <> "" Then
            ReDim Preserve vcfPhoneNumberArray(0 To vcfPhoneNumberCount)
            ReDim Preserve vcfPhoneLabelArray(0 To vcfPhoneNumberCount)
            vcfPhoneNumberArray(vcfPhoneNumberCount) = vcfPhoneHome
            vcfPhoneLabelArray(vcfPhoneNumberCount) = "HOME"
            vcfPhoneNumberCount = vcfPhoneNumberCount + 1
        End If

        'prepare array of addresses and labels.
        Dim vcfAddressArray(0 To 0) As String, vcfAddressLabelArray(0 To 0) As String
        If vcfAddressWork <> "" Then
            ReDim Preserve vcfAddressArray(0 To vcfAddressCount)
            ReDim Preserve vcfAddressLabelArray(0 To vcfAddressCount)
            vcfAddressArray(vcfAddressCount) = vcfAddressWork
            vcfAddressLabelArray(vcfAddressCount) = "WORK"
            vcfAddressCount = vcfAddressCount + 1
        End If
        If vcfAddressHome <> "" Then
            ReDim Preserve vcfAddressArray(0 To vcfAddressCount)
            ReDim Preserve vcfAddressLabelArray(0 To vcfAddressCount)
            vcfAddressArray(vcfAddressCount) = vcfAddressHome
            vcfAddressLabelArray(vcfAddressCount) = "HOME"
            vcfAddressCount = vcfAddressCount + 1
        End If

        'prepare the image object.
        Try
            vcfImageObj = CType(Image.FromFile(vcfImageFN), Bitmap)
        Catch ex As Exception

        End Try

        'write the vcard file.
        Dim writeTorF As Boolean = VCard_WriteContactInfo_V3(vcfPushFileName, vcfFullName, vcfPhoneNumberArray, vcfPhoneLabelArray, vcfPhoneNumberCount, vcfAddressArray, vcfAddressLabelArray, vcfAddressCount, vcfImageObj, vcfEMail, vcfCompany)

        If IsNothing(vcfImageObj) = False Then
            vcfImageObj.Dispose()
        End If




        Dim phoneDvcName As String = cboPhoneList.Text

        'connect to OPP.
        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("OPP", phoneDvcName, oppHandleDvc, oppHandleConn, oppHandleSvc)


        'push the vcard file.
        Dim pushTorF As Boolean, pullTorF As Boolean, exchTorF As Boolean

        If connTorF = True Then
            BlueSoleil_OPP_RegisterStatusCallback(oppHandleConn)

            exchTorF = BlueSoleil_OPP_ExchangeVCards(oppHandleConn, vcfPushFileName, vcfPullPath, pushTorF, pullTorF)

            'unregister status callback?  meh.
            BlueSoleil_OPP_UnregisterStatusCallback(oppHandleConn)

            'disconnect.
            BlueSoleil_DisconnectServiceConn(oppHandleConn)
        End If


        MsgBox("Connect = " & connTorF & "   ObjExchange = " & exchTorF & "    Push = " & pushTorF & "    Pull = " & pullTorF)


    End Sub

    Private Sub lvwRemoteServices_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvwRemoteServices.SelectedIndexChanged

    End Sub

    Private Sub lvwLocalServices_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvwLocalServices.SelectedIndexChanged

    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click



        'get service handles for device.
        Dim svcHandleArray(0 To 0) As UInt32, svcClassArray(0 To 0) As UInt16, svcNameArray(0 To 0) As String, svcHandleCount As Integer = 0

        Dim TorF As Boolean
        TorF = BlueSoleil_GetLocalDeviceServices(svcClassArray, svcHandleArray, svcNameArray, svcHandleCount)

        Dim svcListStr As String = ""

        lvwLocalServices.Items.Clear()

        Dim tempItem As ListViewItem = Nothing

        Dim i As Integer, svcName As String = "", svcClass As UShort = 0
        If TorF = True Then
            For i = 0 To svcHandleCount - 1

                svcName = svcNameArray(i)
                svcClass = svcClassArray(i)

                If svcClass <> 0 Then
                    svcListStr = svcListStr & "0x" & Strings.Right("00000000" & Hex(svcHandleArray(i)), 8) & "    0x" & Strings.Right("0000" & Hex(svcClass), 4) & "    " & svcName & "  " & vbNewLine

                    tempItem = lvwLocalServices.Items.Add("0x" & Strings.Right("00000000" & Hex(svcHandleArray(i)), 8))
                    tempItem.SubItems.Add("0x" & Strings.Right("0000" & Hex(svcClass), 4))
                    tempItem.SubItems.Add(svcName)

                End If

            Next i

            MsgBox("Found " & svcHandleCount & " local services." & vbNewLine & vbNewLine & svcListStr)

        Else

            MsgBox("Unable to retrieve service list from local device.")
        End If

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnGetCardList_Click(sender As Object, e As EventArgs) Handles btnGetCardList.Click


        Dim YorN As MsgBoxResult = MsgBox("Pull CardList and each individual VCard separately?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, "Confirm CardList Pull")
        If YorN <> MsgBoxResult.Yes Then Exit Sub


        Dim writeToPath As String = My.Application.Info.DirectoryPath
        If Strings.Right(writeToPath, 1) <> "\" Then writeToPath = writeToPath & "\"
        writeToPath = writeToPath & "Contacts"

        Try
            If IO.Directory.Exists(writeToPath) = False Then
                IO.Directory.CreateDirectory(writeToPath)
            End If
        Catch ex As Exception

        End Try

        Select Case LCase(cboPhoneBooks.Text)
            Case ("telecom/pb.vcf")

            Case ("telecom/ich.vcf")
                writeToPath = writeToPath & "\History_Incoming"
            Case ("telecom/och.vcf")
                writeToPath = writeToPath & "\History_Outgoing"
            Case ("telecom/mch.vcf")
                writeToPath = writeToPath & "\History_Missed"
            Case ("telecom/cch.vcf")
                writeToPath = writeToPath & "\History_Combined"

            Case Else


        End Select

        Try
            If IO.Directory.Exists(writeToPath) = False Then
                IO.Directory.CreateDirectory(writeToPath)
            End If
        Catch ex As Exception

        End Try


        Dim writeToFile As String = writeToPath & "\CardList.XML"
        Try
            If IO.File.Exists(writeToFile) = False Then
                IO.File.Delete(writeToFile)
            End If
        Catch ex As Exception

        End Try


        tvwCardList.Nodes.Clear()


        Dim phoneDvcName As String = cboPhoneList.Text

        Dim connTorF As Boolean = BlueSoleil_ConnectService_ByName("PBAP", phoneDvcName, pbapHandleDvc, pbapHandleConn, pbapHandleSvc)

        Dim pullListTorF As Boolean = False

        Dim cardHandles(0 To 0) As String, cardNames(0 To 0) As String, cardCount As Integer = 0
        Dim i As Integer
        Dim pullCardTorF As Boolean = False
        Dim cardFN As String = ""

        If connTorF = True Then

            pullListTorF = BlueSoleil_PBAP_PullCardList_ByPath(pbapHandleConn, writeToFile, cboPhoneBooks.Text)


            'read XML
            If pullListTorF = True Then
                cardCount = BlueSoleil_PBAP_XML_GetCardListInfo(writeToFile, cardHandles, cardNames)
                For i = 0 To cardCount - 1
                    cardFN = writeToPath & "\" & cardHandles(i)

                    'remove old VCF file.
                    Try
                        If IO.File.Exists(cardFN) = True Then
                            IO.File.Delete(cardFN)
                        End If
                    Catch ex As Exception

                    End Try

                    pullCardTorF = BlueSoleil_PBAP_PullCard_ByPath(pbapHandleConn, cardHandles(i), cardFN, cboPhoneBooks.Text)

                    If pullCardTorF = True Then

                        Dim contactName As String = "", contactNumbers(0 To 0) As String, contactNumberLabels(0 To 0) As String, contactNumberCount As Integer, contactImage As Bitmap = Nothing, contactOrganization As String = ""
                        Dim contactEMail As String = "", contactLastCallDateTime As DateTime = Nothing, contactLastCallType As String = "", contactAddresses(0 To 0) As String, contactAddressLabels(0 To 0) As String, contactAddressCount As Integer, contactBirthday As String = "", contactGeoPos As String = "", contactNotes As String = ""
                        Dim tempItem As TreeNode = Nothing
                        Dim subItem As TreeNode = Nothing
                        VCard_GetContactInfo(cardFN, 0, contactName, contactEMail, contactNumbers, contactNumberLabels, contactNumberCount, contactImage, contactLastCallDateTime, contactLastCallType, contactAddresses, contactAddressLabels, contactAddressCount, contactBirthday, contactGeoPos, contactNotes, contactOrganization)

                        tempItem = tvwCardList.Nodes.Add(contactName)

                        If IsNothing(contactImage) = False Then
                            'contactImage.Dispose()
                            tempItem.Tag = contactImage
                        End If

                        tempItem.Tag = contactImage

                        For j = 0 To contactNumberCount - 1
                            If contactNumberLabels(j) <> "" Then
                                subItem = tempItem.Nodes.Add(contactNumberLabels(j) & " Phone: " & contactNumbers(j))
                            Else
                                subItem = tempItem.Nodes.Add("Phone: " & contactNumbers(j))
                            End If

                        Next j

                        For j = 0 To contactAddressCount - 1
                            If contactAddressLabels(j) <> "" Then
                                subItem = tempItem.Nodes.Add(contactAddressLabels(j) & " Addr: " & Replace(contactAddresses(j), vbNewLine, " \ "))
                            Else
                                subItem = tempItem.Nodes.Add("Addr: " & Replace(contactAddresses(j), vbNewLine, " \ "))
                            End If

                        Next j

                        If contactEMail <> "" Then
                            subItem = tempItem.Nodes.Add("EMail: " & contactEMail)
                        End If

                        If contactBirthday <> "" Then
                            subItem = tempItem.Nodes.Add("Birthday: " & contactBirthday)
                        End If

                        If IsNothing(contactLastCallDateTime) = False AndAlso contactLastCallDateTime.Year <> 1 Then
                            subItem = tempItem.Nodes.Add("Last Call: " & contactLastCallDateTime.ToShortDateString)
                        End If

                        If contactGeoPos <> "" Then
                            subItem = tempItem.Nodes.Add("GeoPos: " & contactGeoPos)
                        End If

                        If contactNotes <> "" Then
                            subItem = tempItem.Nodes.Add("Notes: " & contactNotes)
                        End If



                    End If

                Next i


            End If


            BlueSoleil_DisconnectServiceConn(pbapHandleConn)

        End If


        tvwCardList.ExpandAll()

        MsgBox("Done.  Connect = " & (pbapHandleConn <> 0) & "  Pull_List = " & pullListTorF & "  Pull_Card = " & pullCardTorF & ".  # of contacts = " & cardCount)



        pbapHandleConn = 0


    End Sub

    Private Sub tvwCardList_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvwCardList.AfterSelect

        Dim currNode As TreeNode = e.Node
        Dim parNode As TreeNode = currNode.Parent

        'get top-most node parent.
        Do
            If IsNothing(parNode) = False Then
                currNode = parNode
                parNode = currNode.Parent
            Else
                Exit Do
            End If
        Loop


        ' pbxPhonebookPic.BackgroundImage.Dispose()
        ' pbxPhonebookPic.BackgroundImage = Nothing

        pbxPhonebookPic.Visible = False
        If IsNothing(currNode.Tag) = False Then
            If TypeOf (currNode.Tag) Is Bitmap Then
                pbxPhonebookPic.BackgroundImageLayout = ImageLayout.Center
                pbxPhonebookPic.BackgroundImage = CType(currNode.Tag, Bitmap)
                pbxPhonebookPic.Refresh()
                pbxPhonebookPic.Visible = True
            End If
        End If


    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub lblHandsFreeStatus_Click(sender As Object, e As EventArgs) Handles lblHandsFreeStatus.Click

    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click

        Dim hfpBool As Boolean = BlueSoleil_HFP_HangUp(hfpHandleConnHFAG)

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles btnRefreshAllDevices.Click


        BlueSoleil_GetAllDevices_NamesAndHandles(dvcNameArray, dvcHandleArray, dvcArrayCount)

        If dvcArrayCount = 0 Then
            BlueSoleil_GetInquiredDevices_NamesAndHandles(dvcNameArray, dvcHandleArray, dvcArrayCount)
        End If

        cboPhoneList.Items.Clear()

        Dim i As Integer
        For i = 0 To dvcArrayCount - 1
            cboPhoneList.Items.Add(dvcNameArray(i))
        Next i
        If dvcArrayCount > 0 Then
            cboPhoneList.SelectedIndex = 0
        End If

        MsgBox("Done")

    End Sub

    Private Sub TabPage6_Click(sender As Object, e As EventArgs) Handles TabPage6.Click

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        avrcpNeedsSupportedEvents = True

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub
End Class
