'modBMSG - Written by Jesse Yeager.   www.CompulsiveCode.com
'
'This is a very basic wrapper for BMSG files used by BlueSoleil to send and receive messages from the MAP profile.
'

Option Explicit On
Option Strict On

Module modBMSG

    Private Function BMSG_HexToByte(ByVal inpHex As String) As Byte

        Return CByte(Val("&H" & inpHex))

    End Function

    Private Function BMSG_CleanUnicode(ByVal inpStr As String) As String
        'subroutine to support escaped UTF8 characters, and UTF8 overall.

        If inpStr = "" Then Return ""

        Dim inpLen As Integer = Len(inpStr)

        Dim outBytes(0 To 0) As Byte
        Dim outPos As Integer = 0


        Dim i As Integer
        For i = 1 To inpLen

            ReDim Preserve outBytes(0 To outPos)

            If Mid(inpStr, i, 1) = "=" And i < (inpLen - 1) Then
                outBytes(outPos) = BMSG_HexToByte(Mid(inpStr, i + 1, 2))

                i = i + 2   'skip two chars.  the NEXT statement will increment for the third char.
            Else
                outBytes(outPos) = CByte(Asc(Mid(inpStr, i, 1)))
            End If
            outPos = outPos + 1

        Next i

        Dim retStr As String = System.Text.Encoding.UTF8.GetString(outBytes)

        Return retStr

    End Function




    Private Sub BMSG_SeparateNameParts(ByVal inpFullName As String, ByRef retLastName As String, ByRef retFirstName As String, retAdditionalName As String, ByRef retNamePrefix As String, ByRef retNameSuffix As String)

        'this subroutine takes the full name, "Mr. Robert Downey, Jr." for example, and parses out the first, last, additional, prefix, and suffix for use in the N: entry of the VCF file.

        retLastName = ""
        retFirstName = ""
        retAdditionalName = ""
        retNamePrefix = ""
        retNameSuffix = ""

        inpFullName = Replace(inpFullName, ",", " , ")

        Dim nameWords(0 To 0) As String
        nameWords = Split(inpFullName, " ")

        Dim commaIdx As Integer = -1
        Dim i As Integer, j As Integer


        'get prefix and suffix if available, and remove them from the words.
        For i = 0 To nameWords.Length - 1
            Select Case UCase(nameWords(i))
                Case "MR", "MR.", "MS", "MS.", "MRS", "MRS."        'duke, earl, lord, sir, etc.
                    retNamePrefix = nameWords(i)
                    nameWords(i) = ""

                Case "JR", "JR.", "SR", "SR.", "II", "III", "IV"
                    retNameSuffix = nameWords(i)
                    nameWords(i) = ""

                Case ","
                    commaIdx = i

            End Select
        Next i

        'if there was a comma, remove comma and anything after it.
        If commaIdx <> -1 Then
            For i = commaIdx To nameWords.Length - 1
                nameWords(i) = ""
            Next i
        End If

        Dim newWordCount As Integer = 0
        For i = 0 To nameWords.Length - 1

            If nameWords(i) = "" Then   'if blank, shift up remaining words.
                For j = i To nameWords.Length - 2
                    nameWords(j) = nameWords(j + 1)
                Next j
                nameWords(nameWords.Length - 1) = ""
            Else
                newWordCount = newWordCount + 1
            End If

        Next i



        Select Case newWordCount
            Case 0
                'problem
                ''

            Case 1
                retFirstName = nameWords(0)

            Case 2
                retFirstName = nameWords(0)
                retLastName = nameWords(1)

            Case 3
                retFirstName = nameWords(0)
                retAdditionalName = nameWords(1)
                retLastName = nameWords(2)

            Case 4  'guessing here a first name like mary ann 
                retFirstName = nameWords(0) & " " & nameWords(1)
                retAdditionalName = nameWords(2)
                retLastName = nameWords(3)

            Case Else
                retFirstName = inpFullName

        End Select


    End Sub


    Public Sub BMSG_GetMessageInfo(ByVal BMsgFileName As String, ByRef retMsgText As String, ByRef retMsgFromName As String, ByRef retMsgFromPhoneNo As String, ByRef retMsgType As String, ByRef retMsgStatus As String, ByRef retMsgRemoteFolder As String, ByRef retAttachBytes() As Byte, ByRef retAttachType As String)

        Dim hFile As IntPtr = FileAPI_OpenFile(BMsgFileName, False)

        Dim inpFLen As Long = FileAPI_GetFileSize(hFile)

        FileAPI_SetFileOffset(hFile, 0)

        Dim tempLine As String = ""
        Dim tempLineType As String = "", tempSubType As String = "", tempValue As String = ""

        retMsgText = ""
        retMsgFromName = ""
        retMsgFromPhoneNo = ""
        retMsgType = ""
        retMsgStatus = ""
        retMsgRemoteFolder = ""
        ReDim retAttachBytes(0 To 0)
        retAttachType = ""

        Dim msgLength As Integer = 0

        Do
            If FileAPI_IsEOF(hFile) <> False Then Exit Do

            tempLine = ""
            FileAPI_ReadLineFromBinaryFile(hFile, -1, inpFLen, vbLf, tempLine)
            tempLine = Replace(tempLine, vbCr, "")

            BMSG_GetLineInfo(tempLine, tempLineType, tempSubType, tempValue)


            Select Case tempLineType

                Case "END"
                    If tempValue = "BMSG" Then Exit Do

                Case "FN"   'full name
                    retMsgFromName = tempValue
                    retMsgFromName = BMSG_CleanUnicode(retMsgFromName)
                    retMsgFromName = retMsgFromName

                Case "N"    'name
                    If retMsgFromName = "" Then
                        retMsgFromName = tempValue
                        retMsgFromName = BMSG_CleanUnicode(retMsgFromName)
                        retMsgFromName = retMsgFromName
                    End If

                Case "EMAIL"    'email
                    If retMsgFromName = "" Then
                        retMsgFromName = tempValue
                        retMsgFromName = BMSG_CleanUnicode(retMsgFromName)
                        retMsgFromName = retMsgFromName
                    End If

                Case "TEL"      'phone number
                    retMsgFromPhoneNo = tempValue

                Case "STATUS"
                    retMsgStatus = tempValue

                Case "TYPE"
                    retMsgType = tempValue

                Case "FOLDER"
                    retMsgRemoteFolder = tempValue

                Case "LENGTH"
                    msgLength = CInt(Val(tempValue))

                Case "BEGIN"
                    If tempValue = "MSG" Then

                        Dim lenToGet As Integer = msgLength - Len(tempLine) - 2
                        FileAPI_GetString(hFile, -1, lenToGet, retMsgText)
                        Dim p1 As Integer = InStrRev(retMsgText, "END:M")
                        If p1 > 0 Then
                            retMsgText = Strings.Left(retMsgText, p1 - 1)
                        End If



                        'Dim numBytesRead As Integer = 0
                        'Do
                        '    If FileAPI_IsEOF(hFile) = True Then Exit Do
                        '    tempLine = ""
                        '    FileAPI_ReadLineFromBinaryFile(hFile, -1, inpFLen, vbLf, tempLine)
                        '    tempLine = Replace(tempLine, vbCr, "")
                        '    numBytesRead = numBytesRead
                        'Loop


                        'remove trailing CrLf
                        If Strings.Right(retMsgText, 2) = vbCrLf Then
                            retMsgText = Strings.Left(retMsgText, Len(retMsgText) - 2)
                        End If

                        retMsgText = BMSG_CleanUnicode(retMsgText)

                        tempValue = retMsgText

                        BMSG_ParseContent(tempValue, retMsgText, retAttachBytes, retAttachType)



                    End If

            End Select

        Loop


        FileAPI_CloseFile(hFile)

    End Sub

    Private Function BMSG_ParseContent_GetData(ByVal dataStartingWithContent As String) As String

        'read past headers to actual data.
        Dim dataLines(0 To 0) As String
        dataLines = Split(dataStartingWithContent, vbLf)

        Dim i As Integer
        Dim firstBlankIdx As Integer = -1
        Dim secondBlankIdx As Integer = -1
        For i = 0 To dataLines.Length - 1
            If Strings.Right(dataLines(i), 1) = vbCr Then
                dataLines(i) = Strings.Left(dataLines(i), Len(dataLines(i)) - 1)
            End If

            If dataLines(i) = "" Then
                If firstBlankIdx = -1 Then
                    firstBlankIdx = i
                Else
                    If secondBlankIdx = -1 Then
                        secondBlankIdx = i
                    End If
                End If
            End If
        Next i

        If secondBlankIdx > 3 Then
            If Strings.Left(dataLines(secondBlankIdx - 1), 6) = "------" Then

                If Strings.Right(dataLines(secondBlankIdx - 1), 2) = "--" Then
                    secondBlankIdx = secondBlankIdx - 1
                End If

            End If

        End If

        Dim numLinesToKeep As Integer = secondBlankIdx - firstBlankIdx - 1

        Dim linesToKeep(0 To numLinesToKeep - 1) As String
        Array.Copy(dataLines, firstBlankIdx + 1, linesToKeep, 0, numLinesToKeep)

        Dim retStr As String = ""
        For i = 0 To numLinesToKeep - 1
            retStr = retStr & linesToKeep(i) & vbLf
        Next

        Return retStr

    End Function

    Private Sub BMSG_ParseContent(ByVal inpMsgString As String, ByRef retMsgText As String, ByRef retMsgAttachmentBytes() As Byte, ByRef retMsgAttachmentType As String)

        ReDim retMsgAttachmentBytes(0 To 0)
        retMsgText = ""
        retMsgAttachmentType = ""

        Dim p1 As Integer = InStr(1, inpMsgString, "Content-Type:", CompareMethod.Text)
        If p1 = 0 Then
            retMsgText = inpMsgString
            Exit Sub
        End If

        p1 = InStr(1, inpMsgString, "Date:", CompareMethod.Text)
        If p1 = 0 Then
            retMsgText = inpMsgString
            Exit Sub
        End If

        Dim isBase64 As Boolean = False

        Dim contentTypePos As Integer = 0
        Dim contentEncodingPos As Integer = 0

        Dim tempContentText As String = ""
        Do
            contentTypePos = InStr(contentTypePos + 1, inpMsgString, "Content-Type:", CompareMethod.Text)
            contentEncodingPos = InStr(contentTypePos + 1, inpMsgString, "Content-Transfer-Encoding:", CompareMethod.Text)
            If contentTypePos < 1 Then Exit Do

            If InStr(Strings.Mid(inpMsgString, contentEncodingPos, 50), "Content-Transfer-Encoding: Base64", CompareMethod.Text) = 1 Then
                isBase64 = True
            End If


            If InStr(Strings.Mid(inpMsgString, contentTypePos, 50), "application/", CompareMethod.Text) = 0 Then
                'not useless garbage.
                tempContentText = Strings.Mid(inpMsgString, contentTypePos)


                If InStr(tempContentText, "Content-Type: text/plain", CompareMethod.Text) = 1 Then
                    'found msg text.
                    retMsgText = BMSG_ParseContent_GetData(tempContentText)

                Else
                    'attachment?
                    retMsgAttachmentType = Strings.Mid(tempContentText, 14)
                    p1 = InStr(retMsgAttachmentType, vbLf)
                    If p1 > 0 Then retMsgAttachmentType = Strings.Left(retMsgAttachmentType, p1 - 1)
                    retMsgAttachmentType = Replace(retMsgAttachmentType, vbCr, "")


                    Dim attachmentBase64 As String = BMSG_ParseContent_GetData(tempContentText)
                    attachmentBase64 = Replace(attachmentBase64, vbLf, "")
                    attachmentBase64 = Replace(attachmentBase64, vbCr, "")
                    If Len(attachmentBase64) Mod 4 <> 0 Then
                        attachmentBase64 = attachmentBase64 & "="
                    End If
                    If Len(attachmentBase64) Mod 4 <> 0 Then
                        attachmentBase64 = attachmentBase64 & "="
                    End If
                    If Len(attachmentBase64) Mod 4 <> 0 Then
                        attachmentBase64 = attachmentBase64 & "="
                    End If

                    Try
                        retMsgAttachmentBytes = Convert.FromBase64String(attachmentBase64)
                    Catch ex As Exception
                        ReDim retMsgAttachmentBytes(0 To 0)
                    End Try



                End If




            End If
        Loop




    End Sub


    Private Sub BMSG_GetLineInfo(ByVal inpLine As String, ByRef retLineType As String, ByRef retSubType As String, ByRef retValue As String)


        Dim p1 As Integer

        p1 = InStr(1, inpLine, ":")
        If p1 < 1 Then
            retLineType = ""
            retSubType = ""
            Return
        End If

        Dim leftPart As String = Left(inpLine, p1 - 1)
        Dim rightPart As String = Mid(inpLine, p1 + 1)

        retValue = rightPart
        Do
            If Left(retValue, 1) = ";" Then
                retValue = Mid(retValue, 2)
            Else
                Exit Do
            End If
        Loop

        p1 = InStr(1, leftPart, ";")
        If p1 = 0 Then
            retLineType = leftPart
            retSubType = ""
            Return
        End If

        retLineType = Left(leftPart, p1 - 1)
        retSubType = Mid(leftPart, p1 + 1)
        retSubType = Replace(retSubType, "TYPE=", "", 1, -1, CompareMethod.Text)

        If retLineType = "PHOTO" Then

        End If

    End Sub



    Public Sub BMSG_WriteMessageFile_Text(ByVal BMsgFileName As String, ByVal recipPhoneNumber As String, ByVal textToSend As String)



        Dim hFile As IntPtr, fLen As Long

        hFile = FileAPI_OpenFile(BMsgFileName, True)
        fLen = FileAPI_GetFileSize(hFile)


        'write out the easy stuff.

        Dim lineStr As String = ""
        Dim lineBytes(0 To 0) As Byte

        lineStr = "BEGIN:BMSG" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "VERSION:1.0" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "STATUS:UNREAD" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "TYPE:SMS_GSM" & vbCrLf                               'change to MMS if writing an MMS file, but that is a bit more complicated.
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        lineStr = "BEGIN:BENV" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        lineStr = "BEGIN:VCARD" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "VERSION:3.0" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "TEL:" & recipPhoneNumber & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "END:VCARD" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        lineStr = "BEGIN:BBODY" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "CHARSET:UTF-8" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        Dim tempMsg As String = "BEGIN:MSG" & vbCrLf & textToSend & vbCrLf & "END:MSG" & vbCrLf
        Dim tempMsgLen As Integer = Len(tempMsg)
        lineStr = "LENGTH:" & tempMsgLen & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = tempMsg
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        lineStr = "END:BBODY" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "END:BENV" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "END:BMSG" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        FileAPI_CloseFile(hFile)

    End Sub



    Private Function BMSG_WriteMessageFile_MMS_Image_BuildMIMEcontent(ByVal textToSend As String, ByVal fileToAttach As String, ByVal fromName As String, ByVal fromAddress As String, ByRef toName As String, ByVal toAddress As String) As String

        Dim retStr As String = ""

        Dim strBoundary As String = "--=_784c66f8-c601-4feb-8399-4dc2946386e7"

        'build MIME headers ourselves..  Screw it.  Why not?  We've come this far.





        Return retStr


    End Function


    Public Sub BMSG_WriteMessageFile_MMS_Image(ByVal BMsgFileName As String, ByVal recipPhoneNumber As String, ByVal textToSend As String, ByVal fileToAttach As String)


        Dim hFile As IntPtr, fLen As Long

        hFile = FileAPI_OpenFile(BMsgFileName, True)
        fLen = FileAPI_GetFileSize(hFile)


        'write out the easy stuff.

        Dim lineStr As String = ""
        Dim lineBytes(0 To 0) As Byte

        lineStr = "BEGIN:BMSG" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "VERSION:1.0" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "STATUS:UNREAD" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "TYPE:MMS_GSM" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        lineStr = "BEGIN:BENV" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        lineStr = "BEGIN:VCARD" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "VERSION:3.0" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "TEL:" & recipPhoneNumber & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "END:VCARD" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        lineStr = "BEGIN:BBODY" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "CHARSET:UTF-8" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        Dim tempMsg As String = "BEGIN:MSG" & vbCrLf & textToSend & vbCrLf & "END:MSG" & vbCrLf
        Dim tempMsgLen As Integer = Len(tempMsg)
        lineStr = "LENGTH:" & tempMsgLen & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = tempMsg
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)


        lineStr = "END:BBODY" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "END:BENV" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        lineStr = "END:BMSG" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        FileAPI_CloseFile(hFile)

    End Sub

End Module
