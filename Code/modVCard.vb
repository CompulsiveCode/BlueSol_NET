'modVCard - Written by Jesse Yeager.  www.CompulsiveCode.com
'
'This is a very basic wrapper for VCard (.VCF) files used by Blue Soleil to send and receive contacts from the PBAP profile.
'

Option Explicit On
Option Strict On

Module modVCard

    Private Function VCard_HexToByte(ByVal inpHex As String) As Byte

        Return CByte(Val("&H" & inpHex))

    End Function

    Private Function VCard_CleanUnicode(ByVal inpStr As String) As String
        'subroutine to support escaped UTF8 characters, and UTF8 overall.

        If inpStr = "" Then Return ""

        Dim inpLen As Integer = Len(inpStr)

        Dim outBytes(0 To 0) As Byte
        Dim outPos As Integer = 0


        Dim i As Integer
        For i = 1 To inpLen

            ReDim Preserve outBytes(0 To outPos)

            If Mid(inpStr, i, 1) = "=" And i < (inpLen - 1) Then
                outBytes(outPos) = VCard_HexToByte(Mid(inpStr, i + 1, 2))

                i = i + 2   'skip two chars.  the NEXT statement will increment for the third char.
            Else
                outBytes(outPos) = CByte(Asc(Mid(inpStr, i, 1)))
            End If
            outPos = outPos + 1

        Next i

        Dim retStr As String = System.Text.Encoding.UTF8.GetString(outBytes)

        retStr = Replace(retStr, "\,", ",")     'do this to clean up escaped commas.

        Return retStr

    End Function

    Private Sub VCard_SeparateAddressParts(ByVal inpAddress As String, ByRef retPObox As String, ByRef retExtAddr As String, ByRef retStreetAddr As String, ByRef retLocalityCity As String, ByRef retRegionState As String, ByRef retPostalCode As String, ByRef retCountry As String, ByRef retFormattedAddress As String)

        retPObox = ""
        retExtAddr = ""
        retStreetAddr = ""
        retLocalityCity = ""
        retRegionState = ""
        retPostalCode = ""
        retCountry = ""

        retFormattedAddress = ""

        Dim strArray(0 To 0) As String, strCount As Integer = 0
        strArray = Split(inpAddress, ";")
        strCount = strArray.Length

        Dim i As Integer
        For i = 0 To strCount - 1
            strArray(i) = Trim(strArray(i))

            If strArray(i) <> "" Then retFormattedAddress = retFormattedAddress & strArray(i) & vbNewLine
        Next i

        If Strings.Right(retFormattedAddress, 2) = vbNewLine Then
            retFormattedAddress = Strings.Left(retFormattedAddress, Len(retFormattedAddress) - 2)
        End If
        retFormattedAddress = VCard_CleanUnicode(retFormattedAddress)

        If strCount > 0 Then retPObox = strArray(0)
        If strCount > 1 Then retExtAddr = strArray(1)
        If strCount > 2 Then retStreetAddr = strArray(2)
        If strCount > 3 Then retLocalityCity = strArray(3)
        If strCount > 4 Then retRegionState = strArray(4)
        If strCount > 5 Then retPostalCode = strArray(5)
        If strCount > 6 Then retCountry = strArray(6)



    End Sub


    Private Sub VCard_SeparateNameParts(ByVal inpFullName As String, ByRef retLastName As String, ByRef retFirstName As String, retAdditionalName As String, ByRef retNamePrefix As String, ByRef retNameSuffix As String)

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

        End Select


    End Sub

    Public Function VCard_WriteContactInfo_V3(ByVal vCardFileName As String, ByVal inpCardFullName As String, ByRef inpPhoneNumbers() As String, ByRef inpPhoneLabels() As String, ByVal inpPhoneCount As Integer, ByRef inpAddresses() As String, ByRef inpAdrsLabels() As String, ByVal inpAdrsCount As Integer, Optional ByVal inpImage As Bitmap = Nothing, Optional ByVal inpCardEMailAddress As String = "", Optional ByVal inpOrganization As String = "", Optional ByVal inpNotes As String = "") As Boolean

        'inpPhoneLabels should be one of:  WORK HOME CELL


        'if vCardFile does not exist, create.
        'otherwise, append to existing file.

        Dim hFile As IntPtr, fLen As Long
        If IO.File.Exists(vCardFileName) = False Then
            hFile = FileAPI_OpenFile(vCardFileName, True)
            fLen = FileAPI_GetFileSize(hFile)
        Else
            hFile = FileAPI_OpenFile(vCardFileName, False)
            fLen = FileAPI_GetFileSize(hFile)
            FileAPI_SetFileOffset(hFile, fLen)

            'could check last two bytes of file to ensure it has EOL (CrLf).
            'or could always add CrLf just to be safe.  I don't think an extra blank line would break anything in a VCF file.

            'DO NOT USE THIS WITHOUT CONSIDERING THE POSSIBILITY THAT THE CONTACT ALREADY EXISTS IN THE EXISTING VCF FILE.
        End If

        If hFile = IntPtr.Zero Then Return False


        'write out the easy stuff.

        Dim lineStr As String = ""
        Dim lineBytes(0 To 0) As Byte

        'begin VCard entry.
        lineStr = "BEGIN:VCARD" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        'specify version.
        lineStr = "VERSION:3.0" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)




        'add name.  format:  last;first;additional;prefix;suffix
        Dim nLast As String = "", nFirst As String = "", nAdditional As String = "", nPrefix As String = "", nSuffix As String = ""
        VCard_SeparateNameParts(inpCardFullName, nLast, nFirst, nAdditional, nPrefix, nSuffix)
        Dim inpCardName As String = nLast & ";" & nFirst & ";" & nAdditional & ";" & nPrefix & ";" & nSuffix
        'lineStr = "N:" & inpCardName & vbCrLf
        lineStr = "N;CHARSET=UTF-8:" & inpCardName & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        'add full name.
        'lineStr = "FN:" & inpCardFullName & vbCrLf
        lineStr = "FN;CHARSET=UTF-8:" & inpCardFullName & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        'add all phone numbers (with labels if present).
        Dim i As Integer
        For i = 0 To inpPhoneCount - 1

            If inpPhoneLabels(i) <> "" Then
                inpPhoneLabels(i) = UCase(inpPhoneLabels(i))
                lineStr = "TEL;TYPE=" & inpPhoneLabels(i) & ":" & inpPhoneNumbers(i) & vbCrLf
            Else
                lineStr = "TEL:" & inpPhoneNumbers(i) & vbCrLf
            End If

            lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
            FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)
        Next i

        'add all addresses (with labels if present).
        For i = 0 To inpAdrsCount - 1
            If inpAddresses(i) <> "" Then
                inpAddresses(i) = Replace(inpAddresses(i), vbNewLine, "\")
                If inpAdrsLabels(i) <> "" Then
                    inpAdrsLabels(i) = UCase(inpAdrsLabels(i))
                    lineStr = "ADR;TYPE=" & inpPhoneLabels(i) & ":" & inpAddresses(i) & vbCrLf
                Else
                    lineStr = "ADR:" & inpAddresses(i) & vbCrLf
                End If

                lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
                FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)
            End If

        Next i



        If IsNothing(inpImage) = False Then
            If inpImage.Width > 0 And inpImage.Height > 0 Then
                'resize to 96x96 ?

                Dim newImage As New Bitmap(inpImage, 96, 96)
                Dim newImageGfx As Graphics
                newImageGfx = Graphics.FromImage(newImage)
                newImageGfx.DrawImage(inpImage, New Rectangle(0, 0, 96, 96))


                'save image to stream.
                Dim imgStream As New IO.MemoryStream()
                newImage.Save(imgStream, Imaging.ImageFormat.Jpeg)      'could also do PNG, but must change encoding type below!

                'get stream bytes.
                Dim imgBytes(0 To 0) As Byte
                ReDim imgBytes(0 To CInt(imgStream.Length - 1))
                imgStream.Position = 0
                imgStream.Read(imgBytes, 0, CInt(imgStream.Length))

                'close and dispose objects.
                imgStream.Close()
                imgStream.Dispose()
                newImageGfx.Dispose()
                newImage.Dispose()

                'encode to base64.
                Dim imgBase64str As String = ""
                imgBase64str = Convert.ToBase64String(imgBytes)

                'add the line header stuff.
                lineStr = "PHOTO;ENCODING=B;TYPE=JPEG:" & imgBase64str

                'after 253 chars, insert CRLF and space.

                i = 74
                Dim tempLeftPart As String = "", tempRightPart As String = ""
                Do
                    If i > Len(lineStr) Then Exit Do

                    tempLeftPart = Strings.Left(lineStr, i)
                    tempRightPart = Strings.Mid(lineStr, i + 1)

                    lineStr = tempLeftPart & vbCrLf & " " & tempRightPart

                    i = i + 76

                Loop

                lineStr = lineStr & vbCrLf & vbCrLf

                lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
                FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

            End If

        End If


        'add email address if present.
        If inpCardEMailAddress <> "" Then
            lineStr = "EMAIL:" & inpCardEMailAddress & vbCrLf
            lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
            FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)
        End If

        'can do ORG: for company name.
        If inpOrganization <> "" Then
            'lineStr = "ORG:" & inpCompanyName & vbCrLf
            lineStr = "ORG;CHARSET=UTF-8:" & inpOrganization & vbCrLf
            lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
            FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)
        End If

        'can do note
        If inpNotes <> "" Then
            'lineStr = "NOTE:" & inpCompanyName & vbCrLf
            lineStr = "NOTE;CHARSET=UTF-8:" & inpNotes & vbCrLf
            lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
            FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)
        End If


        'end the vcard data.
        lineStr = "END:VCARD" & vbCrLf
        lineBytes = System.Text.Encoding.UTF8.GetBytes(lineStr)
        FileAPI_PutBytes(hFile, -1, lineBytes.Length, lineBytes)

        'close the file.
        FileAPI_CloseFile(hFile)

        Return True

    End Function


    Public Sub VCard_GetContactInfo(ByVal vCardFileName As String, ByVal vCardOffset As Long, ByRef retCardName As String, ByRef retCardEmail As String, ByRef retPhoneNumbers() As String, ByRef retPhoneLabels() As String, ByRef retPhoneCount As Integer, ByRef retImage As Bitmap, ByRef retLastCallDateTime As DateTime, ByRef retlastCallType As String, ByRef retAddresses() As String, ByRef retAddressLabels() As String, ByRef retAddressCount As Integer, ByRef retBirthDay As String, ByRef retGeoPosition As String, ByRef retNote As String, ByRef retOrganization As String)

        Dim hFile As IntPtr = FileAPI_OpenFile(vCardFileName, False)

        Dim inpFLen As Long = FileAPI_GetFileSize(hFile)

        FileAPI_SetFileOffset(hFile, vCardOffset)

        Dim tempLine As String = ""
        Dim tempLineType As String = "", tempSubType As String = "", tempValue As String = ""

        retCardName = ""
        retCardEmail = ""
        retPhoneCount = 0
        retAddressCount = 0
        retBirthDay = ""
        retGeoPosition = ""
        retNote = ""
        retLastCallDateTime = Nothing
        retlastCallType = ""       'RECEIVED, MISSED, or DIALED
        retOrganization = ""

        ReDim retAddresses(0 To 0)
        ReDim retAddressLabels(0 To 0)          'HOME, CELL, WORK, PREF
        ReDim retPhoneNumbers(0 To 0)
        ReDim retPhoneLabels(0 To 0)            'HOME, CELL, WORK, PREF

        Dim imgBytes(0 To 0) As Byte



        Do
            If FileAPI_IsEOF(hFile) <> False Then Exit Do

            tempLine = ""
            FileAPI_ReadLineFromBinaryFile(hFile, -1, inpFLen, vbLf, tempLine)
            tempLine = Replace(tempLine, vbCr, "")

            VCard_GetLineInfo(tempLine, tempLineType, tempSubType, tempValue)

            If tempLineType = "PHOTO" Then

                'check encoding.  if not base64, then it's probably a path or URL to a file.
                '
                ''  need to do this.  maybe not, since error-handlers will handle it.


                'process base64 data.  get multi-line data.
                Dim imgData As String = tempValue

                Do
                    tempLine = ""
                    FileAPI_ReadLineFromBinaryFile(hFile, -1, inpFLen, vbLf, tempLine)
                    tempLine = Replace(tempLine, vbCr, "")

                    If Strings.Left(tempLine, 1) <> " " Then
                        'no longer in image data.  take the image data and make an image object...


                        'base64 data length should be a multiple of 4.  let's pad it if necessary so .Net doesn't cry about it.
                        If Len(imgData) Mod 4 <> 0 Then
                            imgData = imgData & "="
                        End If
                        If Len(imgData) Mod 4 <> 0 Then
                            imgData = imgData & "="
                        End If
                        If Len(imgData) Mod 4 <> 0 Then
                            imgData = imgData & "="
                        End If

                        'perform base64 to byte-array conversion.
                        Dim retImageBytes(0 To 0) As Byte
                        Try
                            retImageBytes = Convert.FromBase64String(imgData)
                        Catch ex As Exception
                            ReDim retImageBytes(0 To 0)
                        End Try


                        'convert byte array to image.
                        Dim tempStream As New IO.MemoryStream(retImageBytes)
                        tempStream.Position = 0
                        Dim tempImage As Bitmap = Nothing
                        Try
                            tempImage = CType(Bitmap.FromStream(tempStream), Bitmap)
                            retImage = New Bitmap(tempImage)

                        Catch ex As Exception

                        End Try
                        tempImage = Nothing
                        tempStream.Dispose()

                        'get line info, since we're not in image data anymore.
                        VCard_GetLineInfo(tempLine, tempLineType, tempSubType, tempValue)
                        Exit Do
                    Else
                        imgData = imgData & Mid(tempLine, 2)
                    End If

                Loop
            End If


            Select Case tempLineType

                Case "END"
                    If tempValue = "VCARD" Then Exit Do

                Case "FN"   'full name
                    retCardName = tempValue
                    retCardName = VCard_CleanUnicode(retCardName)
                    retCardName = retCardName

                Case "N"    'name
                    If retCardName = "" Then
                        retCardName = tempValue
                        retCardName = VCard_CleanUnicode(retCardName)
                        retCardName = retCardName

                    End If



                Case "TEL"      'phone number
                    ReDim Preserve retPhoneNumbers(0 To retPhoneCount)
                    ReDim Preserve retPhoneLabels(0 To retPhoneCount)
                    retPhoneNumbers(retPhoneCount) = tempValue
                    retPhoneLabels(retPhoneCount) = tempSubType
                    retPhoneCount = retPhoneCount + 1


                Case "ADR" '
                    Dim tempFormattedAddress As String = ""
                    VCard_SeparateAddressParts(tempValue, "", "", "", "", "", "", "", tempFormattedAddress)

                    ReDim Preserve retAddresses(0 To retAddressCount)
                    ReDim Preserve retAddressLabels(0 To retAddressCount)
                    retAddresses(retAddressCount) = tempFormattedAddress
                    retAddressLabels(retAddressCount) = tempSubType
                    retAddressCount = retAddressCount + 1

                Case "EMAIL"
                    If retCardEmail = "" Then
                        retCardEmail = tempValue
                        retCardEmail = VCard_CleanUnicode(retCardEmail)
                        retCardEmail = retCardEmail
                    End If

                Case "X-IRMC-CALL-DATETIME"
                    retLastCallDateTime = VCard_ConvertStringToDateTime(tempValue)
                    retlastCallType = tempSubType

                Case "BDAY"
                    retBirthDay = tempValue
                    retBirthDay = VCard_CleanUnicode(retBirthDay)

                Case "GEO"
                    If retGeoPosition = "" Then
                        retGeoPosition = tempValue
                        retGeoPosition = VCard_CleanUnicode(retGeoPosition)
                        retGeoPosition = UCase(retGeoPosition).Replace("GEO:", "")
                    End If

                Case "NOTE"
                    If retNote = "" Then
                        retNote = tempValue
                        retNote = VCard_CleanUnicode(retNote)
                        retNote = retNote

                    End If


                Case "ORG"
                    If retOrganization = "" Then
                        retOrganization = tempValue
                        retOrganization = VCard_CleanUnicode(retOrganization)
                        retOrganization = retOrganization

                    End If





            End Select

        Loop


        FileAPI_CloseFile(hFile)

    End Sub



    Public Function VCard_GetContactOffsets(ByVal vCardFileName As String, ByRef vCardOffsets() As Long, ByRef vCardVersion As String) As Integer

        Dim vCardCount As Integer = 0
        ReDim vCardOffsets(0 To 0)
        vCardVersion = ""

        Dim hFile As IntPtr = FileAPI_OpenFile(vCardFileName, False)
        FileAPI_SetFileOffset(hFile, 0)

        Dim inpFLen As Long = FileAPI_GetFileSize(hFile)

        Dim lineOffset As Long = 0

        Dim tempLine As String = ""
        Dim tempLineType As String = "", tempSubType As String = "", tempValue As String = ""
        Do
            If FileAPI_IsEOF(hFile) <> False Then Exit Do

            tempLine = ""

            lineOffset = FileAPI_GetCurrentOffset(hFile)
            FileAPI_ReadLineFromBinaryFile(hFile, -1, inpFLen, vbLf, tempLine)
            tempLine = Replace(tempLine, vbCr, "")

            VCard_GetLineInfo(tempLine, tempLineType, tempSubType, tempValue)

            If tempLineType = "VERSION" Then
                If vCardVersion = "" Then
                    vCardVersion = tempValue
                End If
            End If

            If tempLineType = "BEGIN" And tempValue = "VCARD" Then
                ReDim Preserve vCardOffsets(0 To vCardCount)
                vCardOffsets(vCardCount) = lineOffset    ' FileAPI_GetCurrentOffset(hFile)

                vCardCount = vCardCount + 1
            End If

        Loop


        FileAPI_CloseFile(hFile)

        Return vCardCount

    End Function


    Private Sub VCard_GetLineInfo(ByVal inpLine As String, ByRef retLineType As String, ByRef retSubType As String, ByRef retValue As String)


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
        'Do
        '    If Left(retValue, 1) = ";" Then
        '        retValue = Mid(retValue, 2)
        '    Else
        '        Exit Do
        '    End If
        'Loop

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


    Private Function VCard_ConvertStringToDateTime(ByVal inpDateTimeStr As String) As DateTime

        Dim inpYear As Integer = CInt(Val(Mid(inpDateTimeStr, 1, 4)))
        Dim inpMonth As Integer = CInt(Val(Mid(inpDateTimeStr, 5, 2)))
        Dim inpDay As Integer = CInt(Val(Mid(inpDateTimeStr, 7, 2)))

        'char 9 should be "T"

        Dim inpHour As Integer = CInt(Val(Mid(inpDateTimeStr, 10, 2)))
        Dim inpMinute As Integer = CInt(Val(Mid(inpDateTimeStr, 12, 2)))
        Dim inpSecond As Integer = CInt(Val(Mid(inpDateTimeStr, 14, 2)))

        'if len greater than 15 and char 16 = "Z", time zone present?
        'Dim inpZoneOffset As Integer = CInt(Val(Mid(inpDateTimeStr, 17, 5)))

        Dim retDateTime As New DateTime(inpYear, inpMonth, inpDay, inpHour, inpMinute, inpSecond)

        Return retDateTime

    End Function


    Public Function VCard_ResortFile_ByName(ByVal inpFileName As String, ByVal sortedFileName As String) As Boolean

        'we are going to do this in a very basic way.

        'first, get the offset of each card in the file.
        Dim inpCardOffsets(0 To 0) As Long, inpCardCount As Integer = 0
        inpCardCount = VCard_GetContactOffsets(inpFileName, inpCardOffsets, "")

        If inpCardCount < 1 Then
            'if no cards, error out.
            Return False
        End If


        'get the size of the input file.  we will need it later.
        Dim inpFileinfo As New IO.FileInfo(inpFileName)
        Dim inpFileLen As Long = inpFileinfo.Length

        'build an array of sizes (bytes) for each vCard within the file.
        Dim i As Integer
        Dim inpCardSizes(0 To inpCardCount - 1) As Long
        For i = 0 To inpCardCount - 1

            'calculate the size of the vcard data (number of bytes until next vcard or end of file).
            If i < inpCardCount - 1 Then
                inpCardSizes(i) = inpCardOffsets(i + 1) - inpCardOffsets(i) - 1
            Else
                inpCardSizes(i) = inpFileLen - inpCardOffsets(i) - 1
            End If
        Next i



        'build a string array that starts with the contact name (for sorting purposes), and also contains the offset and size for every card.
        Dim strArray(0 To inpCardCount - 1) As String
        For i = 0 To inpCardCount - 1

            'we dont use most of these variables for anything.  just to call the VCard_GetContactInfo function to get the name.
            Dim retCardName As String = "", notused_PhoneNumbers(0 To 0) As String, notused_PhoneLabels(0 To 0) As String, notused_Addresses(0 To 0) As String, notused_AddressLabels(0 To 0) As String

            VCard_GetContactInfo(inpFileName, inpCardOffsets(i), retCardName, "", notused_PhoneNumbers, notused_PhoneLabels, 0, Nothing, Nothing, "", notused_Addresses, notused_AddressLabels, 0, "", "", "", "")

            retCardName = Replace(retCardName, "|", " ")        'just in case card name contains a pipe |

            'add delimited array item - Name | Offset | Size
            strArray(i) = retCardName & "|" & inpCardOffsets(i) & "|" & inpCardSizes(i)
        Next i

        'sort the string array.  Since the items begin with the card Name, that is the sort key.
        Array.Sort(strArray)


        'open input file and output file
        Dim inpStream As IO.FileStream = Nothing
        Dim outStream As IO.FileStream = Nothing
        Try
            inpStream = New IO.FileStream(inpFileName, IO.FileMode.Open)
            outStream = New IO.FileStream(sortedFileName, IO.FileMode.Create)

        Catch ex As Exception

            Return False

        End Try


        'use file operations to re-write the file in the sorted order.
        For i = 0 To inpCardCount - 1

            Dim strSplit(0 To 0) As String
            strSplit = Split(strArray(i), "|")
            'strSplit(0) is the name, strSplit(1) is the offset, strSplit(2) is the size

            inpStream.Position = CLng(Val(strSplit(1)))         'set stream position

            Dim cardBytes(0 To 0) As Byte
            ReDim cardBytes(0 To CInt(Val(strSplit(2))))

            inpStream.Read(cardBytes, 0, cardBytes.Length)      'read card data.
            outStream.Write(cardBytes, 0, cardBytes.Length)     'write card data.

        Next i

        'close streams and dispose objects.
        inpStream.Close()
        inpStream.Dispose()

        outStream.Flush()
        outStream.Close()
        outStream.Dispose()


        'yay we win.
        Return True


    End Function

End Module
