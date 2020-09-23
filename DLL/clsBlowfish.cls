VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBlowFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       Blowfish Encryption/Decryption
'
' Module:        clsBlowfish.cls
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
' *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING ***
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
'    You acknowledge that this software is subject to the export control
'    laws and regulations of the United States ("U.S.") and agree to abide
'    by those laws and regulations. Under U.S. law, this software may not
'    be downloaded or otherwise exported, reexported, or transferred to
'    restricted countries, restricted end-users, or for restricted
'    end-uses. The U.S. currently has embargo restrictions against Cuba,
'    Iran, Iraq, Libya, North Korea, Sudan, and Syria. The lists of
'    restricted end-users are maintained on the U.S. Commerce Department's
'    Denied Persons List, the Commerce Department's Entity List, the
'    Commerce Department's List of Unverified Persons, and the U.S.
'    Treasury Department's List of Specially Designated Nationals and
'    Blocked Persons. In addition, this software may not be downloaded or
'    otherwise exported, reexported, or transferred to an end-user engaged
'    in activities related to weapons of mass destruction.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
' Description:   Blowfish is a block cipher that was designed in 1993 by
'                Bruce Schneier as a fast, free alternative to existing
'                encryption algorithms. Like Blowfish it can be used as
'                a convenient and much faster substitute for DES. Since
'                then Blowfish has been extensively analyzed and no
'                significant weaknesses have been found. It is considered
'                to be a strong algorithm and has been implemented in over
'                130 commercial applications. Bruce Schneier was also one
'                of the designers of the Blowfish algorithm, which is one of
'                the five finalists selected by NIST in the Advanced Encryption
'                Standard (AES) competition. Blowfish was developed as a
'                successor to the Blowfish cipher, which does not meet the
'                design requirements of the AES.
'
'                Blowfish is a 64-bit block cipher, meaning that data is
'                encrypted and decrypted in 64-bit chunks. The key length
'                can vary from 32 to 448 bits. The algorithm uses 16 rounds,
'                or iterations of the main algorithm. It has been found that
'                the number of rounds is exponentially proportional to the
'                amount of time required to find a key using a brute-force
'                attack. So as the number of rounds increases, the security
'                of the algorithm increases exponentially.
'
'                Private Encryptor's implementation of Blowfish allows the
'                key to vary from 32 to 448 bits, as per the algorithm
'                specification.
'
'                The relative strength of the encryption algorithm is based on
'                key length.  Bruce Schneier, creator of the Blowfish encryption
'                algorithm, has calculated that according to what we know of
'                quantum mechanics today, that the entire energy output of the
'                sun is insufficient to break a 192-bit key.
'
'                The detailed description of the Blowfish algorithm is contained
'                in the original Blowfish paper, written by Bruce Schneier.
'
'                **********************************************************
' SPECIAL NOTE:  The encryption process has been enhanced.
'
'                - Encryption process can be performed multiple times.
'                  See property EncryptRounds().
'
'                - Password key length is expanded to meet and exceed AES
'                  standards.  Key lengths are 32 to 448 in 32 bit
'                  increments.  See property KeyLength().
'                **********************************************************
'
' Warning:       The weakest link in any encryption is the password. This
'                includes the security, length and components that make up
'                the password.  Password length and components are usually
'                dictated by organizational standards.  Security is the
'                responsibility of the user. Do not write down or discuss
'                your password with others.  Lock your screensaver when not
'                at your workstation.
'
' Note:          After encryption, data sizes will not match original sizes.
'                This is due to internal padding and the storing of information
'                required to decrypt the data later.
'
' References:    Original algorithm Written by Bruce Schneier
'                http://www.counterpane.com/blowfish.html
'
'                Tropical software (Algorithm description)
'                http://www.tropsoft.com/strongenc/blowfish.htm
'
'                Visual Basic translation by Fredrik Qvarfort
'                http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=12023&lngWId=1
'                12-Oct-2000
'
'                An Introduction to Using Keys in Cryptography
'                By DI Management Services Pty Ltd, Sydney Australia
'                http://www.di-mgt.com.au/cryptokeys.html
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-JUN-2004  Kenneth Ives  kenaso@tx.rr.com
'              Documented, modified, updated variable names
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added GetWord() and PutWord() routines.  Fixed a logic bug
'              in BlockDecrypt() routine.
' 05-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Renamed lNewValue() to CalcNewValue().  Modifed CalcNewValue()
'              routine for easier maintenance.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified Blowfish_File() and Blowfish_String() routines to
'              encrypt appended data information for additional security.
'              Moved ExpandData() routine to basCommon.bas module.
' 26-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added property KeyLength() to meet design standards. Set key
'              length to increments of 32 bits for 128 to 448.
'              Modified EvaluateKey() routine to use utilize the new key
'              lengths.
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed OffSet range in EncryptData() and DecryptData() routines.
'              Moved GetBlockSize() routine to common module.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
'              Correct data size capture for decryption in Blowfish_String().
' 21-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified EvaluateKey() routine to perform multiple rounds of
'              block encryption when mixing the values for the password key
'              [malngPrimaryKey()] and S-Box [malngSBox()] arrays.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to see progress display while encrypting or
'              decrypting data (EncryptData() and DecryptData() routines).
'              Corrected excess looping in EvaluateKey() routine.
'              Removed obsolete code.  Updated documentation.
' 10-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Removed MixBytes() routine.  I created this routine and now
'              feel that it is no longer needed.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made many updates as per Joe Sova's suggestions
'                1. Removed most of the erase array as they were not needed
'                2. Moved call to display progress to outside processing loop
'                   in EncryptData() and DecryptData() routines.
'                3. Moved Decryption size testing to earlier position within
'                   within Blowfish_File() routine.
'                4. Split size testing in Blowfish_String() routine.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added Public Event CurrentBlockSize() to track current
'                amount of data being processed.
'              - Updated GetBlockSize() routine.
'              - Moved progress bar code from BlockEncrypt() and
'                BlockDecrypt() routines to EncryptData() and DecryptData()
'                routines.
'              - Added new Property CreateNewFile().  Designates if creating
'                a new output file or overwriting the source file after
'                encryption or decryption.  Default value is TRUE.
'              - Updated Blowfish_File() routine to reference new property.
'              - Added new property to designate number of rounds to mix
'                primary key array.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME    As String = "clsBlowfish"
  Private Const MAX_FILE_SIZE  As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MAX_STRING_LEN As Long = &HFFFE&            ' 65534
  Private Const KB_32          As Long = &H8000&            ' 32768
  Private Const BLOCK_SIZE     As Long = 8
  Private Const KEY_ARRAY      As Long = 18
  Private Const CIPHER_ROUNDS  As Long = 16
  Private Const MAX_ROUNDS     As Long = 10
  Private Const MAX_BYTE       As Long = 256

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Byte)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a byt Password
' Variable name:     mabytPassword
'
' ***************************************************************************
  Private mobjMath32        As cMath32
  Private mblnCreateNewFile As Boolean
  Private mcurMaxSize       As Currency
  Private mcurAccrued       As Currency
  Private mcurOrigLength    As Currency
  Private mcurCurrByteCnt   As Currency
  Private mlngRounds        As Long
  Private mlngKeyRounds     As Long
  Private mlngKeyLength     As Long
  Private malngSBox()       As Long
  Private malngPrimaryKey() As Long
  Private mabytPassword()   As Byte


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event CipherProgress(ByVal lngProgress As Long)

' Track amount of data being processed.
Public Event CurrentBlockSize(ByVal lngBlockSize As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let CreateNewFile(ByVal blnNewFile As Boolean)
   
    ' 21-Feb-2012 Added Property to desiginate if input file
    '             is to be overwritten after encryption/decryption.
    '
    ' TRUE - Create new file to hold encrypted/decrypted data
    ' FALSE - Overwrite input file after encryption/decryption
    
    mblnCreateNewFile = blnNewFile

End Property

Public Property Let KeyLength(ByVal lngKeyLength As Long)
        
    ' 01-Jan-2009 Set increments of 32 bits for 32 to 448
    Select Case lngKeyLength
           Case 32, 64, 96, 128, 160, 192, 224, _
                256, 288, 320, 352, 384, 416, 448

                mlngKeyLength = lngKeyLength
                
           Case Else
                mlngKeyLength = 32  ' Use default value
    End Select
    
End Property

Public Property Let EncryptRounds(ByVal lngData As Long)
   
    ' 26-Dec-2008 Added multiple pass functionality.
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngRounds = lngData   ' Good selection
           Case Else:            mlngRounds = 1         ' Default value
    End Select
    
End Property

Public Property Let PrimaryKeyRounds(ByVal lngData As Long)
   
    ' 21-Feb-2011 Number of rounds to mix primary key array data
    '             within EvaluateKey() routine
    Select Case lngData
           Case 1 To 5: mlngKeyRounds = lngData   ' Good selection
           Case Else:   mlngKeyRounds = 1         ' Default value
    End Select
    
End Property

Public Property Let Password(ByVal vntData As Variant)

    ' Used when assigning a value to the password property, on the
    ' left side of an assignment.
    '
    ' Syntax:     X.Password = "User Supplied Password"
    '                            or
    '             X.Password = abytPwd()  <-- Highly recommend using a byte array
    
    Dim intIndex As Integer
    
    Const ROUTINE_NAME As String = "Password Property"
    
    On Error GoTo Password_Error

    Erase mabytPassword()
    
    ' Is password data in an array
    If CBool(IsArrayInitialized(vntData)) Then
    
        ReDim mabytPassword(UBound(vntData))
        
        For intIndex = 0 To UBound(vntData) - 1
            mabytPassword(intIndex) = CByte(vntData(intIndex))
        Next intIndex
    
    ' Is password data in raw string format
    ElseIf Len(Trim$(vntData)) > 0 Then
        mabytPassword() = StringToByteArray(CStr(Trim$(vntData)))
            
    ' no data passed
    Else
        InfoMsg "Missing or invalid password." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If

Password_CleanUp:
    On Error GoTo 0
    Exit Property

Password_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Password_CleanUp
    
End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Blowfish_File
'
' Description:   Function to encrypt/decrypt a file. 
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objBlowfish.Blowfish_File abytFile(), blnEncrypt             ' Encrypt/Decrypt data
' 
' Parameters:    abytFile() - Path/filename in byte array to be hashed
'                blnEncrypt - [Optional] - TRUE - Encrypt incoming data
'                             FALSE - Decrypt incoming data
'                             Default=TRUE
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call for file verification.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MixAppendedData() routine to encrypt/decrypt
'              size information that is to be appended to original data.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Moved Decryption size testing to earlier an position as per
'              Joe Sova's suggestion.
' ***************************************************************************
Public Function Blowfish_File(ByRef abytFile() As Byte, _
                     Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim strPath      As String
    Dim strSource    As String
    Dim strTarget    As String  ' Not used when overlaying source file
    Dim strTempFile  As String
    Dim strFileName  As String
    Dim hFile1       As Long
    Dim hFile2       As Long
    Dim lngBlockSize As Long
    Dim lngLastBlock As Long
    Dim curFilePos1  As Currency
    Dim curFilePos2  As Currency
    Dim curAmtLeft   As Currency
    Dim abytData()   As Byte
    Dim abytSize(11) As Byte
    Dim objBigFiles  As cBigFiles

    Const ROUTINE_NAME As String = "Blowfish_File"
    
    On Error GoTo Blowfish_File_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    Set objBigFiles = New cBigFiles   ' Instantiate class object
    Erase abytData()                  ' Always start with empty arrays
    Erase abytSize()
    ResetVariables                    ' Reset class variables
    
    If Not EvaluateKey() Then
        GoTo Blowfish_File_CleanUp
    End If
    
    strSource = ByteArrayToString(abytFile())           ' convert incoming array to string
    strPath = QualifyPath(GetPath(strSource))           ' Capture source path
    strFileName = GetFilename(strSource)                ' Capture source filename
    SetFileAttributes strSource, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes to normal
    
    ' See if a new file is to be created after
    ' performing encryption or decryption
    If mblnCreateNewFile Then
        If blnEncrypt Then
            ' Append "ENC" extension to target file
            ' designating the file is encrypted.
            strTarget = strPath & strFileName & ENCRYPT_EXT
        Else
            ' Append "DEC" extension to target file
            ' designating the file is decrypted.
            strTarget = strPath & strFileName & DECRYPT_EXT
        End If
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo Blowfish_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file
        
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo Blowfish_File_CleanUp
        End If
            
        ' If performing decryption then do not
        ' process the last few bytes of the file
        ' because this is where we stored the
        ' size of the original file.
        If blnEncrypt Then
            mcurOrigLength = curAmtLeft   ' Capture total byte count
        Else
            curAmtLeft = curAmtLeft - (UBound(abytSize) + 1)  ' Adjust amount that has to be decrypted
            curFilePos1 = curAmtLeft                          ' Set pointer to capture the last 12 bytes
            
            If Not .OpenReadOnly(strSource, hFile1) Then      ' Open the source file
                gblnStopProcessing = True
            End If
            
            .API_ReadFile hFile1, curFilePos1, abytSize()     ' Capture last 12 bytes of source file
            .API_CloseFile hFile1                             ' close source file
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo Blowfish_File_CleanUp
            End If
    
            MixAppendedData abytSize()                         ' Decrypt appended data
            CopyMemory lngLastBlock, abytSize(0), 4&           ' Size of last block of data without padding [0-3]
            CopyMemory mcurOrigLength, abytSize(4), 8&         ' Capture original file size [4-11]
            
            ' Decryption error test
            If (lngLastBlock < 0) Or _
               (lngLastBlock > MAX_FILE_SIZE) Or _
               (curAmtLeft < 1) Or _
               (mcurOrigLength < 0) Or _
               (mcurOrigLength > MAX_FILE_SIZE) Then
            
                InfoMsg "This file cannot be Decrypted, it is" & vbNewLine & _
                        "either not Encrypted, too large or corrupt." & vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curAmtLeft, "#,##0") & " bytes" & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                        "Max allowed size:  " & Format$(MAX_FILE_SIZE, "#,##0") & " bytes" & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                
                gblnStopProcessing = True
                GoTo Blowfish_File_CleanUp
                
            End If
        End If
        
        mcurMaxSize = mcurOrigLength * mlngRounds   ' Save file size for progress calcs
        
        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@
        
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo Blowfish_File_CleanUp
        End If
        
        strTempFile = CreateTempFile()  ' Create a temp target file
        
        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo Blowfish_File_CleanUp
        End If
        
        Do
            Erase abytData()                              ' empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)       ' Process the source file in chunks
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array
            
            If blnEncrypt Then
                lngLastBlock = lngBlockSize               ' Save size of last block of data
            End If
            
            ' Read source file
            If Not .API_ReadFile(hFile1, curFilePos1, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If
                       
            curFilePos1 = curFilePos1 + CCur(UBound(abytData) + 1)  ' Adjust pointers accordingly
            
            If blnEncrypt Then
        
                ' Encrypt data
                If Not EncryptData(abytData()) Then
                    gblnStopProcessing = True
                    Exit Do  ' exit Do..Loop
                End If
            
            Else
                ' Decrypt data
                If DecryptData(abytData()) Then
                
                    If curAmtLeft < 1 Then
                        ReDim Preserve abytData(lngLastBlock - 1)   ' Resize last data chunk
                    End If
                
                Else
                    gblnStopProcessing = True
                    Exit Do  ' exit Do..Loop
                End If
            
            End If
                
            ' Write to target file
            If Not .API_WriteFile(hFile2, curFilePos2, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If
                       
            curFilePos2 = curFilePos2 + CCur(UBound(abytData) + 1)    ' Adjust pointers accordingly
            
            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do  ' exit Do..Loop
            End If
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If
    
        Loop
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo Blowfish_File_CleanUp
        End If
    
        .API_CloseFile hFile1    ' Close all files opened by this routine
        .API_CloseFile hFile2
        
        ' Add important data to end of encrypted file
        If blnEncrypt Then
            
            CopyMemory abytSize(0), lngLastBlock, 4&         ' Save size of last block of data without padding [0-3]
            CopyMemory abytSize(4), mcurOrigLength, 8&       ' Save original file size [4-11]
            .CalcFileSize strTempFile, curFilePos2           ' Get size of target file
            MixAppendedData abytSize()                       ' Encrypt appended data
            
            ' Excessive DoEvents are to keep the write
            ' process in sync with the CPU speed
            DoEvents
            .OpenReadWrite strTempFile, hFile2               ' open target file
            .API_WriteFile hFile2, curFilePos2, abytSize()   ' Append to target file
            .API_CloseFile hFile2                            ' Close target file
            DoEvents
        Else
            ' Final decryption verification
            curAmtLeft = 0@                         ' Verify empty variable
            .CalcFileSize strTempFile, curAmtLeft   ' Get size of decrypted file
            
            ' Compare current file size with
            ' what it is supposed to be
            If (curAmtLeft <> mcurOrigLength) Then
            
                InfoMsg "This file cannot be Decrypted." & vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curAmtLeft, "#,##0") & " bytes" & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                
                gblnStopProcessing = True
                GoTo Blowfish_File_CleanUp
                
            End If

        End If
    End With
    
    ' Move temp file to destination location
    DoEvents
    On Error Resume Next
    DoEvents

    ' See if a new file is to be created after
    ' performing encryption or decryption
    If mblnCreateNewFile Then
        ' Create new file
        MoveFileEx strTempFile, strTarget, _
                   MOVEFILE_COPY_ALLOWED Or _
                   MOVEFILE_REPLACE_EXISTING
    Else
        ' Overwrite source file
        MoveFileEx strTempFile, strSource, _
                   MOVEFILE_COPY_ALLOWED Or _
                   MOVEFILE_REPLACE_EXISTING
    End If

    RaiseEvent CipherProgress(100)   ' Update progress bar

Blowfish_File_CleanUp:
    With objBigFiles
        .API_CloseFile hFile1   ' Close all files opened by this routine
        .API_CloseFile hFile2
    End With
    Set objBigFiles = Nothing   ' Free class object from memory
    
    ResetVariables    ' Reset class variables
    Erase abytData()  ' Always empty arrays when not needed
    Erase abytSize()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Blowfish_File = False
    Else
        Blowfish_File = True
    End If
        
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Blowfish_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Blowfish_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Blowfish_String
'
' Description:   Function to encrypt/decrypt a string of data. 
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)          ' convert string data to byte array
'                objBlowfish.Blowfish_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
' 
' Parameters:    abytData() - data string in byte array
'                blnEncrypt - [Optional] - TRUE - Encrypt incoming data
'                             FALSE - Decrypt incoming data
'                             Default=TRUE
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MixAppendedData() routine to encrypt/decrypt
'              size information that is to be appended to original data.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Split size testing as per Joe Sova's suggestion.
' ***************************************************************************
Public Function Blowfish_String(ByRef abytData() As Byte, _
                       Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte
    
    Const ROUTINE_NAME As String = "Blowfish_String"

    On Error GoTo Blowfish_String_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True      ' Set flag to stop processing
        Exit Function
    End If
    
    Erase abytSize()                                    ' Always start with empty arrays
    ResetVariables                                      ' Reset class variables
    mcurMaxSize = CCur(UBound(abytData)) * mlngRounds   ' Save file size for progress calcs
    mcurAccrued = 0@                                    ' Init accrued progress
    
    If Not EvaluateKey() Then
        GoTo Blowfish_String_CleanUp
    End If
    
    ' encrypt the data
    If blnEncrypt Then
        
        mcurOrigLength = CCur(UBound(abytData))         ' Get string length for progress display
        CopyMemory abytSize(0), mcurOrigLength, 8&      ' Copy original length to temp array
        MixAppendedData abytSize()                      ' Encrypt appended data
        
        If EncryptData(abytData()) Then
            
            curDataLength = UBound(abytData) + 1        ' data length after encryption
            ReDim Preserve abytData(curDataLength + 7)  ' Resize to hold original length
            
            ' Append original data length to end of encrypted data
            CopyMemory abytData(curDataLength), abytSize(0), 8&
            
        Else
            gblnStopProcessing = True
        End If
    
    Else
        ' Decrypt the data
        CopyMemory abytSize(0), abytData(UBound(abytData) - 7), 8&   ' capture the original data length
        ReDim Preserve abytData(UBound(abytData) - 8)                ' Adjust data length
        MixAppendedData abytSize()                                   ' Decrypt appended data
        CopyMemory mcurOrigLength, abytSize(0), 8&                   ' capture the original data length
        
        ' Must have data and cannot exceed size of textbox
        If (UBound(abytData) < 1) Or _
           (mcurOrigLength > MAX_STRING_LEN) Then
           
            InfoMsg "This data cannot be Decrypted, it is" & vbNewLine & _
                    "either not Encrypted, too large or corrupt." & vbNewLine & vbNewLine & _
                    "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                    "Max allowed size:  " & Format$(MAX_STRING_LEN, "#,##0") & " bytes" & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                
            gblnStopProcessing = True
            GoTo Blowfish_String_CleanUp
            
        End If
        
        ' Decrypt the data
        If DecryptData(abytData()) Then
                
            ReDim Preserve abytData(mcurOrigLength)  ' Resize data to exact length
            curDataLength = CCur(UBound(abytData))   ' data length after decryption
            
            ' Compare data length
            If curDataLength <> mcurOrigLength Then
                
                InfoMsg "Encrypted data string is corrupted.  Cannot Decrypt." & _
                        vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curDataLength, "#,##0") & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    
                gblnStopProcessing = True
            
            End If
        Else
            gblnStopProcessing = True
        End If
    
    End If
    
Blowfish_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Blowfish_String = False
    Else
        Blowfish_String = True
    End If
    
    ResetVariables    ' Reset class variables
    Erase abytSize()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Blowfish_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Blowfish_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey() As Boolean

    ' Called by Blowfish_File()
    '           Blowfish_String()
    
    Dim lngIndex     As Long
    Dim lngIdx       As Long
    Dim lngHold      As Long
    Dim lngKeyLen    As Long
    Dim lngPwdIdx    As Long
    Dim lngWholeWord As Long
    Dim lngLeftWord  As Long
    Dim lngRightWord As Long
    Dim abytPwd()    As Byte

    Const ROUTINE_NAME As String = "EvaluateKey"

    On Error GoTo EvaluateKey_Error

    EvaluateKey = False  ' Preset flag
    
    If Not CBool(IsArrayInitialized(mabytPassword)) Then
        InfoMsg "Need a password before data can be encrypted/decrypted." & _
                vbNewLine & vbNewLine & _
                "Min length - " & MIN_PWD_LENGTH & vbNewLine & _
                "Max length - " & MAX_PWD_LENGTH & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    Erase abytPwd()  ' Always start with empty arrays
    
    LoadWorkBoxes                                        ' Preload work arrays
    lngKeyLen = mlngKeyLength \ 8                        ' Ex:  16 bytes = 128 bits \ 8 bits
    abytPwd() = ExpandData(mabytPassword(), lngKeyLen)   ' Expand password key to desired length

    lngPwdIdx = 0                   ' Set temp password array index
    lngHold = mlngRounds            ' 02-Nov-2009 Capture number of rounds for block encryption
    EncryptRounds = mlngKeyRounds   ' 02-Nov-2009 Temp reset number of rounds of block
                                    '             encryption (1-5 rounds is a good mix)
    
    ' Create key-dependant p-boxes
    For lngIndex = 0 To (KEY_ARRAY - 1)
        
        lngWholeWord = 0   ' Preset to zero

        ' Create a unique long integer
        For lngIdx = 1 To 4
            
            ' Overlay last three bytes with first three bytes
            MoveMemory ByVal VarPtr(lngWholeWord) + 1, lngWholeWord, 3&
            lngWholeWord = lngWholeWord Or abytPwd(lngPwdIdx Mod lngKeyLen)
            lngPwdIdx = lngPwdIdx + 1
            
        Next lngIdx

        ' Update primary key array
        malngPrimaryKey(lngIndex) = malngPrimaryKey(lngIndex) Xor lngWholeWord

    Next lngIndex

    ' Additional mixing of the password key data
    ' 29-Jan-2010 Removed outside loop.  Updated
    '             number of encrypt rounds instead
    For lngIndex = 0 To (KEY_ARRAY - 2) Step 2

        lngLeftWord = malngPrimaryKey(lngIndex)        ' Capture 2 long integers
        lngRightWord = malngPrimaryKey(lngIndex + 1)
        BlockEncrypt lngLeftWord, lngRightWord         ' Perform block encryption
        malngPrimaryKey(lngIndex) = lngLeftWord        ' Insert data back into array
        malngPrimaryKey(lngIndex + 1) = lngRightWord

    Next lngIndex
    
    ' Additional mixing of the S-Box data
    ' 29-Jan-2010 Removed outside loop.  Updated
    '             number of encrypt rounds instead
    For lngIndex = 0 To 3
        For lngIdx = 0 To (MAX_BYTE - 2) Step 2
            
            lngLeftWord = malngSBox(lngIndex, lngIdx)        ' Capture 2 long integers
            lngRightWord = malngSBox(lngIndex, lngIdx + 1)
            BlockEncrypt lngLeftWord, lngRightWord           ' Perform block encryption
            malngSBox(lngIndex, lngIdx) = lngLeftWord        ' Insert data back into array
            malngSBox(lngIndex, lngIdx + 1) = lngRightWord

        Next lngIdx
    Next lngIndex
        
EvaluateKey_CleanUp:
    Erase abytPwd()           ' Always empty arrays when not needed
    EncryptRounds = lngHold   ' 02-Nov-2009 Reset number of rounds of block
                              '             encryption back to original value
    ' An error occurred or user
    ' opted to STOP processing
    If gblnStopProcessing Then
        EvaluateKey = False
    Else
        EvaluateKey = True
    End If

    On Error GoTo 0           ' Nullify this error trap
    Exit Function

EvaluateKey_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ResetVariables
    gblnStopProcessing = True
    Resume EvaluateKey_CleanUp

End Function

Private Function EncryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Blowfish_File()
    '           Blowfish_String()
    
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngLeftWord   As Long
    Dim lngRightWord  As Long
    Dim lngDataLength As Long
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte

    Const ROUTINE_NAME As String = "EncryptData"

    On Error GoTo EncryptData_Error

    lngDataLength = UBound(abytData) + 1  ' Get size of incoming array
    
    ' the lngDataLength must be a multiple of 8 bytes
    If (lngDataLength Mod BLOCK_SIZE <> 0) Then
        lngDataLength = (lngDataLength + BLOCK_SIZE) - (lngDataLength Mod BLOCK_SIZE)
        ReDim Preserve abytData(lngDataLength - 1)
    End If

    ' 1-Jan-2009 Fixed lngOffSet range
    ' Encrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Transfer incoming data to a temp work array
        CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE
                
        ' Convert bytes to two long integers
        GetWord lngLeftWord, abytInput(), 0&
        GetWord lngRightWord, abytInput(), 4&

        ' Encrypt block of data
        BlockEncrypt lngLeftWord, lngRightWord
        
        ' Convert two long integers to bytes
        PutWord lngLeftWord, abytOutput(), 0&
        PutWord lngRightWord, abytOutput(), 4&
                
        ' Transfer encrypted data back to original array
        CopyMemory abytData(lngOffSet), abytOutput(0), BLOCK_SIZE

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
    Next lngOffSet
    
    ' 19-Feb-2012 Update progress bar. Separate
    '             steps for easier debugging
    mcurAccrued = mcurAccrued + CCur(lngDataLength * mlngRounds)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent CipherProgress(lngProgress)
            
EncryptData_CleanUp:
    DoEvents
    If gblnStopProcessing Then
        EncryptData = False   ' An error occurred or user opted to STOP processing
    Else
        EncryptData = True    ' Good Finish
    End If

    On Error GoTo 0
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

Private Function DecryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Blowfish_File()
    '           Blowfish_String()
    
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngLeftWord   As Long
    Dim lngRightWord  As Long
    Dim lngDataLength As Long
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte

    Const ROUTINE_NAME As String = "DecryptData"

    On Error GoTo DecryptData_Error

    lngDataLength = UBound(abytData) + 1  ' Get size of incoming array
    
    ' 1-Jan-2009 Fixed lngOffSet range
    ' Decrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Transfer incoming data to a temp work array
        CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE
                
        ' Convert bytes to two long integers
        GetWord lngLeftWord, abytInput(), 0&
        GetWord lngRightWord, abytInput(), 4&

        ' Decrypt block of data
        BlockDecrypt lngLeftWord, lngRightWord
                
        ' Convert two long integers to bytes
        PutWord lngLeftWord, abytOutput(), 0&
        PutWord lngRightWord, abytOutput(), 4&
                
        ' Transfer decrypted data back to the original array
        CopyMemory abytData(lngOffSet), abytOutput(0), BLOCK_SIZE

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
    Next lngOffSet
    
    ' 19-Feb-2012 Update progress bar. Separate
    '             steps for easier debugging
    mcurAccrued = mcurAccrued + CCur(lngDataLength * mlngRounds)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent CipherProgress(lngProgress)
            
DecryptData_CleanUp:
    DoEvents
    If gblnStopProcessing Then
        DecryptData = False   ' An error occurred or user opted to STOP processing
    Else
        DecryptData = True    ' Good Finish
    End If

    On Error GoTo 0
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

Private Static Sub BlockEncrypt(ByRef lngLeftWord As Long, _
                                ByRef lngRightWord As Long)

    ' Called by EncryptData()
    
    Dim lngIdx   As Long
    Dim lngTemp  As Long
    Dim lngIndex As Long
    Dim lngRound As Long
    
    For lngRound = 1 To mlngRounds
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        lngIdx = 0
    
        For lngIndex = 0 To (CIPHER_ROUNDS \ 2 - 1)
    
            lngLeftWord = lngLeftWord Xor malngPrimaryKey(lngIdx)
            lngRightWord = lngRightWord Xor CalcNewValue(lngLeftWord)
            
            lngRightWord = lngRightWord Xor malngPrimaryKey(lngIdx + 1)
            lngLeftWord = lngLeftWord Xor CalcNewValue(lngRightWord)
            lngIdx = lngIdx + 2
    
        Next lngIndex
    
        lngTemp = lngRightWord
        lngRightWord = lngLeftWord Xor malngPrimaryKey(CIPHER_ROUNDS)
        lngLeftWord = lngTemp Xor malngPrimaryKey(CIPHER_ROUNDS + 1)
        
    Next lngRound
    
End Sub

Private Static Sub BlockDecrypt(ByRef lngLeftWord As Long, _
                                ByRef lngRightWord As Long)

    ' Called by DecryptData()
    
    Dim lngIdx   As Long
    Dim lngTemp  As Long
    Dim lngIndex As Long
    Dim lngRound As Long

    For lngRound = 1 To mlngRounds
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        lngTemp = lngRightWord
        lngRightWord = lngLeftWord Xor malngPrimaryKey(CIPHER_ROUNDS + 1)
        lngLeftWord = lngTemp Xor malngPrimaryKey(CIPHER_ROUNDS)
        lngIdx = CIPHER_ROUNDS - 2
        
        For lngIndex = 0 To (CIPHER_ROUNDS \ 2 - 1)
    
            lngLeftWord = lngLeftWord Xor CalcNewValue(lngRightWord)
            lngRightWord = lngRightWord Xor malngPrimaryKey(lngIdx + 1)
            
            lngRightWord = lngRightWord Xor CalcNewValue(lngLeftWord)
            lngLeftWord = lngLeftWord Xor malngPrimaryKey(lngIdx)
            lngIdx = lngIdx - 2
    
        Next lngIndex

    Next lngRound
    
End Sub

Private Function CalcNewValue(ByVal lngValue As Long) As Long

    ' Called by BlockEncrypt()
    '           BlockDecrypt()
    
    Dim abytData(4) As Byte
    Dim lngTemp     As Long  ' Added variable for easier maintenance
    
    CopyMemory abytData(0), lngValue, 4&   ' Convert long integer to byte array
    
    With mobjMath32
        lngTemp = .UnsignedAdd(malngSBox(0, abytData(3)), malngSBox(1, abytData(2)))
        lngTemp = lngTemp Xor malngSBox(2, abytData(1))
        CalcNewValue = .UnsignedAdd(lngTemp, malngSBox(3, abytData(0)))
    End With
    
End Function

Private Sub GetWord(ByRef lngValue As Long, _
                    ByRef abytData() As Byte, _
                    ByVal lngIndex As Long)

    ' Called by EncryptData()
    '           DecryptData()
    '
    ' Convert byte array to long integer in
    ' Big-Endian format
    
    Dim abytTemp(4) As Byte
    
    abytTemp(3) = abytData(lngIndex)       ' Invert the bytes
    abytTemp(2) = abytData(lngIndex + 1)
    abytTemp(1) = abytData(lngIndex + 2)
    abytTemp(0) = abytData(lngIndex + 3)
    
    CopyMemory lngValue, abytTemp(0), 4&   ' Convert byte array to long integer
  
End Sub

Private Sub PutWord(ByVal lngValue As Long, _
                    ByRef abytData() As Byte, _
                    ByVal lngIndex As Long)

    ' Called by EncryptData()
    '           DecryptData()
    '
    ' Convert long integer to byte array in
    ' Little-Endian format
    
    Dim abytTemp(4) As Byte
    
    CopyMemory abytTemp(0), lngValue, 4&   ' Convert long integer to byte array
    
    abytData(lngIndex) = abytTemp(3)       ' Invert the bytes
    abytData(lngIndex + 1) = abytTemp(2)
    abytData(lngIndex + 2) = abytTemp(1)
    abytData(lngIndex + 3) = abytTemp(0)
    
End Sub

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the data to be processed.
'
' Parameters:    curAmtLeft - Amount of data left
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= KB_32: GetBlockSize = KB_32
           Case Else:        GetBlockSize = CLng(curAmtLeft)
    End Select
    
    RaiseEvent CurrentBlockSize(GetBlockSize)
    
End Function

Private Sub LoadWorkBoxes()

    ' Called by EvaluateKey()
    
    Dim lngIndex As Long
    Dim strData  As String
    Dim avntData As Variant

    ReDim malngPrimaryKey(KEY_ARRAY)
    ReDim malngSBox(0 To 3, MAX_BYTE)
    
    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Empty variant

    ' Reference:  DI Management Services Pty Ltd, Sydney Australia
    '             http://www.di-mgt.com.au/src/blowfish.txt
    '
    ' password key data
    strData = strData & "243F6A88 85A308D3 13198A2E 03707344 A4093822 299F31D0 "
    strData = strData & "082EFA98 EC4E6C89 452821E6 38D01377 BE5466CF 34E90C6C "
    strData = strData & "C0AC29B7 C97C50DD 3F84D5B5 B5470917 9216D5D9 8979FB1B "

    ' Separate into a variant array
    avntData = Split(strData, Chr$(32))

    ' Transfer data to password key array
    For lngIndex = 0 To (KEY_ARRAY - 1)
        malngPrimaryKey(lngIndex) = Val("&H" & avntData(lngIndex))
    Next lngIndex

    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Empty variant

    ' Initialize s-boxes
    strData = strData & "D1310BA6 98DFB5AC 2FFD72DB D01ADFB7 B8E1AFED 6A267E96 BA7C9045 F12C7F99 "
    strData = strData & "24A19947 B3916CF7 0801F2E2 858EFC16 636920D8 71574E69 A458FEA3 F4933D7E "
    strData = strData & "0D95748F 728EB658 718BCD58 82154AEE 7B54A41D C25A59B5 9C30D539 2AF26013 "
    strData = strData & "C5D1B023 286085F0 CA417918 B8DB38EF 8E79DCB0 603A180E 6C9E0E8B B01E8A3E "
    strData = strData & "D71577C1 BD314B27 78AF2FDA 55605C60 E65525F3 AA55AB94 57489862 63E81440 "
    strData = strData & "55CA396A 2AAB10B6 B4CC5C34 1141E8CE A15486AF 7C72E993 B3EE1411 636FBC2A "
    strData = strData & "2BA9C55D 741831F6 CE5C3E16 9B87931E AFD6BA33 6C24CF5C 7A325381 28958677 "
    strData = strData & "3B8F4898 6B4BB9AF C4BFE81B 66282193 61D809CC FB21A991 487CAC60 5DEC8032 "
    strData = strData & "EF845D5D E98575B1 DC262302 EB651B88 23893E81 D396ACC5 0F6D6FF3 83F44239 "
    strData = strData & "2E0B4482 A4842004 69C8F04A 9E1F9B5E 21C66842 F6E96C9A 670C9C61 ABD388F0 "
    strData = strData & "6A51A0D2 D8542F68 960FA728 AB5133A3 6EEF0B6C 137A3BE4 BA3BF050 7EFB2A98 "
    strData = strData & "A1F1651D 39AF0176 66CA593E 82430E88 8CEE8619 456F9FB4 7D84A5C3 3B8B5EBE "
    strData = strData & "E06F75D8 85C12073 401A449F 56C16AA6 4ED3AA62 363F7706 1BFEDF72 429B023D "
    strData = strData & "37D0D724 D00A1248 DB0FEAD3 49F1C09B 075372C9 80991B7B 25D479D8 F6E8DEF7 "
    strData = strData & "E3FE501A B6794C3B 976CE0BD 04C006BA C1A94FB6 409F60C4 5E5C9EC2 196A2463 "
    strData = strData & "68FB6FAF 3E6C53B5 1339B2EB 3B52EC6F 6DFC511F 9B30952C CC814544 AF5EBD09 "
    strData = strData & "BEE3D004 DE334AFD 660F2807 192E4BB3 C0CBA857 45C8740F D20B5F39 B9D3FBDB "
    strData = strData & "5579C0BD 1A60320A D6A100C6 402C7279 679F25FE FB1FA3CC 8EA5E9F8 DB3222F8 "
    strData = strData & "3C7516DF FD616B15 2F501EC8 AD0552AB 323DB5FA FD238760 53317B48 3E00DF82 "
    strData = strData & "9E5C57BB CA6F8CA0 1A87562E DF1769DB D542A8F6 287EFFC3 AC6732C6 8C4F5573 "
    strData = strData & "695B27B0 BBCA58C8 E1FFA35D B8F011A0 10FA3D98 FD2183B8 4AFCB56C 2DD1D35B "
    strData = strData & "9A53E479 B6F84565 D28E49BC 4BFB9790 E1DDF2DA A4CB7E33 62FB1341 CEE4C6E8 "
    strData = strData & "EF20CADA 36774C01 D07E9EFE 2BF11FB4 95DBDA4D AE909198 EAAD8E71 6B93D5A0 "
    strData = strData & "D08ED1D0 AFC725E0 8E3C5B2F 8E7594B7 8FF6E2FB F2122B64 8888B812 900DF01C "
    strData = strData & "4FAD5EA0 688FC31C D1CFF191 B3A8C1AD 2F2F2218 BE0E1777 EA752DFE 8B021FA1 "
    strData = strData & "E5A0CC0F B56F74E8 18ACF3D6 CE89E299 B4A84FE0 FD13E0B7 7CC43B81 D2ADA8D9 "
    strData = strData & "165FA266 80957705 93CC7314 211A1477 E6AD2065 77B5FA86 C75442F5 FB9D35CF "
    strData = strData & "EBCDAF0C 7B3E89A0 D6411BD3 AE1E7E49 00250E2D 2071B35E 226800BB 57B8E0AF "
    strData = strData & "2464369B F009B91E 5563911D 59DFA6AA 78C14389 D95A537F 207D5BA2 02E5B9C5 "
    strData = strData & "83260376 6295CFA9 11C81968 4E734A41 B3472DCA 7B14A94A 1B510052 9A532915 "
    strData = strData & "D60F573F BC9BC6E4 2B60A476 81E67400 08BA6FB5 571BE91F F296EC6B 2A0DD915 "
    strData = strData & "B6636521 E7B9F9B6 FF34052E C5855664 53B02D5D A99F8FA1 08BA4799 6E85076A "

    ' Separate into a variant array
    avntData = Split(strData, Chr$(32))

    ' Transfer data to contants array
    For lngIndex = 0 To (MAX_BYTE - 1)
        malngSBox(0, lngIndex) = Val("&H" & avntData(lngIndex))
    Next lngIndex

    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Empty variant

    strData = strData & "4B7A70E9 B5B32944 DB75092E C4192623 AD6EA6B0 49A7DF7D 9CEE60B8 8FEDB266 "
    strData = strData & "ECAA8C71 699A17FF 5664526C C2B19EE1 193602A5 75094C29 A0591340 E4183A3E "
    strData = strData & "3F54989A 5B429D65 6B8FE4D6 99F73FD6 A1D29C07 EFE830F5 4D2D38E6 F0255DC1 "
    strData = strData & "4CDD2086 8470EB26 6382E9C6 021ECC5E 09686B3F 3EBAEFC9 3C971814 6B6A70A1 "
    strData = strData & "687F3584 52A0E286 B79C5305 AA500737 3E07841C 7FDEAE5C 8E7D44EC 5716F2B8 "
    strData = strData & "B03ADA37 F0500C0D F01C1F04 0200B3FF AE0CF51A 3CB574B2 25837A58 DC0921BD "
    strData = strData & "D19113F9 7CA92FF6 94324773 22F54701 3AE5E581 37C2DADC C8B57634 9AF3DDA7 "
    strData = strData & "A9446146 0FD0030E ECC8C73E A4751E41 E238CD99 3BEA0E2F 3280BBA1 183EB331 "
    strData = strData & "4E548B38 4F6DB908 6F420D03 F60A04BF 2CB81290 24977C79 5679B072 BCAF89AF "
    strData = strData & "DE9A771F D9930810 B38BAE12 DCCF3F2E 5512721F 2E6B7124 501ADDE6 9F84CD87 "
    strData = strData & "7A584718 7408DA17 BC9F9ABC E94B7D8C EC7AEC3A DB851DFA 63094366 C464C3D2 "
    strData = strData & "EF1C1847 3215D908 DD433B37 24C2BA16 12A14D43 2A65C451 50940002 133AE4DD "
    strData = strData & "71DFF89E 10314E55 81AC77D6 5F11199B 043556F1 D7A3C76B 3C11183B 5924A509 "
    strData = strData & "F28FE6ED 97F1FBFA 9EBABF2C 1E153C6E 86E34570 EAE96FB1 860E5E0A 5A3E2AB3 "
    strData = strData & "771FE71C 4E3D06FA 2965DCB9 99E71D0F 803E89D6 5266C825 2E4CC978 9C10B36A "
    strData = strData & "C6150EBA 94E2EA78 A5FC3C53 1E0A2DF4 F2F74EA7 361D2B3D 1939260F 19C27960 "
    strData = strData & "5223A708 F71312B6 EBADFE6E EAC31F66 E3BC4595 A67BC883 B17F37D1 018CFF28 "
    strData = strData & "C332DDEF BE6C5AA5 65582185 68AB9802 EECEA50F DB2F953B 2AEF7DAD 5B6E2F84 "
    strData = strData & "1521B628 29076170 ECDD4775 619F1510 13CCA830 EB61BD96 0334FE1E AA0363CF "
    strData = strData & "B5735C90 4C70A239 D59E9E0B CBAADE14 EECC86BC 60622CA7 9CAB5CAB B2F3846E "
    strData = strData & "648B1EAF 19BDF0CA A02369B9 655ABB50 40685A32 3C2AB4B3 319EE9D5 C021B8F7 "
    strData = strData & "9B540B19 875FA099 95F7997E 623D7DA8 F837889A 97E32D77 11ED935F 16681281 "
    strData = strData & "0E358829 C7E61FD6 96DEDFA1 7858BA99 57F584A5 1B227263 9B83C3FF 1AC24696 "
    strData = strData & "CDB30AEB 532E3054 8FD948E4 6DBC3128 58EBF2EF 34C6FFEA FE28ED61 EE7C3C73 "
    strData = strData & "5D4A14D9 E864B7E3 42105D14 203E13E0 45EEE2B6 A3AAABEA DB6C4F15 FACB4FD0 "
    strData = strData & "C742F442 EF6ABBB5 654F3B1D 41CD2105 D81E799E 86854DC7 E44B476A 3D816250 "
    strData = strData & "CF62A1F2 5B8D2646 FC8883A0 C1C7B6A3 7F1524C3 69CB7492 47848A0B 5692B285 "
    strData = strData & "095BBF00 AD19489D 1462B174 23820E00 58428D2A 0C55F5EA 1DADF43E 233F7061 "
    strData = strData & "3372F092 8D937E41 D65FECF1 6C223BDB 7CDE3759 CBEE7460 4085F2A7 CE77326E "
    strData = strData & "A6078084 19F8509E E8EFD855 61D99735 A969A7AA C50C06C2 5A04ABFC 800BCADC "
    strData = strData & "9E447A2E C3453484 FDD56705 0E1E9EC9 DB73DBD3 105588CD 675FDA79 E3674340 "
    strData = strData & "C5C43465 713E38D8 3D28F89E F16DFF20 153E21E7 8FB03D4A E6E39F2B DB83ADF7 "

    ' Separate into a variant array
    avntData = Split(strData, Chr$(32))

    ' Transfer data to contants array
    For lngIndex = 0 To (MAX_BYTE - 1)
        malngSBox(1, lngIndex) = Val("&H" & avntData(lngIndex))
    Next lngIndex

    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Empty variant

    strData = strData & "E93D5A68 948140F7 F64C261C 94692934 411520F7 7602D4F7 BCF46B2E D4A20068 "
    strData = strData & "D4082471 3320F46A 43B7D4B7 500061AF 1E39F62E 97244546 14214F74 BF8B8840 "
    strData = strData & "4D95FC1D 96B591AF 70F4DDD3 66A02F45 BFBC09EC 03BD9785 7FAC6DD0 31CB8504 "
    strData = strData & "96EB27B3 55FD3941 DA2547E6 ABCA0A9A 28507825 530429F4 0A2C86DA E9B66DFB "
    strData = strData & "68DC1462 D7486900 680EC0A4 27A18DEE 4F3FFEA2 E887AD8C B58CE006 7AF4D6B6 "
    strData = strData & "AACE1E7C D3375FEC CE78A399 406B2A42 20FE9E35 D9F385B9 EE39D7AB 3B124E8B "
    strData = strData & "1DC9FAF7 4B6D1856 26A36631 EAE397B2 3A6EFA74 DD5B4332 6841E7F7 CA7820FB "
    strData = strData & "FB0AF54E D8FEB397 454056AC BA489527 55533A3A 20838D87 FE6BA9B7 D096954B "
    strData = strData & "55A867BC A1159A58 CCA92963 99E1DB33 A62A4A56 3F3125F9 5EF47E1C 9029317C "
    strData = strData & "FDF8E802 04272F70 80BB155C 05282CE3 95C11548 E4C66D22 48C1133F C70F86DC "
    strData = strData & "07F9C9EE 41041F0F 404779A4 5D886E17 325F51EB D59BC0D1 F2BCC18F 41113564 "
    strData = strData & "257B7834 602A9C60 DFF8E8A3 1F636C1B 0E12B4C2 02E1329E AF664FD1 CAD18115 "
    strData = strData & "6B2395E0 333E92E1 3B240B62 EEBEB922 85B2A20E E6BA0D99 DE720C8C 2DA2F728 "
    strData = strData & "D0127845 95B794FD 647D0862 E7CCF5F0 5449A36F 877D48FA C39DFD27 F33E8D1E "
    strData = strData & "0A476341 992EFF74 3A6F6EAB F4F8FD37 A812DC60 A1EBDDF8 991BE14C DB6E6B0D "
    strData = strData & "C67B5510 6D672C37 2765D43B DCD0E804 F1290DC7 CC00FFA3 B5390F92 690FED0B "
    strData = strData & "667B9FFB CEDB7D9C A091CF0B D9155EA3 BB132F88 515BAD24 7B9479BF 763BD6EB "
    strData = strData & "37392EB3 CC115979 8026E297 F42E312D 6842ADA7 C66A2B3B 12754CCC 782EF11C "
    strData = strData & "6A124237 B79251E7 06A1BBE6 4BFB6350 1A6B1018 11CAEDFA 3D25BDD8 E2E1C3C9 "
    strData = strData & "44421659 0A121386 D90CEC6E D5ABEA2A 64AF674E DA86A85F BEBFE988 64E4C3FE "
    strData = strData & "9DBC8057 F0F7C086 60787BF8 6003604D D1FD8346 F6381FB0 7745AE04 D736FCCC "
    strData = strData & "83426B33 F01EAB71 B0804187 3C005E5F 77A057BE BDE8AE24 55464299 BF582E61 "
    strData = strData & "4E58F48F F2DDFDA2 F474EF38 8789BDC2 5366F9C3 C8B38E74 B475F255 46FCD9B9 "
    strData = strData & "7AEB2661 8B1DDF84 846A0E79 915F95E2 466E598E 20B45770 8CD55591 C902DE4C "
    strData = strData & "B90BACE1 BB8205D0 11A86248 7574A99E B77F19B6 E0A9DC09 662D09A1 C4324633 "
    strData = strData & "E85A1F02 09F0BE8C 4A99A025 1D6EFE10 1AB93D1D 0BA5A4DF A186F20F 2868F169 "
    strData = strData & "DCB7DA83 573906FE A1E2CE9B 4FCD7F52 50115E01 A70683FA A002B5C4 0DE6D027 "
    strData = strData & "9AF88C27 773F8641 C3604C06 61A806B5 F0177A28 C0F586E0 006058AA 30DC7D62 "
    strData = strData & "11E69ED7 2338EA63 53C2DD94 C2C21634 BBCBEE56 90BCB6DE EBFC7DA1 CE591D76 "
    strData = strData & "6F05E409 4B7C0188 39720A3D 7C927C24 86E3725F 724D9DB9 1AC15BB4 D39EB8FC "
    strData = strData & "ED545578 08FCA5B5 D83D7CD3 4DAD0FC4 1E50EF5E B161E6F8 A28514D9 6C51133C "
    strData = strData & "6FD5C7E7 56E14EC4 362ABFCE DDC6C837 D79A3234 92638212 670EFA8E 406000E0 "

    ' Separate into a variant array
    avntData = Split(strData, Chr$(32))

    ' Transfer data to contants array
    For lngIndex = 0 To (MAX_BYTE - 1)
        malngSBox(2, lngIndex) = Val("&H" & avntData(lngIndex))
    Next lngIndex

    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Empty variant

    strData = strData & "3A39CE37 D3FAF5CF ABC27737 5AC52D1B 5CB0679E 4FA33742 D3822740 99BC9BBE "
    strData = strData & "D5118E9D BF0F7315 D62D1C7E C700C47B B78C1B6B 21A19045 B26EB1BE 6A366EB4 "
    strData = strData & "5748AB2F BC946E79 C6A376D2 6549C2C8 530FF8EE 468DDE7D D5730A1D 4CD04DC6 "
    strData = strData & "2939BBDB A9BA4650 AC9526E8 BE5EE304 A1FAD5F0 6A2D519A 63EF8CE2 9A86EE22 "
    strData = strData & "C089C2B8 43242EF6 A51E03AA 9CF2D0A4 83C061BA 9BE96A4D 8FE51550 BA645BD6 "
    strData = strData & "2826A2F9 A73A3AE1 4BA99586 EF5562E9 C72FEFD3 F752F7DA 3F046F69 77FA0A59 "
    strData = strData & "80E4A915 87B08601 9B09E6AD 3B3EE593 E990FD5A 9E34D797 2CF0B7D9 022B8B51 "
    strData = strData & "96D5AC3A 017DA67D D1CF3ED6 7C7D2D28 1F9F25CF ADF2B89B 5AD6B472 5A88F54C "
    strData = strData & "E029AC71 E019A5E6 47B0ACFD ED93FA9B E8D3C48D 283B57CC F8D56629 79132E28 "
    strData = strData & "785F0191 ED756055 F7960E44 E3D35E8C 15056DD4 88F46DBA 03A16125 0564F0BD "
    strData = strData & "C3EB9E15 3C9057A2 97271AEC A93A072A 1B3F6D9B 1E6321F5 F59C66FB 26DCF319 "
    strData = strData & "7533D928 B155FDF5 03563482 8ABA3CBB 28517711 C20AD9F8 ABCC5167 CCAD925F "
    strData = strData & "4DE81751 3830DC8E 379D5862 9320F991 EA7A90C2 FB3E7BCE 5121CE64 774FBE32 "
    strData = strData & "A8B6E37E C3293D46 48DE5369 6413E680 A2AE0810 DD6DB224 69852DFD 09072166 "
    strData = strData & "B39A460A 6445C0DD 586CDECF 1C20C8AE 5BBEF7DD 1B588D40 CCD2017F 6BB4E3BB "
    strData = strData & "DDA26A7E 3A59FF45 3E350A44 BCB4CDD5 72EACEA8 FA6484BB 8D6612AE BF3C6F47 "
    strData = strData & "D29BE463 542F5D9E AEC2771B F64E6370 740E0D8D E75B1357 F8721671 AF537D5D "
    strData = strData & "4040CB08 4EB4E2CC 34D2466A 0115AF84 E1B00428 95983A1D 06B89FB4 CE6EA048 "
    strData = strData & "6F3F3B82 3520AB82 011A1D4B 277227F8 611560B1 E7933FDC BB3A792B 344525BD "
    strData = strData & "A08839E1 51CE794B 2F32C9B7 A01FBAC9 E01CC87E BCC7D1F6 CF0111C3 A1E8AAC7 "
    strData = strData & "1A908749 D44FBD9A D0DADECB D50ADA38 0339C32A C6913667 8DF9317C E0B12B4F "
    strData = strData & "F79E59B7 43F5BB3A F2D519FF 27D9459C BF97222C 15E6FC2A 0F91FC71 9B941525 "
    strData = strData & "FAE59361 CEB69CEB C2A86459 12BAA8D1 B6C1075E E3056A0C 10D25065 CB03A442 "
    strData = strData & "E0EC6E0E 1698DB3B 4C98A0BE 3278E964 9F1F9532 E0D392DF D3A0342B 8971F21E "
    strData = strData & "1B0A7441 4BA3348C C5BE7120 C37632D8 DF359F8D 9B992F2E E60B6F47 0FE3F11D "
    strData = strData & "E54CDA54 1EDAD891 CE6279CF CD3E7E6F 1618B166 FD2C1D05 848FD2C5 F6FB2299 "
    strData = strData & "F523F357 A6327623 93A83531 56CCCD02 ACF08162 5A75EBB5 6E163697 88D273CC "
    strData = strData & "DE966292 81B949D0 4C50901B 71C65614 E6C6C7BD 327A140A 45E1D006 C3F27B9A "
    strData = strData & "C9AA53FD 62A80F00 BB25BFE2 35BDD2F6 71126905 B2040222 B6CBCF7C CD769C2B "
    strData = strData & "53113EC0 1640E3D3 38ABBD60 2547ADF0 BA38209C F746CE76 77AFA1C5 20756060 "
    strData = strData & "85CBFE4E 8AE88DD8 7AAAF9B0 4CF9AA7E 1948C25C 02FB8A8C 01C36AE4 D6EBE1F9 "
    strData = strData & "90D4F869 A65CDEA0 3F09252D C208E69F B74E6132 CE77E25B 578FDFE3 3AC372E6 "

    ' Separate into a variant array
    avntData = Split(strData, Chr$(32))

    ' Transfer data to contants array
    For lngIndex = 0 To (MAX_BYTE - 1)
        malngSBox(3, lngIndex) = Val("&H" & avntData(lngIndex))
    Next lngIndex

    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Always empty variants when not needed

End Sub

Private Sub ResetVariables()

    ' Always empty arrays and variants
    ' when not needed.  Saves resources.
    Erase malngPrimaryKey()
    Erase malngSBox()
    
    mcurOrigLength = 0@
    mcurCurrByteCnt = 0@
        
End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objBlowfish = new cBlowfish
                
    ResetVariables
    Erase mabytPassword()
        
    Set mobjMath32 = Nothing
    Set mobjMath32 = New cMath32
    
    ' Set property values
    KeyLength = 32         ' Minimum key length
    EncryptRounds = 1      ' Number of rounds of encryption
    PrimaryKeyRounds = 1   ' Number of rounds to mix primary key array
    CreateNewFile = True   ' Default is to create a new output file

End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objBlowfish = Nothing
            
    Set mobjMath32 = Nothing
    ResetVariables
    Erase mabytPassword()

End Sub
