VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBigFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
'  Module:     clsBigFiles.cls
'
'  Purpose:    This module calculates the size of a file.  Can handle
'              file sizes greater than 2gb.
'
' Reference:   Richard Newcombe  22-Jan-2007
'              Getting Past the 2 Gb File Limit
'              http://www.codeguru.com/vb/controls/vb_file/directory/article.php/c12917__1/
'
'              How To Seek Past VBA's 2GB File Limit
'              http://support.microsoft.com/kb/189981
'
'              How To Write Data to a File Using WriteFile API
'              http://support.microsoft.com/kb/165942/EN-US/
'
'              File Caching
'              http://msdn.microsoft.com/en-us/library/aa364218(VS.85).aspx
'
'              CreateFile Function
'              http://msdn.microsoft.com/en-us/library/aa363858(VS.85).aspx
'
'              INFO: FILE_FLAG_WRITE_THROUGH and FILE_FLAG_NO_BUFFERING
'              http://support.microsoft.com/kb/99794
'
' Description: The descriptions in this module are excerts from Richard
'              Newcombe's article.
'
'              When working in the IDE, any numbers that are entered are
'              limited to a Long variable type. Actually, as far I've
'              found, the IDE uses Longs for most numeric storage within
'              the projects that you write.
'
'              Okay, so what's the problem with Longs? Well, by definition
'              they are a signed 4-byte variable, in hex &H7FFFFFFF, with a
'              lower limit of -2,147,483,648 and an upper limit of
'              2,147,483,647 (2 Gb). &H80000000 stores the sign of the
'              value. Even when you enter values in Hex, they are stored in
'              a Long.
'
'              Working with random access files, you quite often use a Long
'              to store the filesize and current position, completely
'              unaware that if the file you access is just one byte over
'              the 2 Gb size, you can cause your application to corrupt the
'              file when writing to it.
'
'              Unfortunately, there is no quick fix for this. To get around
'              the problem, you need to write your own file handling
'              module, one that uses windows APIs to open, read, write, and
'              close any file.
'
'              The API's expect the Low and High 32-bit values in unsigned
'              format. Also, the APIs return unsigned values. So, the first
'              thing you have to do is decide on a variable type that you
'              can use to store values higher than 2 Gb. After some serious
'              thought, I decided to use a Currency type (64-bit scaled
'              integer) this gives you a 922,337 gig upper file limit, way
'              bigger that the largest hard drive available today.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              http://www.codeguru.com/vb/controls/vb_file/directory/article.php/c12917__1/
' 03-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Documented and modified
' 01-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added flag to OpenReadWrite() operations that will not go
'              through any intermediate cache but will go directly to disk.
'              Thanks to Lorin for pointing this out.
' 15-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced FileExists() routine with IsPathValid() routine.
' 25-Nov-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added boolean flag to parameters.  See CreateBigFile().
' ***************************************************************************
Option Explicit

' ********************************************************************
' Constants
' ********************************************************************
  Private Const MODULE_NAME             As String = "clsBigFiles"
  Private Const FILE_ATTRIBUTE_NORMAL   As Long = &H80
  Private Const FILE_BEGIN              As Long = 0
  Private Const FILE_SHARE_READ         As Long = &H1
  Private Const FILE_SHARE_WRITE        As Long = &H2
  Private Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000
  Private Const GENERIC_READ            As Long = &H80000000
  Private Const GENERIC_WRITE           As Long = &H40000000
  Private Const OPEN_EXISTING           As Long = 3
  Private Const OPEN_ALWAYS             As Long = 4
  Private Const INVALID_HANDLE_VALUE    As Long = -1
  Private Const MAX_LONG                As Long = &H7FFFFFFF        ' 2147483647
  Private Const MAX_DWORD               As Currency = (2# ^ 32) - 1 ' 4294967295

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long

  ' The CreateFile function creates or opens the following objects and
  ' returns a handle that can be used to access the object:
  '
  '       Files
  '       pipes
  '       mailslots
  '       communications resources
  '       disk devices (Windows NT only)
  '       consoles
  '       directories (open only)
  Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
          (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
          ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, _
          ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
          ByVal hTemplateFile As Long) As Long

  ' CloseHandle invalidates the specified object handle, decrements
  ' the object’s handle count, and performs object retention checks.
  ' Once the last handle to an object is closed, the object is removed
  ' from the operating system. In other words, closes an open file.
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hFile As Long) As Long

  ' This API is used to set the current position in the open file.
  ' This function is very important because any reads or writes to
  ' the file do not automatically forward the file position.
  Private Declare Function SetFilePointer Lib "kernel32" _
          (ByVal hFile As Long, ByVal lDistanceToMove As Long, _
          lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

  ' Use SetEndOfFile to set the new end of the file marker. This
  ' function moves the end-of-file position for the specified file
  ' to the current position of the file pointer.
  Private Declare Function SetEndOfFile Lib "kernel32" _
          (ByVal hFile As Long) As Long

  ' This function writes data to a file, starting at the position
  ' indicated by the file pointer. After the write operation has
  ' been completed, the file pointer is adjusted by the number of
  ' bytes written.
  Private Declare Function WriteFile Lib "kernel32" _
          (ByVal hFile As Long, lpBuffer As Any, _
          ByVal nNumberOfBytesToWrite As Long, _
          lpNumberOfBytesWritten As Long, _
          ByVal lpOverlapped As Any) As Long

  ' This function reads data from a file, starting at the position
  ' indicated by the file pointer. After the read operation has
  ' been completed, the file pointer is adjusted by the number of
  ' bytes read.
  Private Declare Function ReadFile Lib "kernel32" _
          (ByVal hFile As Long, lpBuffer As Any, _
          ByVal nNumberOfBytesToRead As Long, _
          lpNumberOfBytesRead As Long, _
          ByVal lpOverlapped As Any) As Long

  ' GetFileSize determines the size of the file. The file size is
  ' given in a 64-bit value that is split into two 32-bit values.
  ' The high-order half is put into the variable passed as
  ' lpFileSizeHigh; the low-order half is returned by the function.
  ' To get the size, you can either put the binary or hexadecimal
  ' values of the two variables side-by-side, or use the formula
  ' filesize = lpFileSizeHigh * 2^32 + return value. If an error
  ' occurs, the function instead returns -1.
  Private Declare Function GetFileSize Lib "kernel32.dll" _
          (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

  ' Forces Windows to flush the write cache.
  Private Declare Function FlushFileBuffers Lib "kernel32" _
          (ByVal hFile As Long) As Long

  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Private Declare Function SetFileAttributes Lib "kernel32" _
          Alias "SetFileAttributesA" _
          (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long


' ***************************************************************************
' ***                           Methods                                   ***
' ***************************************************************************

' ***************************************************************************
' Routine:       OpenReadOnly
'
' Description:   Open a file to be used as input.  The file must already
'                exist.  If the file does not exist, an error will occur.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName - Fully qualified path and file name
'                hFile - Numeric value designating an open file
'
' Returns:       TRUE - Successfully opened file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function OpenReadOnly(ByVal strFileName As String, _
                             ByRef hFile As Long) As Boolean

    Const ROUTINE_NAME As String = "OpenReadOnly"

    On Error GoTo OpenReadOnly_Error

    OpenReadOnly = False   ' Preset to bad ending

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' make sure the file does exist
    If Not IsPathValid(strFileName) Then
        Exit Function
    End If

    ' Open a file that is to be used as input (read only).
    hFile = CreateFile(strFileName & Chr$(0), _
                       GENERIC_READ, FILE_SHARE_READ, _
                       0&, OPEN_EXISTING, 0&, 0&)

    ' error opening the file
    If hFile = INVALID_HANDLE_VALUE Then
        Exit Function
    End If

    ' Set the pointer to start at the beginning of the file
    SetFilePointer hFile, 0, 0, FILE_BEGIN

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        API_CloseFile hFile     ' Verify file handle has been released
    Else
        OpenReadOnly = True
    End If

OpenReadOnly_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

OpenReadOnly_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    API_CloseFile hFile     ' Verify file handle has been released
    OpenReadOnly = False
    Resume OpenReadOnly_CleanUp

End Function

' ***************************************************************************
' Routine:       OpenReadWrite
'
' Description:   Open a file to update.  If the file exist, if will be
'                opened.  If the file does not exist, it will be created.
'                Use carefully.  If you open an existing file and something
'                goes wrong, the file may become a zero byte file.  There
'                is no recovery of the data available.  I use this to access
'                a temporary work file only.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName - Fully qualified path and file name
'                hFile - Numeric value designating an open file
'
' Returns:       TRUE - Successfully opened file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added flag when the file is opened that will allow data to not
'              go through any intermediate cache but will go directly to disk.
'              Thanks to Lorin for pointing this out.
' ***************************************************************************
Public Function OpenReadWrite(ByVal strFileName As String, _
                              ByRef hFile As Long) As Boolean

    Const ROUTINE_NAME As String = "OpenReadWrite"

    On Error GoTo OpenReadWrite_Error

    OpenReadWrite = False  ' Preset to bad ending

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Open a file that is to be updated.
    '
    ' The flag FILE_FLAG_WRITE_THROUGH is specified so that
    ' system caching is not in effect, then the data is
    ' immediately flushed to disk without going through the
    ' system cache. The operating system also requests a
    ' write-through of the hard disk cache to persistent media.
    '
    ' Note:  Not all hardware supports this write-through capability
    '        with both flags being used at the same time.
    '        FILE_FLAG_WRITE_THROUGH and FILE_FLAG_NO_BUFFERING
    '
    hFile = CreateFile(strFileName & Chr$(0), _
                       GENERIC_READ Or GENERIC_WRITE, _
                       FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                       0&, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL Or _
                       FILE_FLAG_WRITE_THROUGH, 0&)

    ' error opening the file
    If hFile = INVALID_HANDLE_VALUE Then
        Err.Clear
        Exit Function
    End If

    ' Set the pointer to start at the beginning of the file
    SetFilePointer hFile, 0, 0, FILE_BEGIN

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        API_CloseFile hFile     ' Verify file handle has been released
    Else
        OpenReadWrite = True
    End If

OpenReadWrite_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

OpenReadWrite_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    API_CloseFile hFile     ' Verify file handle has been released
    Resume OpenReadWrite_CleanUp

End Function

' ***************************************************************************
' Routine:       CalcFileSize
'
' Description:   This routine is used to open a file as read only and
'                calculate it's size.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName  - Name of file
'                curFileSize  - Returned file size in bytes
'                strBitsInHex - OPTIONAL - Return the file size calculated
'                               into bits and in hex format without leading
'                               zeroes. Used for calculating a hash.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 15-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated file size calculations.
' ***************************************************************************
Public Sub CalcFileSize(ByVal strFileName As String, _
                        ByRef curFilesize As Currency, _
               Optional ByRef strBitsInHex As String = vbNullString)

    Dim hFile        As Long      ' receives a handle to the file
    Dim lngLowOrder  As Long      ' receive the low-order half of the file size
    Dim lngHighOrder As Long      ' receive the high-order half of the file size
    Dim curHexTemp   As Currency  ' Holding area if greater than 2gb

    Const ROUTINE_NAME As String = "CalcFileSize"

    On Error GoTo CalcFileSize_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    curFilesize = 0@
    lngHighOrder = 0&

    ' Open file as read only and get the file handle
    If OpenReadOnly(strFileName, hFile) Then

        SetFilePointer hFile, 0, 0, FILE_BEGIN              ' Set pointer to beginning of file
        lngLowOrder = GetFileSize(hFile, lngHighOrder)      ' Get API file size values
        API_CloseFile hFile                                 ' Close file handle

        curFilesize = Long2Size(lngLowOrder, lngHighOrder)  ' Calculate actual file size

        '************************
        '* Optional return data *
        '************************
        ' Test for overflow when calculating the
        ' number of bits. This information is
        ' used for generating file hashes.
        If (curFilesize * 8) >= MAX_LONG Then

            ' File size is greater than or equal to 2gb
            curHexTemp = (curFilesize * 8)
            strBitsInHex = NumberToHex(curHexTemp)

        Else
            ' File size is less than 2gb
            strBitsInHex = Hex$(curFilesize * 8)
        End If

    Else
        curFilesize = CCur(INVALID_HANDLE_VALUE)
    End If

CalcFileSize_CleanUp:
    If hFile > 0 Then
        API_CloseFile hFile
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

CalcFileSize_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    curFilesize = 0
    strBitsInHex = vbNullString
    Resume CalcFileSize_CleanUp

End Sub

' ***************************************************************************
' Routine:       API_ReadFile
'
' Description:   This routine is used to read data from an opened file.
'
'                This works much like the standard Read # command. You pass
'                the function the file handle you got when opening the file,
'                the byte position you want to start our read from, the size
'                of the data block you require, and a buffer. For the buffer,
'                you use an array of bytes because this seems to be the best
'                method to ensure that any type of data can be read and
'                written with it.
'
'                You use your conversion function to split the position info
'                into the two 32-bit variables needed for the SetFilePointer
'                API. Set the location in your file and then call your
'                ReadFile API. And, for when you read at the end of the file,
'                you pass the number of bytes that have been read.
'
' Parameters:    hFile - Numeric value designating an open file
'                curPosition - Current position within the file
'                abytData() - Byte array to hold the data that was read
'
' Returns:       TRUE - Successfully read the file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function API_ReadFile(ByVal hFile As Long, _
                             ByVal curPosition As Currency, _
                             ByRef abytData() As Byte) As Boolean

    Dim lngRetCode     As Long
    Dim lngLowOrder    As Long
    Dim lngHighOrder   As Long
    Dim lngBytesRead   As Long
    Dim lngBytesToRead As Long

    Const ROUTINE_NAME As String = "API_ReadFile"

    On Error GoTo API_ReadFile_Error

    API_ReadFile = False   ' Preset to bad ending
    lngLowOrder = 0
    lngHighOrder = 0

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Get size of data to write
    lngBytesToRead = (UBound(abytData) + 1) * LenB(abytData(0))

    ' Calculate current position within file
    Size2Long curPosition, lngLowOrder, lngHighOrder

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo API_ReadFile_CleanUp
    End If

    ' Set the pointer to start at specific position within the file
    SetFilePointer hFile, lngLowOrder, lngHighOrder, FILE_BEGIN

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo API_ReadFile_CleanUp
    End If

    ' Test for successful file pointer
    If lngHighOrder = &HFFFFFFFF Then
        InfoMsg "Failed to set read file pointer." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo API_ReadFile_CleanUp
    End If

    ' Read this portion of the file
    lngRetCode = ReadFile(hFile, abytData(0), lngBytesToRead, lngBytesRead, 0&)

    ' Test for successful file read
    If lngRetCode = 0 Then
        InfoMsg "Failed to read file." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
    Else
        API_ReadFile = True   ' Set flag for successful finish
    End If

API_ReadFile_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        API_CloseFile hFile     ' Verify file handle has been released
        API_ReadFile = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

API_ReadFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    API_CloseFile hFile     ' Verify file handle has been released
    API_ReadFile = False
    Resume API_ReadFile_CleanUp

End Function

' ***************************************************************************
' Routine:       API_WriteFile
'
' Description:   This routine is used to write data to the file.
'
'                This works much like the standard Write # command. You pass
'                the function the file handle you got when opening the file,
'                the byte position you want to start our read from, the size
'                of the data block you require, and a buffer. For the buffer,
'                you use an array of bytes because this seems to be the best
'                method to ensure that any type of data can be read and
'                written with it.
'
'                You use your conversion function to split the position info
'                into the two 32-bit variables needed for the SetFilePointer
'                API. Set the location in your file and then call your
'                ReadFile API. And, for when you read at the end of the file,
'                you pass the number of bytes that have been written.
'
' Parameters:    hFile - Numeric value designating an open file
'                curPosition - Current position within the file
'                abytData() - Byte array that holds the data to be written
'
' Returns:       TRUE - Successfully updated the file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function API_WriteFile(ByVal hFile As Long, _
                              ByVal curPosition As Currency, _
                              ByRef abytData() As Byte) As Boolean

    Dim lngRetCode      As Long
    Dim lngLowOrder     As Long
    Dim lngHighOrder    As Long
    Dim lngBytesToWrite As Long
    Dim lngBytesWritten As Long

    Const ROUTINE_NAME As String = "API_WriteFile"

    On Error GoTo API_WriteFile_Error

    API_WriteFile = False   ' Preset to bad ending
    lngLowOrder = 0
    lngHighOrder = 0

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Get the length of data to write
    lngBytesToWrite = (UBound(abytData) + 1) * LenB(abytData(0))

    ' Calculate current position within file
    Size2Long curPosition, lngLowOrder, lngHighOrder

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo API_WriteFile_CleanUp
    End If

    ' Set the pointer to start at specific position within the file
    SetFilePointer hFile, lngLowOrder, lngHighOrder, FILE_BEGIN

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo API_WriteFile_CleanUp
    End If

    ' Test for successful file pointer
    If lngHighOrder = &HFFFFFFFF Then
        InfoMsg "Failed to set write file pointer." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo API_WriteFile_CleanUp
    End If

    ' Write to the file
    lngRetCode = WriteFile(hFile, abytData(0), lngBytesToWrite, lngBytesWritten, 0&)

    ' Test for successful file read
    If lngRetCode = 0 Then
        InfoMsg "Failed to write to file." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
    Else
        FlushFileBuffers hFile  ' Flush file buffers to force writing
        DoEvents                '     of the data immediately
        API_WriteFile = True    ' Set flag for successful finish
    End If

API_WriteFile_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        API_CloseFile hFile     ' Verify file handle has been released
        API_WriteFile = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

API_WriteFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    API_CloseFile hFile     ' Verify file handle has been released
    API_WriteFile = False
    Resume API_WriteFile_CleanUp

End Function

' ***************************************************************************
' Routine:       API_SetEndOfFile
'
' Description:   Sets the pointer to the end of the file designating that
'                we are now finished with this file.
'
' Parameters:    hFile - Numeric value designating an open file
'                curPosition - Current position within the file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub API_SetEndOfFile(ByVal hFile As Long, _
                            ByVal curPosition As Currency)

    Dim lngRetCode   As Long
    Dim lngLowOrder  As Long
    Dim lngHighOrder As Long

    Const ROUTINE_NAME As String = "API_SetEndOfFile"

    On Error GoTo API_SetEndOfFile_Error

    lngLowOrder = 0
    lngHighOrder = 0

    ' Calculate current position within file
    Size2Long curPosition, lngLowOrder, lngHighOrder

    ' Set the pointer to the end of the file
    lngRetCode = SetFilePointer(hFile, lngLowOrder, lngHighOrder, FILE_BEGIN)

    ' Test for successful file pointer
    If lngHighOrder <> &HFFFFFFFF Then
        ' Set end of file marker
        lngRetCode = SetEndOfFile(hFile)
    Else
        InfoMsg "Failed to set end of file pointer." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If

API_SetEndOfFile_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

API_SetEndOfFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume API_SetEndOfFile_CleanUp

End Sub

' ***************************************************************************
' Routine:       API_CloseFile
'
' Description:   Closes an open file.
'
' Parameters:    hFile - Numeric value designating an open file
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub API_CloseFile(ByRef hFile As Long)

    ' Always close a file when not in use or
    ' undesired consequences may happen.
    If hFile > 0 Then
        CloseHandle hFile  ' Release file handle
        hFile = 0          ' Reset handle value
    End If

End Sub

' **************************************************************************
' Routine:       CreateBigFile
'
' Description:   Creates a file from 1 byte to greater than 2gb filled with
'                null values. I have created files greater than 5gb without
'                any problems.
'
' Parameters:    strFilename - Name of file to be created
'                curFileSize - Size of new file
'                blnVerify   - OPTIONAL - Verify if file was successfully
'                              created and matches the size requested.
'                              Default = TRUE
'
' Returns:       TRUE - File was successfully created
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Original routine
' 25-Nov-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added boolean flag to parameters
' ***************************************************************************
Public Function CreateBigFile(ByVal strFileName As String, _
                              ByVal curFilesize As Currency, _
                     Optional ByVal blnVerify As Boolean = True) As Boolean

    Dim hFile      As Long
    Dim blnGood    As Boolean
    Dim curNewSize As Currency

    Const ROUTINE_NAME As String = "CreateBigFile"

    On Error GoTo CreateBigFile_Error

    blnGood = False

    ' If a file by this name already exist,
    ' then make sure it is empty.
    DoEvents
    If IsPathValid(strFileName) Then

        ' Make sure file attributes are normal
        DoEvents
        SetFileAttributes strFileName, FILE_ATTRIBUTE_NORMAL

        ' Empty the file
        hFile = FreeFile
        Open strFileName For Output As #hFile
        Close #hFile

    End If

    ' Open the file with read/write access
    If OpenReadWrite(strFileName, hFile) Then

        ' Load file with null values
        blnGood = LoadWithNullValues(hFile, curFilesize)

        ' Flush file buffers to force
        ' writing of the data immediately
        FlushFileBuffers hFile
        DoEvents

        ' Always release file handle
        ' when no longer needed.
        API_CloseFile hFile

        ' See if file was successsfully
        ' created and loaded with nulls
        If blnGood Then

            If blnVerify Then
                ' Calculate the size of the new file to
                ' see if it matches what was requested
                CalcFileSize strFileName, curNewSize

                ' Compare new file size to requested size
                If curNewSize <> curFilesize Then
                    blnGood = False   ' File sizes do not match
                End If
            End If
        End If

    End If

CreateBigFile_CleanUp:
    CreateBigFile = blnGood   ' Set flag based on type of completion
    API_CloseFile hFile       ' Verify file handle has been released

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CreateBigFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    blnGood = False
    Resume CreateBigFile_CleanUp

End Function

' **************************************************************************
' Routine:       LoadWithNullValues
'
' Description:   Updates a file from 1 byte to greater than 2gb with null
'                values.
'
' Parameters:    hFile - Numeric designator of file to be updated
'                curFileSize - Amount of data to be written to file
'
' Returns:       TRUE - File was successfully loaded
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Original routine
' ***************************************************************************
Public Function LoadWithNullValues(ByVal hFile As Long, _
                                   ByVal curFilesize As Currency) As Boolean

    Dim lngDistanceToMove     As Long
    Dim lngDistanceToMoveHigh As Long

    LoadWithNullValues = False  ' Preset to bad ending

    ' If no file handle then leave
    If hFile < 1 Then
        Exit Function
    End If

    ' Calculate size of new file
    Size2Long curFilesize, lngDistanceToMove, lngDistanceToMoveHigh

    ' Set the file pointers to mark the size of the file
    If SetFilePointer(hFile, lngDistanceToMove, _
                      lngDistanceToMoveHigh, FILE_BEGIN) > 0 Then

        ' Test to see if we can identify
        ' the end of the new file
        If SetEndOfFile(hFile) Then
            LoadWithNullValues = True  ' Successful finish
        End If
    End If

End Function


' ***************************************************************************
' ****                Internal Procedures and Functions                  ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Size2Long
'
' Description:   This routine will work out the higher 32 bits. This code
'                looks like it could be done with a simple division, but
'                you have the problem of the IDE using longs. So, rather
'                than running the risk of the IDE using a long somewhere
'                in the calculations (tests have returned varied results
'                where the division is often 1 out on certain file sizes),
'                you may find a division method that will give the correct
'                value every time, but for now this method will suffice.
'
'                With your max filesize being 922,337 Gb, the highest
'                value that LongHigh will hold is 2,147,630, way below the
'                maximum positive value that a Long can hold. This means
'                that you do not need to monitor it.
'
' Parameters:    curFileSize - File size to be evaluated
'                lngLowOrder - The highest value this will hold is 2,147,483,647
'                lngHighOrder - The highest value this will hold is 2,147,630
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub Size2Long(ByVal curFilesize As Currency, _
                      ByRef lngLowOrder As Long, _
                      ByRef lngHighOrder As Long)

    ' Called by API_ReadFile()
    '           API_WriteFile()
    '           API_SetEndOfFile()
    '           LoadWithNullValues()

    lngHighOrder = 0   ' Init return values
    lngLowOrder = 0

    ' If the filesize is smaller than 4 Gb, the Do loop does not execute.
    ' Between 4 and 8 Gb the Do loop executes once.  Between 8-12 Gb, it
    ' executes twice.
    Do Until curFilesize < MAX_DWORD
        lngHighOrder = lngHighOrder + 1
        curFilesize = curFilesize - MAX_DWORD
    Loop

    ' Take the remainder and decide whether it needs the sign bit of the
    ' long to hold a value (&H80000000). If it's smaller than 2,147,483,648
    ' you can pass the value directly to the low order. If not, you have to
    ' convert the value to the hex equivalent in signed 32bit.
    If curFilesize > MAX_LONG Then
        ' Larger than 2gb
        lngLowOrder = CLng((MAX_DWORD - (curFilesize - 1)) * -1)
    Else
        ' Less than 2gb
        lngLowOrder = CLng(curFilesize)
    End If

End Sub

' ***************************************************************************
' Routine:       Long2Size
'
' Description:   This routine will convert two Long values into one
'                Currency value.  The multiplication here does not
'                give any erroneous results, but note that the currency
'                type variable must be listed first; otherwise, VB will
'                try to use a Long type variable to temporally store
'                the result. (This is one of the minor issues in VB6.
'                When doing calculations, VB uses the same variable
'                type of your first variable in the calculation and not
'                the variable type of the destination variable. This
'                has been documented on several sites.)
'
' Parameters:    lngLowOrder  - Highest value = 2,147,483,647
'                lngHighOrder - Highest value = 2,147,630
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Renamed, modified and documented
' ***************************************************************************
Private Function Long2Size(ByVal lngLowOrder As Long, _
                           ByVal lngHighOrder As Long) As Currency

    ' Called by CalcFileSize()

    Dim curFilesize As Currency

    curFilesize = MAX_DWORD * CCur(lngHighOrder)

    If lngLowOrder < 0 Then
        Long2Size = curFilesize + (MAX_DWORD + CCur(lngLowOrder + 1))
    Else
        Long2Size = curFilesize + CCur(lngLowOrder)
    End If

End Function

' ***************************************************************************
' Routine:       NumberToHex
'
' Description:   Convert a whole number to a hex string
'
' Parameters:    vntNumber - Number to be converted
'
' Returns:       Hex string
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Function NumberToHex(ByVal vntNumber As Variant) As String

    ' Called by CalcFileSize()

    Dim vntHex As Variant
    Dim strHex As String

    vntHex = Empty  ' Always start with an empty variant
    strHex = vbNullString
    vntNumber = CDec(vntNumber)

    If vntNumber < 1 Then
        GoTo NumberToHex_CleanUp
    End If

    Do While vntNumber > 0

        vntHex = Fix(vntNumber / 16)
        vntHex = vntNumber - (vntHex * 16)
        vntNumber = Int(vntNumber / 16)

        Select Case vntHex
               Case 10:   strHex = "A" & strHex
               Case 11:   strHex = "B" & strHex
               Case 12:   strHex = "C" & strHex
               Case 13:   strHex = "D" & strHex
               Case 14:   strHex = "E" & strHex
               Case 15:   strHex = "F" & strHex
               Case Else: strHex = CStr(vntHex) & strHex
        End Select

    Loop

NumberToHex_CleanUp:
    NumberToHex = strHex  ' Return hex string
    vntHex = Empty        ' Always empty variants when not needed
    strHex = vbNullString

End Function

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))

End Function



