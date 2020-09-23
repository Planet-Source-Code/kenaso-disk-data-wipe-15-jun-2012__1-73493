Attribute VB_Name = "modCommon"
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
Option Explicit

' ***************************************************************************
' Global Constants
' ***************************************************************************
  Public Const PGM_NAME                    As String = "kiWipe DLL"
  Public Const DLL_NAME                    As String = "kiWipe"
  Public Const ENCRYPT_EXT                 As String = ".ENC"
  Public Const DECRYPT_EXT                 As String = ".DEC"
  Public Const TMP_FILE_PREFIX             As String = "~ki"   ' User defined prefix
  Public Const FILE_ATTRIBUTE_NORMAL       As Long = &H80&
  Public Const MOVEFILE_REPLACE_EXISTING   As Long = &H1&
  Public Const MOVEFILE_COPY_ALLOWED       As Long = &H2&
  Public Const MIN_PWD_LENGTH              As Long = 5    ' Minimum password length
  Public Const MAX_PWD_LENGTH              As Long = 50   ' Maximum password length

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "modCommon"
  Private Const GB_4        As Currency = 4294967296@
  Private Const KB_1        As Long = &H400&       ' 1024
  Private Const KB_32       As Long = &H8000&      ' 32768
  Private Const KB_64       As Long = &H10000      ' 65536
  Private Const KB_128      As Long = &H20000      ' 131072
  Private Const KB_256      As Long = &H40000      ' 262144
  Private Const KB_512      As Long = &H80000      ' 524288
  Private Const MB_1        As Long = &H100000     ' 1048576
  Private Const MAX_BYTE    As Long = 256
  Private Const MAX_SIZE    As Long = 260
  Private Const MAX_LONG    As Long = &H7FFFFFFF   ' 2147483647
  
' ***************************************************************************
' Type Structures
' ***************************************************************************
  ' Used for determining memory size
  Private Type LARGE_INTEGER
      LowPart  As Long
      HighPart As Long
  End Type

  ' Used for determining memory size
  Private Type MEMORYSTATUSEX
      dwLength                As Long
      dwMemoryLoad            As Long
      ullTotalPhys            As LARGE_INTEGER
      ullAvailPhys            As LARGE_INTEGER
      ullTotalPageFile        As LARGE_INTEGER
      ullAvailPageFile        As LARGE_INTEGER
      ullTotalVirtual         As LARGE_INTEGER
      ullAvailVirtual         As LARGE_INTEGER
      ullAvailExtendedVirtual As LARGE_INTEGER
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long
  
  ' The GetTempPath function retrieves the path of the directory designated
  ' for temporary files.  The GetTempPath function gets the temporary file
  ' path as follows:
  '   1.  The path specified by the TMP environment variable.
  '   2.  The path specified by the TEMP environment variable, if TMP
  '       is not defined.
  '   3.  The current directory, if both TMP and TEMP are not defined.
  Private Declare Function GetTempPath Lib "kernel32.dll" _
          Alias "GetTempPathA" (ByVal nBufferLength As Long, _
          ByVal lpBuffer As String) As Long

  ' The GetTempFileName function creates a name for a temporary file.
  ' The filename is the concatenation of specified path and prefix strings,
  ' a hex string formed from a specified integer, and the .TMP extension.
  Private Declare Function GetTempFileName Lib "kernel32.dll" _
          Alias "GetTempFileNameA" (ByVal lpszPath As String, _
          ByVal lpPrefixString As String, ByVal wUnique As Long, _
          ByVal lpTempFileName As String) As Long

  ' Parses a path to determine if it is a directory root. Returns TRUE
  ' if the specified path is a root, or FALSE otherwise.
  Private Declare Function PathIsRoot Lib "shlwapi" Alias "PathIsRootA" _
          (ByVal pszPath As String) As Long
  
  ' Truncates a path to fit within a certain number of characters by
  ' replacing path components with elipses.  Called by ShrinkTofit().
  Private Declare Function PathCompactPathEx Lib "shlwapi" Alias "PathCompactPathExA" _
          (ByVal pszOut As String, ByVal pszSrc As String, _
          ByVal cchMax As Long, ByVal dwFlags As Long) As Long
 
  ' Retrieves information about the system's current usage of both physical
  ' and virtual memory.  You can use the GlobalMemoryStatusEx function to
  ' determine how much memory your application can allocate without severely
  ' impacting other applications.  The information returned by the
  ' GlobalMemoryStatusEx function is volatile.  There is no guarantee that
  ' two sequential calls to this function will return the same information.
  Private Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" _
          (ByRef lpBuffer As MEMORYSTATUSEX) As Long

' ***************************************************************************
' Global API Declares
' ***************************************************************************
  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Public Declare Function GetTickCount Lib "kernel32" () As Long
              
  ' ZeroMemory is used for clearing contents of a type structure.
  Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
         (Destination As Any, ByVal Length As Long)

  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
         (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

  ' MoveFileEx Function moves an existing file or directory, including its
  ' children, with various move options.  If successful then return code is
  ' nonzero.
  Public Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" _
         (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
         ByVal dwFlags As Long) As Long

' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Currency)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m cur Memory
' Variable name:     mcurMemory
'
' ***************************************************************************
  Private mcurMemory As Currency    ' Total physical or available memory


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Function QualifyPath(ByVal strPath As String, _
                   Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If
    
End Function

' ***************************************************************************
' Routine:       UnQualifyPath
'
' Description:   Removes a trailing character from the path
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to remove
'                          Default = "\"
'
' Returns:       Fully qualified path without a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Function UnQualifyPath(ByVal strPath As String, _
                     Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        UnQualifyPath = Left$(strPath, Len(strPath) - 1)
    Else
        UnQualifyPath = strPath
    End If
    
End Function

' ***************************************************************************
' Routine:       GetPath
'
' Description:   Capture complete path up to filename.  Path must end with
'                a backslash.
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Complete path to last backslash
'
' Example:       "C:\Kens Software" = "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetPath(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetPath = objFSO.GetParentFolderName(strPathFile)
    Set objFSO = Nothing
    
End Function

' ***************************************************************************
' Routine:       GetFilename
'
' Description:   Capture file name
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Just the file name
'
' Example:       "Gif89.dll" = "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetFilename(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetFilename = objFSO.GetFilename(strPathFile)
    Set objFSO = Nothing
    
End Function

' ***************************************************************************
' Routine:       GetFilenameExt
'
' Description:   Capture file name extension
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       File name extension
'
' Example:       "dll" = "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetFilenameExt(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetFilenameExt = objFSO.GetExtensionName(strPathFile)
    Set objFSO = Nothing
    
End Function

' ***************************************************************************
' Routine:       GetVersion
'
' Description:   Capture file version information
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Version information
'
' Example:       "1.0.0.1" = "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetVersion(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetVersion = objFSO.GetFileVersion(strPathFile)
    Set objFSO = Nothing
    
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
Public Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))
   
End Function
  
' ***************************************************************************
' Routine:       CloseAllFiles
'
' Description:   Closes any files that were opened within this application.
'                The FreeFile() function returns an integer representing the
'                next file handle opened by this application.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function CloseAllFiles() As Boolean

    While FreeFile > 1
        Close #FreeFile - 1
    Wend
    
End Function

' **************************************************************************
' Routine:       DisplayNumber
'
' Description:   Return a string representing the value in string format
'                to requested number of decimal positions.
'
'                    Bytes  Bytes
'                    KB     Kilobytes
'                    MB     Megabytes
'                    GB     Gigabytes
'                    TB     Terabytes
'                    PB     Petabytes
'
'                Ex:  2,530,096 bytes --> 2.4 MB
'
' Parameters:    dblCapacity - value to be reformatted
'                lngDecimals - [OPTIONAL] number of decimal positions.
'                     Valid values are 0-5.  Change to meet special needs.
'                     Default value = 1 decimal position
'
' Returns:       Reformatted string representation
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-DEC-2001  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 12-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated output
' ***************************************************************************
Public Function DisplayNumber(ByVal dblCapacity As Double, _
                     Optional ByVal lngDecimals As Long = 1) As String
  
    ' Called by GetDriveInfo()
    
    Dim intCount As Long
    Dim dblValue As Double
    
    Const KB_1 As Double = 1024#
    
    On Error GoTo DisplayNumber_Error
    
    dblValue = dblCapacity   ' I do this for debugging purposses
    intCount = 0
    DisplayNumber = vbNullString
    
    If dblValue > 0 Then
        
        ' Must be a positive value
        If lngDecimals < 0 Then
            lngDecimals = 0
        End If
        
        ' Maximum of 5 decimal positions.
        If lngDecimals > 5 Then
            lngDecimals = 5      ' Change to meet special needs.
        End If
    
        ' loop thru the value and determine how
        ' many times it can be divided by 1kb
        Do While dblValue > (KB_1 - 1)
            dblValue = dblValue / KB_1
            intCount = intCount + 1
        Loop
        
        If lngDecimals = 0 Then
            ' Format value with no decimal positions
            DisplayNumber = Format$(Fix(dblValue), "0")
        Else
            ' Format value with requested decimal positions
            DisplayNumber = FormatNumber(dblValue, lngDecimals)
        End If
        
        ' Append extension (Ex:  "70.1 GB")
        DisplayNumber = DisplayNumber & " " & _
                        Choose(intCount + 1, "Bytes", "KB", "MB", "GB", "TB", "PB")
    Else
    
        ' No value was passed to this routine
        If lngDecimals = 0 Then
            ' Format value with no decimal positions
            DisplayNumber = "0 Bytes"
        Else
            ' Format value with one decimal position
            DisplayNumber = "0.0 Bytes"
        End If
    
    End If

DisplayNumber_Error:

End Function

' ***************************************************************************
' Procedure:     GetLongName
'
' Description:   The Dir() function can be used to return a long filename
'                but it does not include path information. By parsing a
'                given short path/filename into its constituent directories,
'                you can use the Dir() function to build a long path/filename.
'
' Example:       Syntax:
'                   GetLongName C:\DOCUME~1\KENASO\LOCALS~1\Temp\~ki6A.tmp
'
'                Returns:
'                   "C:\Documents and Settings\Kenaso\Local Settings\Temp\~ki6A.tmp"
'
' Parameters:    strShortName - Path or file name to be converted.
'
' Returns:       A readable path or file name.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jul-2004  http://support.microsoft.com/kb/154822
'              "How To Get a Long Filename from a Short Filename"
' 09-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function GetLongName(ByVal strShortName As String) As String

    Dim strTemp     As String
    Dim strLongName As String
    Dim intPosition As Integer
    
    On Error Resume Next
    
    strLongName = vbNullString
    
    ' Remove all double quotes
    strShortName = Replace(strShortName, Chr$(34), "")
    
    ' Add a backslash to short name, if needed,
    ' to prevent Instr() function from failing.
    strShortName = QualifyPath(strShortName)
    
    ' Start at position 4 so as to ignore
    ' "[Drive Letter]:\" characters.
    intPosition = InStr(4, strShortName, "\")
    
    ' Pull out each string between
    ' backslash character for conversion.
    Do While intPosition > 0
    
        ' Progressively parse path to verify
        ' each portion does exist and
        ' capture its expanded version.
        strTemp = Dir$(Left$(strShortName, intPosition - 1), _
                       vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
        
        ' Cannot read path
        If Len(strTemp) = 0 Then
            strShortName = vbNullString  ' Empty variables
            strLongName = vbNullString
            Exit Do            ' Exit this loop
        End If
        
        ' Append new elongated portion to output string
        ' after converting it to propercase format.
        strLongName = strLongName & "\" & StrConv(strTemp, vbProperCase)
        
        ' Find next backslash
        intPosition = InStr(intPosition + 1, strShortName, "\")
    
    Loop
    
    ' Prefix with drive letter and colon
    GetLongName = UCase$(Left$(strShortName, 2)) & strLongName

    On Error GoTo 0   ' Nullify this error trap
    
End Function

' ***************************************************************************
' Routine:       IsPathARoot
'
' Description:   Parses a path to determine if it is a directory root.
'                The passed path does not have to exist.
'
'                Returns True for paths such as "\", "X:\", "\\server\share",
'                or "\\server\", or for paths that begin with those strings.
'                Paths such as "..\path2" will return False.
'
' Parameters:    strPath - full path/folder name
'
' Returns:       True if the specified path is a root, or False otherwise.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 03-DEC-2001  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function IsPathARoot(ByVal strPath As String) As Boolean
    
    IsPathARoot = CBool(PathIsRoot(strPath))
    
End Function
    
' ***************************************************************************
' Routine:       UnsignedToLong
'
' Description:   This function takes a Double containing a value in the
'                range of an unsigned Long and returns a Long Integer.
'
' Parameters:    dblValue - Number to be converted
'
' Returns:       Positive long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewote routine
' ***************************************************************************
Public Function UnsignedToLong(ByVal dblValue As Double) As Long

    Do
        Do While dblValue > MAX_LONG
            dblValue = dblValue - GB_4
        Loop
        
        Do While dblValue < 0
            dblValue = dblValue + MAX_LONG
        Loop
    
    Loop Until (dblValue > 0) And (dblValue <= MAX_LONG)
    
    UnsignedToLong = CLng(dblValue)
    
End Function

' ***************************************************************************
' Routine:       ShrinkToFit
'
' Description:   This routine creates the ellipsed string by specifying
'                the size of the desired string in characters.  Adds
'                ellipses to a file path whose maximum length is specified
'                in characters.
'
' Parameters:    strPath - Path to be resized for display
'                intMaxLength - Maximum length of the return string
'
' Returns:       Resized path
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-May-2004  Randy Birch
'              http://vbnet.mvps.org/code/fileapi/pathcompactpathex.htm
' 22-Jun-2004  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Function ShrinkToFit(ByVal strPath As String, _
                            ByVal intMaxLength As Integer) As String

    Dim strBuffer As String
    
    strPath = TrimStr(strPath)
    
    ' See if ellipses need to be inserted into the path
    If Len(strPath) <= intMaxLength Then
        ShrinkToFit = strPath
        Exit Function
    End If
    
    ' intMaxLength is the maximum number of characters to be contained in the
    ' new string, **including the terminating NULL character**. For example,
    ' if intMaxLength = 8, the resulting string would contain a maximum of
    ' seven characters plus the termnating null.
    '
    ' Because of this, add 1 to the value passed as intMaxLength to ensure
    ' the resulting string is the size requested.
    intMaxLength = intMaxLength + 1
    strBuffer = Space$(MAX_SIZE)
    PathCompactPathEx strBuffer, strPath, intMaxLength, 0&
    
    ' Return the readjusted data string
    ShrinkToFit = TrimStr(strBuffer)
    
End Function

' ***************************************************************************
' Routine:       EmptyCollection
'
' Description:   Properly empty and deactivate a collection
'
' Parameters:    colData - Collection to be processed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub EmptyCollection(ByRef colData As Collection)

    ' Has collection been deactivated?
    If colData Is Nothing Then
        Exit Sub
    End If
    
    ' Is the collection empty?
    Do While colData.Count > 0
        
        ' Parse backwards thru collection and delete data.
        ' Backwards parsing prevents a collection from
        ' having to reindex itself after each data removal.
        colData.Remove colData.Count
    Loop
    
    ' Free collection object from memory
    Set colData = Nothing
    
End Sub

' ***************************************************************************
' Routine:       CreateTempFile
'
' Description:   System generated temporary file. The file will be located
'                in the Windows default temporary work folder.
'
' Returns:       path\name of temporary file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function CreateTempFile() As String

    Dim strTmpFile   As String
    Dim strTmpFolder As String

    strTmpFile = Space$(MAX_SIZE)   ' preload with spaces, not nulls
    strTmpFolder = GetTempFolder()  ' Find Windows default temp folder

    ' Create a unique temporary file name consisting of my defined
    ' prefix+system generated hex string+TMP extension (ex: ~kiC151.TMP).
    ' Concatenate path and file name.  Create an empty temp file.
    ' (ex:  C:\Windows\Temp\~kiC151.TMP)
    GetTempFileName strTmpFolder, TMP_FILE_PREFIX, 0&, strTmpFile
    DoEvents
    
    ' Remove any trailing null values
    strTmpFile = TrimStr(strTmpFile)

    ' Convert data string to a readable format and in propercase
    strTmpFile = GetLongName(strTmpFile)
    
    CreateTempFile = strTmpFile  ' Return complete path\filename
    
End Function

' ***************************************************************************
' Routine:       GetTempFolder
'
' Description:   Find system generated temporary folder.
'
' Returns:       Path to windows default temp folder
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetTempFolder() As String

    Dim strTmpFolder As String

    strTmpFolder = Space$(MAX_SIZE)  ' preload with spaces, not nulls

    If GetTempPath(MAX_SIZE, strTmpFolder) <> 0 Then

        ' Found Windows default Temp folder. Remove
        ' any trialing nulls and double quotes.
        strTmpFolder = TrimStr(strTmpFolder)
        strTmpFolder = UnQualifyPath(strTmpFolder)
        strTmpFolder = Replace(strTmpFolder, Chr$(34), "")
        
        If IsPathValid(strTmpFolder) Then
            strTmpFolder = QualifyPath(strTmpFolder)   ' Append backslash
        Else
            strTmpFolder = "C:\"  ' should never happen
        End If
    Else
        ' Did not find Windows default temp folder
        ' therefore, use root level of drive C:
        strTmpFolder = "C:\"  ' should never happen
    End If

    GetTempFolder = strTmpFolder  ' Return path name

End Function

' ***************************************************************************
' Routine:       IsArrayInitialized
'
' Description:   This is an ArrPtr function that determines if the passed
'                array is initialized, and if so will return the pointer
'                to the safearray header. If the array is not initialized,
'                it will return zero. Normally you need to declare a VarPtr
'                alias into msvbvm50.dll or msvbvm60.dll depending on the
'                VB version, but this function will work with vb5 or vb6.
'                It is handy to test if the array is initialized as the
'                return value is non-zero.  Use CBool to convert the return
'                value into a boolean value.
'
'                This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
'                ex:  If CBool(IsArrayInitialized(array_being_tested)) Then ...
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Mar-2008  RD Edwards
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69970
' ***************************************************************************
Public Function IsArrayInitialized(ByVal avntData As Variant) As Long

    Dim intDataType As Integer   ' Must be an integer

    On Error GoTo IsArrayInitialized_Exit
    
    IsArrayInitialized = 0  ' preset to FALSE
    
    ' Get the real VarType of the argument, this is similar
    ' to VarType(), but returns also the VT_BYREF bit
    CopyMemory intDataType, avntData, 2&

    ' if a valid array was passed
    If (intDataType And vbArray) = vbArray Then
        
        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena and Monte Hansen.
        CopyMemory IsArrayInitialized, ByVal VarPtr(avntData) + 8&, 4&
    
    End If
    
IsArrayInitialized_Exit:
    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       MixAppendedData
'
' Description:   Performs simple Encryption/Decryption on the information
'                that is to be appended to the original data after normal
'                encryption.  By mixing the appended data you are keeping
'                prying eyes from knowing required information needed to
'                perform decryption easily.  When calling this routine
'                while performing decryption, the data will be decrypted.
'
' Parameters:    abytData() - Byte array to be encrypted/decrypted
'                lngMixCount - Optional - Number of passes to mix the data
'                        Default = 5
'
' Returns:       Return data in a byte array.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Simplified mixing process
' 01-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Mix count is now an optional value
'              - Updated documentation
' ***************************************************************************
Public Sub MixAppendedData(ByRef abytData() As Byte, _
                  Optional ByVal lngMixCount As Long = 5)

    Dim lngHigh    As Long
    Dim lngStep    As Long
    Dim lngLoop    As Long
    Dim lngIndex   As Long
    Dim abytTemp() As Byte
    
    ReDim abytTemp(MAX_BYTE)  ' Size temp array
    
    ' Verify number of mixing loops
    ' are within an acceptable range
    Select Case lngMixCount
           Case Is < 2:  lngMixCount = 2   ' Set to minimum
           Case Is > 10: lngMixCount = 10  ' set to maximum
    End Select
    
    lngHigh = UBound(abytData)
    lngStep = (lngHigh + lngMixCount) Mod MAX_BYTE
    
    ' Load with ASCII decimal values (0-255)
    For lngIndex = 0 To (MAX_BYTE - 1)
        abytTemp(lngIndex) = CByte(lngIndex)
    Next lngIndex
        
    ' Extra looping for additional security
    For lngLoop = 1 To lngMixCount
        
        ' Perform simple encryption/decryption using Xor
        For lngIndex = 0 To lngHigh
            abytData(lngIndex) = abytData(lngIndex) Xor abytTemp((lngStep + lngIndex) Mod MAX_BYTE)
        Next lngIndex
        
    Next lngLoop
    
    Erase abytTemp()   ' Always empty array when not needed
    
End Sub

' ***************************************************************************
' Routine:       ExpandData
'
' Description:   Expand byte array to a designated length.
'
' Parameters:    abytInput() - Incoming byte array
'                lngReturnLen - Output length of return byte array
'
' Returns:       Expanded byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function ExpandData(ByRef abytInput() As Byte, _
                           ByVal lngReturnLen As Long) As Byte()

    Dim lngIndex     As Long
    Dim lngStart     As Long
    Dim lngTmpIdx    As Long
    Dim lngInputLen  As Long
    Dim abytTemp()   As Byte
    Dim abytOutput() As Byte

    Const ROUTINE_NAME As String = "ExpandData"

    On Error GoTo ExpandData_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo ExpandData_CleanUp
    End If

    ReDim abytOutput(lngReturnLen)   ' Resize output array
    lngInputLen = UBound(abytInput)  ' Capture length of input array

    ' Load output array
    For lngIndex = 0 To lngInputLen - 1
        
        ' Copy data from input array to output array
        abytOutput(lngIndex) = abytInput(lngIndex)
        
        ' If there is more input data than output
        ' array can hold then exit this loop
        If lngIndex = (lngReturnLen - 1) Then
            Exit For
        End If
        
    Next lngIndex
    
    ' If length of incoming data is less than
    ' new output length then append extra data
    ' to output array
    If lngInputLen < lngReturnLen Then

        lngTmpIdx = 0                            ' Init temp array index
        lngStart = lngIndex                      ' Save last output array position
        abytTemp() = LoadXBoxArray(abytInput())  ' Load temp array with 0-255 mixed
        
        ' Load rest of output array
        For lngIndex = lngStart To lngReturnLen - 1
            abytOutput(lngIndex) = abytTemp(lngTmpIdx)  ' Copy temp array to output array
            lngTmpIdx = (lngTmpIdx + 1) Mod MAX_BYTE    ' increment temp array index
        Next lngIndex
                        
    End If
 
    ExpandData = abytOutput()   ' Return expanded data
    
ExpandData_CleanUp:
    Erase abytOutput()  ' Always empty arrays when not needed
    Erase abytTemp()
    On Error GoTo 0     ' Nullify this error trap
    Exit Function

ExpandData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ExpandData_CleanUp
 
End Function
 
' ***************************************************************************
' Routine:       LoadXBoxArray
'
' Description:   The incoming data array (n bytes) is passed to become part
'                of the mixing process. This routine does not duplicate data
'                in the x-Box array (0-255), just rearranges it.  Duplication
'                allows for missing values in the original data.  Be aware of
'                other mixing routines because they may produce duplicate
'                values during the mixing process.  Note that I do not
'                randomly select any data.  The selection process must be
'                repeatable to be able to encrypt\decrypt data.
'
'                WARNING:  If you make any changes to this routine, verify
'                the end results are repeatable.  Remember, this mixing
'                process deals with both encryption and decryption.
'
' Parameters:    abytInput() - Input byte array
'                lngMixCount - [Optional] - number of iterations used for
'                    mixing the data.  Default = 25
'
' Returns:       Byte array contaning mixed ASCII values 0-255 with no
'                duplicates.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function LoadXBoxArray(ByRef abytInput() As Byte, _
                     Optional ByVal lngMixCount As Long = 25) As Byte()

    Dim lngHigh     As Long   ' Number of array elements
    Dim lngLoop     As Long   ' Loop counter
    Dim lngIndex    As Long   ' Loop counter
    Dim lngNewIdx   As Long   ' Calculated index for swapping
    Dim abytMixed() As Byte   ' Array of mixed values 0-255
    Dim abytTemp()  As Byte   ' Holds input data multiple times
    
    Const ROUTINE_NAME As String = "LoadXBoxArray"

    On Error GoTo LoadXBoxArray_Error

    ReDim abytTemp(MAX_BYTE)      ' Size temp array
    ReDim abytMixed(MAX_BYTE)     ' Size output array
    
    lngHigh = UBound(abytInput)   ' Capture size of incoming array
    lngNewIdx = 7                 ' Starting index (my birth day)
    
    ' Verify number of mixing loops
    ' are within an acceptable range
    Select Case lngMixCount
           Case Is < 25: lngMixCount = 25   ' Set to minimum
           Case Is > 99: lngMixCount = 99   ' set to maximum
    End Select
    
    ' Load work arrays
    For lngIndex = 0 To MAX_BYTE - 1
        abytMixed(lngIndex) = CByte(lngIndex)                  ' load ASCII decimal array (0-255)
        abytTemp(lngIndex) = abytInput(lngIndex Mod lngHigh)   ' load array based on input data
    Next lngIndex
            
    ' Outer loop is for obtaining a good mix
    For lngLoop = 1 To lngMixCount
        
        ' Calculate new index (0-255)
        lngNewIdx = (lngNewIdx + abytTemp(lngNewIdx) + abytMixed(lngNewIdx)) Mod MAX_BYTE

        ' Loop thru array and rearrange data
        For lngIndex = 0 To (MAX_BYTE - 1)
        
            ' Calculate new index
            lngNewIdx = (lngNewIdx + abytMixed(lngIndex)) Mod MAX_BYTE

            ' If current index and new index are not
            ' the same then swap data with each other
            If lngIndex <> lngNewIdx Then
                SwapBytes abytMixed(lngIndex), abytMixed(lngNewIdx)
            End If
        
        Next lngIndex
    Next lngLoop
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo LoadXBoxArray_CleanUp
    End If
                     
    LoadXBoxArray = abytMixed()   ' Return mixed data
    
LoadXBoxArray_CleanUp:
    Erase abytMixed()   ' Always empty arrays when not needed
    Erase abytTemp()
        
    On Error GoTo 0     ' Nullify error trap in this routine
    Exit Function

LoadXBoxArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadXBoxArray_CleanUp

End Function
 
' **************************************************************************
' Routine:       CalcProgress
'
' Description:   Calculates current amount of completion
'
' Parameters:    dblCurrAmt   - Current value
'                dblMaxAmount - Maximum value
'
' Returns:       percentage of progression
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function CalcProgress(ByVal dblCurrAmt As Double, _
                             ByVal dblMaxAmount As Double) As Long

    Dim lngPercent As Long

    Const MAX_PERCENT As Long = 100
    
    ' Make sure current value does
    ' not exceed maximum value
    If (dblMaxAmount - dblCurrAmt) < 0 Then
        lngPercent = 0
    Else
        ' Calculate percentage based
        ' on current and maximum value
        lngPercent = Val(Round(dblCurrAmt / dblMaxAmount, 3) * MAX_PERCENT)
    End If
            
    ' Validate percentage so we
    ' do not exceed our bounds
    If lngPercent < 0 Then
        lngPercent = 0
    ElseIf lngPercent > MAX_PERCENT Then
        lngPercent = MAX_PERCENT
    End If

    CalcProgress = lngPercent
    
End Function

' ***************************************************************************
' Routine:       FormatTimeDisplay
'
' Description:   Formats time display
'
' Reference:     Karl E. Peterson, http://vb.mvps.org/
'
' Parameters:    lngMilliseconds - Time in milliseconds
'
' Returns:       Formatted output
'                01:23:45.678  <- 1 hour 23 minutes 45 seconds 678 thousandths
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function FormatTimeDisplay(ByVal lngMilliseconds As Long) As String

    ' Called by CalculateProgress()
    
    Dim lngDays As Long
    
    Const ONE_DAY As Long = 86400000        ' Number of milliseconds in a day
    
    FormatTimeDisplay = vbNullString        ' Verify output string is empty
    lngDays = (lngMilliseconds \ ONE_DAY)   ' Calculate number of days
        
    ' See if one or more days has passed
    If lngDays > 0 Then
        FormatTimeDisplay = CStr(lngDays) & " day(s)  "           ' Start loading output string
        lngMilliseconds = lngMilliseconds - (ONE_DAY * lngDays)   ' Calculate number of milliseconds left
    End If

    ' Continue formatting output string as HH:MM:SS
    FormatTimeDisplay = FormatTimeDisplay & Format$(DateAdd("s", (lngMilliseconds \ 1000), #12:00:00 AM#), "HH:MM:SS")
    
    ' Calc number of milliseconds left
    lngMilliseconds = lngMilliseconds - ((lngMilliseconds \ 1000) * 1000)
    
    ' Append thousandths to output string
    FormatTimeDisplay = FormatTimeDisplay & "." & Format$(lngMilliseconds, "000")
   
End Function

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the record to be written.  The
'                write process has been speeded up by 50% or more by
'                adjusting the record length based on amount of data left
'                to write.
'
' Parameters:    curAmtLeft - Amount of data left to be written
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 13-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added new block size selections based on physical memory
' 18-Oct-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated available physical memory selection
' ***************************************************************************
Public Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Determine record size to write.
    Select Case mcurMemory
    
           Case Is > 2750   ' Approx 3gb of memory or more
                Select Case curAmtLeft
                       Case Is >= MB_1:   GetBlockSize = MB_1    ' 1,048,576 bytes
                       Case Is >= KB_512: GetBlockSize = KB_512  '   524,288
                       Case Is >= KB_256: GetBlockSize = KB_256  '   262,144
                       Case Is >= KB_128: GetBlockSize = KB_128  '   131,072
                       Case Is >= KB_64:  GetBlockSize = KB_64   '    65,536
                       Case Is >= KB_32:  GetBlockSize = KB_32   '    32,768
                       Case Else:         GetBlockSize = CLng(curAmtLeft)
                End Select

           Case Is > 1750   ' Approx 2gb of memory or more
                Select Case curAmtLeft
                       Case Is >= KB_512: GetBlockSize = KB_512
                       Case Is >= KB_256: GetBlockSize = KB_256
                       Case Is >= KB_128: GetBlockSize = KB_128
                       Case Is >= KB_64:  GetBlockSize = KB_64
                       Case Is >= KB_32:  GetBlockSize = KB_32
                       Case Else:         GetBlockSize = CLng(curAmtLeft)
                End Select
           
           Case Is > 750    ' Approx 1gb of memory or more
                Select Case curAmtLeft
                       Case Is >= KB_128: GetBlockSize = KB_128
                       Case Is >= KB_64:  GetBlockSize = KB_64
                       Case Is >= KB_32:  GetBlockSize = KB_32
                       Case Else:         GetBlockSize = CLng(curAmtLeft)
                End Select
           
           Case Else   ' Less than 1gb of memory
                Select Case curAmtLeft
                       Case Is >= KB_32:  GetBlockSize = KB_32
                       Case Else:         GetBlockSize = CLng(curAmtLeft)
                End Select
    End Select
    
End Function

' ***************************************************************************
' Routine:       GetMemorySize
'
' Description:   Capture memory footprint of this PC.  Used to determine
'                maximum record sizes when writing to disk.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub GetMemorySize()

    Dim typMemStat As MEMORYSTATUSEX  ' For checking PC memory

    On Error GoTo GetMemorySize_Error

    mcurMemory = 0@                         ' Initialize variable (recommended)
    ZeroMemory typMemStat, Len(typMemStat)  ' Clear type structure (recommended)
    typMemStat.dwLength = Len(typMemStat)   ' Load type structure length (required)
    GlobalMemoryStatusEx typMemStat         ' Load memory information into structure
    
    ' Convert large integer to currency
    CopyMemory mcurMemory, typMemStat.ullTotalPhys, 8&     ' Total physical memory
    'CopyMemory mcurMemory, typMemStat.ullAvailPhys, 8&     ' Available physical memory
    'CopyMemory mcurMemory, typMemStat.ullAvailVirtual, 8&  ' Available virtual memory
    
    ' Convert currency to non-decimal.
    ' Example is based on 512mb of physical memory
    ' with approximate calculated values.
    '    46922.1376 ->  447mb Total physical memory
    '    12031.1808 ->  114mb Available physical memory
    '   204222.464  -> 1947mb Available virtual memory
    mcurMemory = CCur(Fix((mcurMemory * 10000@) / MB_1))

GetMemorySize_CleanUp:
    ZeroMemory typMemStat, Len(typMemStat)  ' Clear type structure (recommended)
    On Error GoTo 0
    Exit Sub

GetMemorySize_Error:
    Err.Clear           ' Clear error flag
    mcurMemory = 512@   ' set a minimum value
    Resume GetMemorySize_CleanUp
    
End Sub

' **************************************************************************
' Routine:       CreateNewName
'
' Description:   Renames the old file or folder 26 times.  Beginning with
'                the letter "A" and ending with the letter "Z".  By the
'                time this routine is called, the original file contents
'                have already been overwritten at least once.
'
' Example:        1. C:\Temp\Test File.txt --> C:\Temp\AAAAAAAAA.AAA
'                 2. C:\Temp\AAAAAAAAA.AAA --> C:\Temp\BBBBBBBBB.BBB
'                             ...
'                26. C:\Temp\YYYYYYYYY.YYY --> C:\Temp\ZZZZZZZZZ.ZZZ
'
' Parameters:    strPath - Path and file name
'
' Returns:       new name of the file (Path\new_Filename)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated to use API MoveFileEX()
' ***************************************************************************
Public Function CreateNewName(ByVal strPath As String) As String

    Dim lngIndex    As Long      ' loop counter
    Dim lngLength   As Long      ' length of original name
    Dim lngPointer  As Long      ' pointer in string
    Dim strNewName  As String    ' new folder or file name
    Dim strPrevName As String    ' previous path and filename
    Dim strTempName As String    ' Temp Name only

    Const ROUTINE_NAME As String = "CreateNewName"

    On Error GoTo CreateNewName_Error

    ' if this is a folder then leave
    If Right$(strPath, 1) = "\" Then
        Exit Function
    End If
    
    SetFileAttributes strPath, FILE_ATTRIBUTE_NORMAL           ' Reset path attributes
    strNewName = vbNullString                                            ' Init new name
    strPrevName = strPath                                      ' capture original path\filename
    lngPointer = InStrRev(strPath, "\", Len(strPath))          ' find mode path name
    strTempName = Mid$(strPath, lngPointer + 1)                ' capture filename
    strPath = Left$(strPath, lngPointer)                       ' capture path
    lngLength = Len(strTempName)                               ' capture total length of name
    lngPointer = InStrRev(strTempName, ".", Len(strTempName))  ' find last period in the name
    
    ' Rename file 26 times
    ' Excessive DoEvents are here for a purpose
    For lngIndex = 1 To 26

        DoEvents
        strTempName = String$(lngLength, Chr$(64 + lngIndex))  ' Create new name
        
        ' Re-insert last period, if any
        If lngPointer > 0 Then
            Mid$(strTempName, lngPointer, 1) = "."
        End If
        
        strNewName = strPath & strTempName   ' append new name to path
        
        ' Rename previous file with new filename
        DoEvents
        If MoveFileEx(strPrevName, strNewName, _
                      MOVEFILE_COPY_ALLOWED Or _
                      MOVEFILE_REPLACE_EXISTING) <> 0 Then
        
            strPrevName = strNewName   ' update variable with new name

        Else
            ' An error occurred.  Use last
            ' known file name and leave.
            strNewName = strPrevName
            Exit For    ' exit For..Next loop
        
        End If
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

    Next lngIndex


CreateNewName_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        CreateNewName = strPrevName   ' return last known file name
    Else
        CreateNewName = strNewName    ' return new name
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CreateNewName_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume CreateNewName_CleanUp

End Function

Public Sub Wait(ByVal lngMilliseconds As Long)

    Dim lngPause As Long
        
    ' Calculate a pause
    lngPause = GetTickCount() + lngMilliseconds
    
    Do
        DoEvents
    Loop While lngPause > GetTickCount()
    
End Sub

' ***************************************************************************
' Routine:       ByteArrayToString
'
' Description:   Converts a byte array to string data
'
' Parameters:    abytData - array of bytes
'
' Returns:       Data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function ByteArrayToString(ByRef abytData() As Byte) As String

    ByteArrayToString = StrConv(abytData(), vbUnicode)

End Function

' ***************************************************************************
' Routine:       StringToByteArray
'
' Description:   Converts string data to a byte array
'
' Parameters:    strData - Data string to be converted
'
' Returns:       byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function StringToByteArray(ByVal strData As String) As Byte()

     StringToByteArray = StrConv(strData, vbFromUnicode)

End Function

' ***************************************************************************
' Routine:       GetComplement
'
' Description:   This routine will determine the inverse value of a byte
'                value (0-255).  Using bitwise NOT sometimes generates
'                the direct opposite (15 to -15) because of bit flipping.
'                What we need is a positive representation of the byte.
'
'                      Input        Complement
'                Ex:   15       --> 240       (decimal value)
'                      00001111 --> 11110000  (binary format)
'
' Parameters:    bytData - value to be evaluated
'
' Returns:       Inverse value of an ASCII decimal value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function GetComplement(ByVal bytData As Byte) As Byte

    ' Called by cWipe.MiscPatterns()
    '           cWipe.DoDPatterns()
    '           cWipe.EuroPatterns()
    
    GetComplement = CByte((Not bytData) And &HFF&)   ' return complement byte
    
End Function

' ***************************************************************************
' Routine:       CalcTempFiles
'
' Description:   Calculate number and size of temp files to be created when
'                filling freespace on a partitioned drive.
'
' Parameters:    dblAmtLeft - Amount of freespace on partitioned drive
'                blnNTFS    - Flag designating if this file system is
'                             NTFS or FAT32
'
' Returns:       Two deminsional array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 14-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Added  8GB file size calculation
'              - Updated documentation
' 20-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed an overflow bug while calculating number of files to be
'              created.
' ***************************************************************************
Public Function CalcTempFiles(ByVal dblAmtLeft As Double, _
                              ByVal blnNTFS As Boolean) As Double()

    ' Called by cWipe.FillWithTempFiles()
    
    Dim lngIndex   As Long
    Dim dblFileCnt As Double
    Dim adblData() As Double
    
    On Error Resume Next
    
    Erase adblData()                 ' Always start with empty arrays
    ReDim adblData(0 To 11, 0 To 1)  ' Size two dimensional array
    
    ' On NTFS systems sometimes the last 4 KB are protected.
    ' On FAT32 systems sometimes it is the last 8 KB.
    dblAmtLeft = dblAmtLeft - IIf(blnNTFS, 4096, 8192)
    
    ' Preload file sizes
    ' Odd sizes to prevent boundry errors.
    adblData(0, 1) = 8589946876#   '   8 GB + 4092 bytes
    adblData(1, 1) = 4294971388#   '   4 GB + 4092
    adblData(2, 1) = 2147483644#   '   2 GB - 4
    adblData(3, 1) = 1073741820#   '   1 GB - 4
    adblData(4, 1) = 134217724#    ' 100 MB - 4
    adblData(5, 1) = 16777212#     '  10 MB - 4
    adblData(6, 1) = 1048572#      '   1 MB - 4
    adblData(7, 1) = 65532#        '  64 KB - 4
    adblData(8, 1) = 32764#        '  32 KB - 4
    adblData(9, 1) = 16380#        '  16 KB - 4
    adblData(10, 1) = 4092#        '   4 KB - 4
    
    ' Calculate number of files
    ' to create for each size
    For lngIndex = 0 To 10
        
        DoEvents
        Select Case blnNTFS
        
               Case True    ' NTFS file system.
                    dblFileCnt = Fix(dblAmtLeft / adblData(lngIndex, 1))            ' Calculate number of files
                    adblData(lngIndex, 0) = IIf(dblFileCnt > 0#, dblFileCnt, 0#)    ' Number of files to be created
                    dblAmtLeft = dblAmtLeft - (adblData(lngIndex, 1) * dblFileCnt)  ' Calculate amount left
                    
               Case False   ' FAT32 file system
                    Select Case lngIndex
                           Case 0, 1, 2   ' No files larger than 1 GB (Too slow)
                                adblData(lngIndex, 0) = 0#
                           Case Else
                                dblFileCnt = Fix(dblAmtLeft / adblData(lngIndex, 1))            ' Calculate number of files
                                adblData(lngIndex, 0) = IIf(dblFileCnt > 0#, dblFileCnt, 0#)    ' Number of files to be created
                                dblAmtLeft = dblAmtLeft - (adblData(lngIndex, 1) * dblFileCnt)  ' Calculate amount left
                    End Select
        End Select
            
    Next lngIndex
        
    If dblAmtLeft > 0# Then
        adblData(11, 0) = 1#           ' Create one file
        adblData(11, 1) = dblAmtLeft   ' this size
    Else
        adblData(11, 0) = 0#    ' Nothing to create
        adblData(11, 1) = 0#
    End If
    
    CalcTempFiles = adblData()  ' Return two dimensional array
        
    Erase adblData()   ' Always empty arrays when not needed
    On Error GoTo 0    ' nullify this error trap
    
End Function

Private Sub SwapBytes(ByRef AA As Byte, _
                      ByRef BB As Byte)

    ' Swap byte values
    
    AA = AA Xor BB
    BB = BB Xor AA
    AA = AA Xor BB

End Sub


