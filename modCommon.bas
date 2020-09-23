Attribute VB_Name = "modCommon"
' ***************************************************************************
' Module:   modCommon
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
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME    As String = "modCommon"
  Private Const MF_BYPOSITION  As Long = &H400
  Private Const MF_REMOVE      As Long = &H1000
  Private Const MAX_SIZE       As Long = 260
  Private Const KB_1           As Long = 1024        ' 1 kilobyte
  Private Const HWND_TOPMOST   As Long = -1          ' bring to top and stay there
  Private Const HWND_NOTOPMOST As Long = -2
  Private Const SWP_NOMOVE     As Long = 2           ' don't move window
  Private Const SWP_NOSIZE     As Long = 1           ' don't size window
  Private Const HWND_FLAGS     As Long = SWP_NOMOVE Or SWP_NOSIZE
  Private Const MAX_LONG       As Long = &H7FFFFFFF  ' 2147483647
  Private Const GB_4           As Double = (2# ^ 32)  ' 4294967296
    
' ***************************************************************************
' API Declares - Local
' ***************************************************************************
  ' Parses a path to determine if it is a directory root. Returns TRUE
  ' if the specified path is a root, or FALSE otherwise.
  Private Declare Function PathIsRoot Lib "shlwapi" Alias "PathIsRootA" _
          (ByVal pszPath As String) As Long
  
  ' Truncates a path to fit within a certain number of characters by
  ' replacing path components with elipses.  Called by ShrinkTofit().
  Private Declare Function PathCompactPathEx Lib "shlwapi" Alias "PathCompactPathExA" _
          (ByVal pszOut As String, ByVal pszSrc As String, _
          ByVal cchMax As Long, ByVal dwFlags As Long) As Long

  ' The GetWindowsDirectory function retrieves the path of the Windows
  ' directory. The Windows directory contains such files as Windows-based
  ' applications, initialization files, and Help files.
  Private Declare Function GetWindowsDirectory Lib "kernel32" _
          Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
          ByVal nSize As Long) As Long

  ' Changes the size, position, and Z order of a child, pop-up, or top-level
  ' window. These windows are ordered according to their appearance on the
  ' screen. The topmost window receives the highest rank and is the first
  ' window in the Z order.
  Private Declare Function SetWindowPos Lib "user32" _
          (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
          ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
          ByVal cy As Long, ByVal wFlags As Long) As Long
  
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  ' ====== DisableX API Declares ============================================
  ' The DrawMenuBar function redraws the menu bar of the specified window.
  ' If the menu bar changes after Windows has created the window, this
  ' function must be called to draw the changed menu bar.  If the function
  ' fails, the return value is zero.
  Private Declare Function DrawMenuBar Lib "user32" _
          (ByVal hwnd As Long) As Long
  
  ' The GetMenuItemCount function determines the number of items in the
  ' specified menu.  If the function fails, the return value is -1.
  Private Declare Function GetMenuItemCount Lib "user32" _
          (ByVal hMenu As Long) As Long
  
  ' The GetSystemMenu function allows the application to access the window
  ' menu (also known as the System menu or the Control menu) for copying
  ' and modifying.  If the bRevert parameter is FALSE (0&), the return
  ' value is the handle of a copy of the window menu.  If the function
  ' fails, the return value is zero.
  Private Declare Function GetSystemMenu Lib "user32" _
          (ByVal hwnd As Long, ByVal bRevert As Long) As Long
  
  ' The RemoveMenu function deletes a menu item from the specified menu.
  ' If the menu item opens a drop-down menu or submenu, RemoveMenu does
  ' not destroy the menu or its handle, allowing the menu to be reused.
  ' Before this function is called, the GetSubMenu function should retrieve
  ' the handle of the drop-down menu or submenu.  If the function fails,
  ' the return value is zero.
  Private Declare Function RemoveMenu Lib "user32" _
          (ByVal hMenu As Long, ByVal nPosition As Long, _
          ByVal wFlags As Long) As Long
  ' =========================================================================


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

Public Sub AlwaysOnTop(ByRef frm As Form, _
                       ByVal blnOnTop As Boolean)

    ' This routine uses an argument to determine whether
    ' to make the specified form always on top or not
    '
    ' Syntax:
    '        AlwaysOnTop form_name, True    ' Place on top
    On Error GoTo AlwaysOnTop_Error

    If blnOnTop Then
        ' stay as topmost window
        SetWindowPos frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, HWND_FLAGS
    Else
        ' not on top anymore
        SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, HWND_FLAGS
    End If

AlwaysOnTop_CleanUp:
    On Error GoTo 0
    Exit Sub

AlwaysOnTop_Error:
    ErrorMsg MODULE_NAME, "AlwaysOnTop", Err.Description
    Resume AlwaysOnTop_CleanUp

End Sub

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
'                Ex:  75231309824 -> 70.1 GB
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
    
    Const MAX_DECIMALS As Long = 5   ' Change to meet special needs
    
    On Error GoTo DisplayNumber_Error
    
    dblValue = dblCapacity   ' I do this for debugging purposes
    intCount = 0             ' Counter for KB_1
    DisplayNumber = vbNullString
    
    If dblValue > 0 Then
        
        ' Must be a positive value
        If lngDecimals < 1 Then
            lngDecimals = 0
        End If
        
        ' Maximum of 5 decimal positions.
        If lngDecimals > MAX_DECIMALS Then
            lngDecimals = MAX_DECIMALS
        End If
    
        ' Loop thru input value and determine how
        ' many times it can be divided by 1024 (1 KB)
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
        
        DisplayNumber = DisplayNumber & " " & _
                        Choose(intCount + 1, "Bytes", "KB", "MB", "GB", "TB", "PB")
    Else
    
        ' No value was passed to this routine
        If lngDecimals = 0 Then
            DisplayNumber = "0 Bytes"     ' Format value with no decimal positions
        Else
            DisplayNumber = "0.0 Bytes"   ' Format value with one decimal position
        End If
    
    End If

DisplayNumber_Error:

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
    
' **************************************************************************
' Routine:       IsWindowsFolder
'
' Description:   Determines if this drive/folder is where the Windows
'                folder resides.
'
' Parameters:    strData - Drive/folder to be evaluated
'
' Returns:       TRUE - This is the Windows path/folder
'                FALSE - Windows path/folder not found
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Feb-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated testing logic
' ***************************************************************************
Public Function IsWindowsFolder(ByVal strData As String) As Boolean

    Dim strWinFolder As String
    Dim lngDataLen   As Long
    Dim lngRetLen    As Long
    
    Const ROUTINE_NAME As String = "IsWindowsFolder"

    On Error GoTo IsWindowsFolder_Error

    IsWindowsFolder = False          ' Preset flag to FALSE
    lngDataLen = Len(strData)        ' Capture length of incoming data
    strWinFolder = Space$(MAX_SIZE)  ' Preload buffer
    
    ' Make API call to determine where the Windows folder resides
    lngRetLen = GetWindowsDirectory(strWinFolder, MAX_SIZE)
    
    ' API return length should be length of path
    ' Ex:  lngRetLen = 10 = Len("C:\Windows")
    If lngRetLen > 0 Then
        
        strWinFolder = Left$(strWinFolder, lngRetLen)   ' Remove trailing nulls
        
        ' See if just the drive letter ("C:\") was
        ' passed and if it matches the drive letter
        ' of the Windows drive.
        If StrComp(Left$(strWinFolder, lngDataLen), strData, vbTextCompare) = 0 Then
            IsWindowsFolder = True   ' Found Windows folder name
            Exit Function
        End If
        
        ' See if the folder name passed to here
        ' contains the Windows path/folder name.
        If StrComp(strWinFolder, Left$(strData, lngRetLen), vbTextCompare) = 0 Then
            IsWindowsFolder = True   ' Found Windows folder name
        Else
            IsWindowsFolder = False  ' Cannot find Windows folder name
        End If

    End If

IsWindowsFolder_CleanUp:
    On Error GoTo 0
    Exit Function

IsWindowsFolder_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    IsWindowsFolder = False
    Resume IsWindowsFolder_CleanUp
    
End Function

' ***************************************************************************
' Routine:       ReadIniFile
'
' Description:   Capture values from INI file.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Public Sub ReadIniFile(Optional ByRef avntOption As Variant = Empty)

    With gobjINIMgr
        
        glngWipeMethod = Val(.GetKeyValue(gstrINI, INI_DEFAULT, "Default", "1"))
        gstrDescription = .GetKeyValue(gstrINI, INI_DEFAULT, "Description", "All zeroes  [ 0x00 * 1 ]")
        gstrPreviousPath = .GetKeyValue(gstrINI, INI_DEFAULT, "PreviousPath", "C:\")
        glngPasses = .GetKeyValue(gstrINI, INI_DEFAULT, "NbrOfPasses", "1")
        
        ' optional return values
        If IsArrayInitialized(avntOption) Then
        
            avntOption(0) = CByte(.GetKeyValue(gstrINI, INI_DEFAULT, "LogData", "0"))
            avntOption(1) = CByte(.GetKeyValue(gstrINI, INI_DEFAULT, "DisplayFinishMsg", "0"))
            avntOption(2) = CByte(.GetKeyValue(gstrINI, INI_DEFAULT, "Verify", "0"))
            avntOption(3) = CByte(.GetKeyValue(gstrINI, INI_DEFAULT, "DisplayVerifyMsg", "1"))
            avntOption(4) = CByte(.GetKeyValue(gstrINI, INI_DEFAULT, "ZeroLastWrite", "0"))
            avntOption(5) = CByte(.GetKeyValue(gstrINI, INI_DEFAULT, "LogEncryptParms", "0"))
            
            gblnLogData = CBool(avntOption(0))
            gblnDisplayFinishMsg = CBool(avntOption(1))
            gblnVerifyData = CBool(avntOption(2))
            gblnDisplayVerifyMsgs = CBool(avntOption(3))
            gblnZeroLastWrite = CBool(avntOption(4))
            gblnLogEncryptParms = CBool(avntOption(5))
                
        Else
        
            gblnLogData = CBool(.GetKeyValue(gstrINI, INI_DEFAULT, "LogData", "0"))
            gblnDisplayFinishMsg = CBool(.GetKeyValue(gstrINI, INI_DEFAULT, "DisplayFinishMsg", "0"))
            gblnVerifyData = CBool(.GetKeyValue(gstrINI, INI_DEFAULT, "Verify", "0"))
            gblnDisplayVerifyMsgs = CBool(.GetKeyValue(gstrINI, INI_DEFAULT, "DisplayVerifyMsg", "1"))
            gblnZeroLastWrite = CBool(.GetKeyValue(gstrINI, INI_DEFAULT, "ZeroLastWrite", "0"))
            gblnLogEncryptParms = CBool(.GetKeyValue(gstrINI, INI_DEFAULT, "LogEncryptParms", "0"))
            
        End If
    
    End With
    
End Sub

' ***************************************************************************
' Routine:       WriteIniFile
'
' Description:   Update values in INI file.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Public Sub WriteIniFile(Optional ByRef avntOption As Variant = Empty)

    With gobjINIMgr
    
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "Default", CStr(glngWipeMethod)
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "Description", gstrDescription
        .SaveOneKeyValue gstrINI, INI_DEFAULT, "NbrOfPasses", glngPasses

        ' optional update values
        If IsArrayInitialized(avntOption) Then
            
            .SaveOneKeyValue gstrINI, INI_DEFAULT, "LogData", avntOption(0)
            .SaveOneKeyValue gstrINI, INI_DEFAULT, "DisplayFinishMsg", avntOption(1)
            .SaveOneKeyValue gstrINI, INI_DEFAULT, "Verify", avntOption(2)
            .SaveOneKeyValue gstrINI, INI_DEFAULT, "DisplayVerifyMsg", avntOption(3)
            .SaveOneKeyValue gstrINI, INI_DEFAULT, "ZeroLastWrite", avntOption(4)
            .SaveOneKeyValue gstrINI, INI_DEFAULT, "LogEncryptParms", avntOption(5)
            
            gblnLogData = CBool(avntOption(0))
            gblnDisplayFinishMsg = CBool(avntOption(1))
            gblnVerifyData = CBool(avntOption(2))
            gblnDisplayVerifyMsgs = CBool(avntOption(3))
            gblnZeroLastWrite = CBool(avntOption(4))
            gblnLogEncryptParms = CBool(avntOption(5))
            
        End If
    
    End With
    
End Sub

' ***************************************************************************
' Routine:       BuildIniFile
'
' Description:   Create an INI file.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 19-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated INI file format
' ***************************************************************************
Public Sub BuildIniFile()

    Dim strSection As String
    
    ' Create an INI file with a heading
    gobjINIMgr.CreateINI gstrINI
        
    ' Format section keys and values
    strSection = vbNullString
    strSection = strSection & "Item_1=01|All zeroes  [ 0x00 * 1 ] |Low | " & vbNewLine
    strSection = strSection & "Item_2=02|Random data  [ Random * 1 ] | Medium |Optional last write with zeroes" & vbNewLine
    strSection = strSection & "Item_3=03|US DoD Short  [ 3 writes ] | Medium |Optional last write with zeroes" & vbNewLine
    strSection = strSection & "Item_4=04|US DoD Long  [ 7 writes ] | High |Optional last write with zeroes" & vbNewLine
    strSection = strSection & "Item_5=05|N.A.T.O.  [ 7 writes ] | High |Optional last write with zeroes" & vbNewLine
    strSection = strSection & "Item_6=06|German VSITR  [ 7 writes ] | High |Optional last write with zeroes" & vbNewLine
    strSection = strSection & "Item_7=07|Bruce Schneier  [ 7 writes ] | High |Optional last write with zeroes" & vbNewLine
    strSection = strSection & "Item_8=08|Peter Gutmann  [ 35 writes ] | High |Optional last write with zeroes" & vbNewLine
    strSection = strSection & "Item_9=09|Rijndael (AES) Encryption | High |Random password, key and block size" & vbNewLine
    strSection = strSection & "Item_10=10|Blowfish Encryption | High |Random password and key length" & vbNewLine
    strSection = strSection & "Item_11=11|Twofish Encryption | High |Random password and key length" & vbNewLine
    strSection = strSection & "Item_12=12|ArcFour Encryption | High |Random password and key length" & vbNewLine
    
    ' Update INI file. Create section with keys and values
    gobjINIMgr.SaveCompleteSection gstrINI, "PERMANENT_METHODS", strSection
    
    ' Format section keys and values
    strSection = vbNullString
    strSection = strSection & "Default=1" & vbNewLine
    strSection = strSection & "Description=All zeroes  [ 0x00 * 1 ]" & vbNewLine
    strSection = strSection & "LogName=NO_FILE" & vbNewLine
    strSection = strSection & "LogData=0" & vbNewLine
    strSection = strSection & "Verify=0" & vbNewLine
    strSection = strSection & "DisplayMsg=0" & vbNewLine
    strSection = strSection & "NbrOfPasses=1" & vbNewLine
    strSection = strSection & "DisplayFinishMsg=0" & vbNewLine
    strSection = strSection & "DisplayVerifyMsg=0" & vbNewLine
    strSection = strSection & "ZeroLastWrite=0" & vbNewLine
    strSection = strSection & "LogEncryptParms=0" & vbNewLine
    strSection = strSection & "PreviousPath=C:\" & vbNewLine
    
    ' Update INI file. Create section with keys and values.
    gobjINIMgr.SaveCompleteSection gstrINI, "DEFAULT_SELECTION", strSection
    
    ' Update INI file. Create section title only.
    gobjINIMgr.SaveSectionTitle gstrINI, "CUSTOM_PATTERNS"
    
End Sub

' ***************************************************************************
' Routine:       BuildLogFile
'
' Description:   Create a log file and folder if they do not exist.
'
' Parameters:    blnCreateNew - OPTIONAL - Create a new log file.
'                    DEFAULT - FALSE do not create a file.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Public Sub BuildLogFile()

    Dim hFile As Integer    ' first free file handle
    
    Const LINE1 As String = "DiscDataWipe Log File"
    Const LINE2 As String = "dd.mm.yyyy hh:mm:ss  Description"
    
    On Error GoTo BuildLogFile_Error

    ' if the log folder does not exist then
    ' create folder and log file.
    If Not IsPathValid(gstrLogFolder) Then
        
        MkDir gstrLogFolder   ' create log folder
        CreateLogFileName
        
        hFile = FreeFile
        Open gstrLogFile For Output As #hFile
        Print #hFile, LINE1
        Print #hFile, LINE2
        Print #hFile, String$(80, "*")
        Close #hFile

    Else
        ' if the log file does not exist then
        ' create a log file.
        If Not IsPathValid(gstrLogFile) Then
            
            CreateLogFileName
            
            hFile = FreeFile
            Open gstrLogFile For Output As #hFile
            Print #hFile, LINE1
            Print #hFile, LINE2
            Print #hFile, String$(80, "*")
            Close #hFile
        End If
    
    End If
    
BuildLogFile_CleanUp:
    On Error GoTo 0
    Exit Sub

BuildLogFile_Error:
    ErrorMsg MODULE_NAME, "BuildLogFile", Err.Description
    Resume BuildLogFile_CleanUp

End Sub

' **************************************************************************
' Routine:       UpdateLogFile
'
' Description:   Updates the log file with the date, time, and a message.
'
' Parameters:    strMsg - message string to be written to the log file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Sub UpdateLogFile(ByRef strMsg As String)

    Dim hFile        As Integer
    Dim strTimeStamp As String
    
    On Error GoTo UpdateLogFile_Error

    BuildLogFile                            ' Verify there is a log file
    strTimeStamp = GetTimeStamp()           ' capture system timestamp
    
    hFile = FreeFile                        ' get first free file handle
    Open gstrLogFile For Append As #hFile   ' open file for append
    Print #hFile, strTimeStamp & strMsg     ' Append the data record
    Close #hFile                            ' close the log file
    
UpdateLogFile_CleanUp:
    strMsg = vbNullString
    On Error GoTo 0
    Exit Sub

UpdateLogFile_Error:
    ErrorMsg MODULE_NAME, "UpdateLogFile", Err.Description
    Resume UpdateLogFile_CleanUp
    
End Sub

Public Sub GetLogFilename()

    Dim strLogDate    As String
    Dim strTemp       As String
    Dim strJulianDate As String
    
    On Error GoTo GetLogFilename_Error

    gblnLogData = gobjINIMgr.GetKeyValue(gstrINI, INI_DEFAULT, "LogData", "0")
    strTemp = gobjINIMgr.GetKeyValue(gstrINI, INI_DEFAULT, "LogName", "NO_FILE")
    
    If gblnLogData Then
                
        If Len(Trim$(strTemp)) = 0 Or _
           StrComp(strTemp, "NO_FILE") = 0 Then
           
            CreateLogFileName
           
        ElseIf IsNumeric(Mid$(strTemp, 3, 5)) Then
        
            strLogDate = Mid$(strTemp, 3, 5)  ' strip out the julian date
            strJulianDate = GetJulianDate()   ' get the current julian date
            
            ' do the julian dates match
            If StrComp(strJulianDate, strLogDate) <> 0 Then
            
                ' dates are different
                CreateLogFileName
            End If
        Else
            ' create a log filename
            CreateLogFileName
        End If
    
        gobjINIMgr.SaveOneKeyValue gstrINI, INI_DEFAULT, "LogName", gstrLogFilename
        
    Else
        ' make sure we have a default value
        gobjINIMgr.SaveOneKeyValue gstrINI, INI_DEFAULT, "LogName", "NO_FILE"
    End If

GetLogFilename_CleanUp:
    On Error GoTo 0
    Exit Sub

GetLogFilename_Error:
    ErrorMsg MODULE_NAME, "GetLogFilename", Err.Description
    Resume GetLogFilename_CleanUp
    
End Sub

Public Sub CreateLogFileName()

    ' create a log file name consisting of the Julian data
    ' Example:  02/01/2004  -->  DD04032.log
    
    gstrLogFilename = "DD" & GetJulianDate() & ".log"
    gstrLogFile = QualifyPath(gstrLogFolder) & gstrLogFilename
    
End Sub

Public Function GetTimeStamp() As String

    ' Format system date and time
    GetTimeStamp = Format$(Now(), "dd.mm.yyyy") & " " & _
                   Format$(Now(), "hh:mm:ss") & "  "
    
End Function

' ***************************************************************************
' Routine:       GetJulianDate
'
' Description:   This procedure takes a normal date format (that is, 1/1/94)
'                and converts it to the appropriate Julian date (yyddd).
'
'                Most government agencies and contractors require the use of
'                Julian dates. A Julian date starts with a two-digit year,
'                and then counts the number of days from January 1 of that
'                year.
'
'                The Julian Date is returned in string format so as to
'                display all 5 digits to include any leading zeroes. If it
'                were returned in numeric format the Julian Date might be
'                truncated.  For example the Julian Date "00001"
'                (Jan 1, 2000) would appear as 1 because all leading
'                zeroes would be dropped.
'
' Parameters:    datDate - Date to be converted
'
' Returns:       Formatted date in string format to display all 5 digits to
'                include any leading zeroes.
'                Ex:  8/17/2007 --> "07229"
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Aug-2007  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetJulianDate(Optional ByVal datDate As Date = Empty) As String

    ' See if date variable has been
    ' properly initialised
    If datDate = "12:00:00 AM" Then
        datDate = Now()   ' Use current system date
    End If
    
    GetJulianDate = Format$(datDate, "yy") & _
                    Format$(DatePart("y", datDate), "000")

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
' Routine:       ByteArrayToLong (formerly named BytesToLong)
'
' Description:   Convert data from a byte array into a long integer. This
'                routine assumes that the byte array will have at least
'                4 elements.
'
' Reference:     Convert 4 Bytes to Long in VB
'                Filipe Lage
'                http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=56804&lngWId=1
'
' Parameters:    abytData() - Array to hold the data
'                lngIdx     - position to start within the array
'
' Returns:       Long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Renamed and updated this routine
' ***************************************************************************
Public Function ByteArrayToLong(ByRef abytData() As Byte, _
                       Optional ByVal lngIdx As Long = 0) As Long

    Const ROUTINE_NAME As String = "ByteArrayToLong"

    On Error GoTo ByteArrayToLong_Error

    ' make sure there is at least one byte
    If UBound(abytData) < 1 Then

        InfoMsg "There is not enough data in the incoming array to " & _
                "convert to a long integer." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo ByteArrayToLong_CleanUp

    End If

    ' Test pointer value
    Select Case lngIdx

           Case Is < 0
                InfoMsg "The starting position must zero or greater." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                GoTo ByteArrayToLong_CleanUp

           Case Is >= UBound(abytData)
                InfoMsg "Starting position in byte array exceeds size of array." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                GoTo ByteArrayToLong_CleanUp

           Case Is > (UBound(abytData) - 3)
                InfoMsg "Incoming array does not have enough data to convert." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                GoTo ByteArrayToLong_CleanUp
    End Select

    ' Convert to hex string then to long integer
    ByteArrayToLong = Val("&H" & Right$("0" & Hex$(abytData(lngIdx)), 2) & _
                                 Right$("0" & Hex$(abytData(lngIdx + 1)), 2) & _
                                 Right$("0" & Hex$(abytData(lngIdx + 2)), 2) & _
                                 Right$("0" & Hex$(abytData(lngIdx + 3)), 2))

ByteArrayToLong_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

ByteArrayToLong_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    ByteArrayToLong = 0
    Resume ByteArrayToLong_CleanUp

End Function

' ***************************************************************************
' Routine:       UnsignedToLong
'
' Description:   This function takes a Double containing a value in the
'                range of an unsigned Long and returns a Long that you can
'                pass to an API that requires an unsigned LongConvert an
'                unsigned double into a long.
'
' Parameters:    dblValue - Number to be converted
'
' Returns:       A long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-JUL-1998  Microsoft Knowledge Base Article - 189323
'              HOWTO: Convert Between Signed and Unsigned Numbers
'              http://support.microsoft.com/default.aspx?scid=kb;en-us;189323
' ***************************************************************************
Public Function UnsignedToLong(ByVal dblValue As Double) As Long

    ' If an overflow occures then ignore it
    If dblValue < 0 Or dblValue >= GB_4 Then Error 6 ' Overflow

    ' Make sure the data does not exceed that of a long integer.
    If dblValue <= MAX_LONG Then
        UnsignedToLong = dblValue
    Else
        UnsignedToLong = dblValue - GB_4  ' subtract if we exceed a long
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

Public Sub Wait(ByVal sngSecond As Single)

    Dim sngTarget As Single
        
    sngTarget = Timer  ' Capture current time
        
    Do While (Timer - sngTarget) < sngSecond
            
        ' If we cross midnight, back up one day
        DoEvents
        If Timer < sngTarget Then
            sngTarget = sngTarget - 86400  ' Number of seconds in a day
        End If
        
    Loop
    
End Sub

' ***************************************************************************
' Routine:       DisableX
'
' Description:   Remove the "X" from the window and menu
'
'                A VB developer may find themselves developing an application
'                whose integrity is crucial, and therefore must prevent the
'                user from accidentally terminating the application during
'                its life, while still displaying the system menu.  And while
'                Visual Basic does provide two places to cancel an impending
'                close (QueryUnload and Unload form events) such a sensitive
'                application may need to totally prevent even activation of
'                the shutdown.
'
'                Although it is not possible to simply disable the Close button
'                while the Close system menu option is present, just a few
'                lines of API code will remove the system menu Close option
'                and in doing so permanently disable the titlebar close button.
'
' Parameters:    frmName - Name of form
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-Jul-1998  Randy Birch
'              RemoveMenu: Killing the Form's Close Menu and 'X' Button
'              http://www.mvps.org/vbnet/index.html
' ***************************************************************************
Public Sub DisableX(ByRef frmName As Form)

    Dim hMenu          As Long
    Dim lngMenuItemCnt As Long
        
    ' Obtain the handle to the form's system menu
    hMenu = GetSystemMenu(frmName.hwnd, 0&)
    
    If hMenu Then
        
        ' Obtain the handle to the form's system menu
        lngMenuItemCnt = GetMenuItemCount(hMenu)
        
        ' Remove the system menu Close menu item.
        ' The menu item is 0-based, so the last
        ' item on the menu is lngMenuItemCnt - 1
        RemoveMenu hMenu, lngMenuItemCnt - 1, _
                   MF_REMOVE Or MF_BYPOSITION
        
        ' Remove the system menu separator line
        RemoveMenu hMenu, lngMenuItemCnt - 2, _
                   MF_REMOVE Or MF_BYPOSITION
        
        ' Force a redraw of the menu. This
        ' refreshes the titlebar, dimming the X
        DrawMenuBar frmName.hwnd
    
    End If
    
End Sub


