VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIniMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class module designed for the manipulation of data with either the WIN.INI or private INI files."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsIniMgr
'
' Description:   This class is used to create and manipulate INI files
'
' SPECIAL NOTE:  When making API calls to these routines, REMEMBER the parameters
'                you pass DO NOT HAVE:
'
'                    1.  section name does not have "[ ]"
'                    2.  key name does not have "=" suffix
'                    3.  key value does not have "=" prefix
'
' Important:      This module must have access to modTrimStr.bas
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote class
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed bug. Testing an empty varialbe for existance of data.
'              - Renamed return parameter to a more meaningful name.  See
'                GetAllSectionData(), GetAllSectionTitles() routines.
'              - Moved new parameter from second to last position.  See
'                GetAllSectionTitles() routine.
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Deleted RemoveTrailingNulls() routine from this module.
'              - Changed call to RemoveTrailingNulls() to TrimStr module
'                due to speed and accuracy.
' ***************************************************************************
Option Explicit

' ******************************************************************************
' Constants
' ******************************************************************************
  Private Const MODULE_NAME As String = "clsIniMgr"
  Private Const KB_10       As Long = &H2800&    ' 10240
  Private Const MAX_BYTES   As Long = 256
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long
  
  ' The GetProfileString function retrieves the string associated with
  ' the specified key in the given section of the WIN.INI file. This
  ' function is provided for compatibility with 16-bit Windows-based
  ' applications. Win32-based applications should store initialization
  ' information in the registry.
  Private Declare Function GetProfileString Lib "kernel32" _
          Alias "GetProfileStringA" (ByVal lpAppName As String, _
          ByVal lpKeyName As String, ByVal lpDefault As String, _
          ByVal lpReturnedString As String, ByVal nSize As Long) As Long

  ' The GetProfileSection function retrieves all of the keys and
  ' values for the specified section of the WIN.INI file. This function
  ' is provided for compatibility with 16-bit Windows-based applications.
  ' Win32-based applications should store initialization information in
  ' the registry.
  Private Declare Function GetProfileSection Lib "kernel32" _
          Alias "GetProfileSectionA" (ByVal lpAppName As String, _
          ByVal lpReturnedString As String, ByVal nSize As Long) As Long

  ' The WriteProfileSection function replaces the contents of the
  ' specified section in the WIN.INI file with the specified keys and
  ' values.  This function is provided for compatibility with 16-bit
  ' Windows applications. Win32-based applications should store
  ' initialization information in the registry
  Private Declare Function WriteProfileSection Lib "kernel32" _
          Alias "WriteProfileSectionA" (ByVal lpAppName As String, _
          ByVal lpString As String) As Long

  ' The WriteProfileString function copies a string into the specified
  ' section of the WIN.INI file.  This function is provided for
  ' compatibility with 16-bit Windows-based applications. Win32-based
  ' applications should store initialization information in the registry.
  Private Declare Function WriteProfileString Lib "kernel32" _
          Alias "WriteProfileStringA" (ByVal lpszSection As String, _
          ByVal lpszKeyName As String, ByVal lpszString As String) As Long


' ******************************************************************************
' ================ PRIVATE INI FILES =========================
' ******************************************************************************

  ' The GetPrivateProfileString function retrieves a string from the
  ' specified section in an initialization file. This function is
  ' provided for compatibility with 16-bit Windows-based applications.
  ' Win32-based applications should store initialization information
  ' in the registry.
  Private Declare Function GetPrivateProfileString Lib "kernel32" _
          Alias "GetPrivateProfileStringA" _
          (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
          ByVal lpDefault As String, ByVal lpReturnedString As String, _
          ByVal nSize As Long, ByVal lpFileName As String) As Long

  ' The GetPrivateProfileSection function retrieves all of the keys
  ' and values for the specified section from an initialization file.
  ' This function is provided for compatibility with 16-bit applications
  ' written for Windows. Win32-based applications should store
  ' initialization information in the registry.
  Private Declare Function GetPrivateProfileSection Lib "kernel32" _
          Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
          ByVal lpReturnedString As String, ByVal nSize As Long, _
          ByVal lpFileName As String) As Long

  ' The WritePrivateProfileSection function replaces the keys and
  ' values under the specified section in an initialization file.
  ' This function is provided for compatibility with 16-bit
  ' Windows-based applications. Win32-based applications should
  ' store initialization information in the registry.
  Private Declare Function WritePrivateProfileSection Lib "kernel32" _
          Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
          ByVal lpString As String, ByVal lpFileName As String) As Long

  ' The WritePrivateProfileString function copies a string into
  ' the specified section of the specified initialization file.
  ' This function is provided for compatibility with 16-bit
  ' Windows-based applications. WIn32-based applications should
  ' store initialization information in the registry.
  Private Declare Function WritePrivateProfileString Lib "kernel32" _
          Alias "WritePrivateProfileStringA" _
          (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
          ByVal lpString As Any, ByVal lpFileName As String) As Long

  ' The lstrcpy function copies a string to a buffer. If the function
  ' succeeds, the return value is a pointer to the buffer else null.
  Private Declare Function lstrcpyA Lib "kernel32" _
          (ByVal retval As String, ByVal Ptr As Long) As Long
          
  ' The lstrlen function returns the length in bytes (ANSI version) or
  ' WCHAR values (Unicode version) of the specified string (not
  ' including the terminating null character).
  Private Declare Function lstrlenA Lib "kernel32" _
          (ByVal Ptr As Any) As Long
          
  ' Searches a path for a file name. Returns a pointer to the address
  ' of the string if successful, or a pointer to the beginning of the
  ' path otherwise.
  Private Declare Function PathFindFileName Lib "shlwapi" _
          Alias "PathFindFileNameA" (ByVal pszPath As String) As Long



' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CreateINI
'
' Description:   Determines if an INI file already exist.  If it does not,
'                one is created.  It will not overlay an existing INI files.
'
' Parameters:    strINI_Name - Name of INI file to be created
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Public Function CreateINI(ByVal strINI_Name As String) As Boolean
Attribute CreateINI.VB_Description = "Determines if an INI file already exist.  If it does not, one is created."

    Dim hFile       As Integer
    Dim intPosition As Integer
    Dim strFilename As String
    Dim strAsteriks As String

    On Error GoTo CreateINI_Error

    strAsteriks = String$(60, 42)

    ' Determine if we are to build INI in the application directory or somewhere
    ' else.  Capture just the name of the INI for the heading after we build the
    ' file.
    intPosition = InStrRev(strINI_Name, "\")
    If intPosition = 0 Then
        strFilename = strINI_Name
        strINI_Name = App.Path & "\" & strINI_Name
    Else
        strFilename = Mid$(strINI_Name, intPosition + 1)
    End If

    ' If the file already exist, then leave else create it.
    On Error GoTo CreateINI_Error

    ' Create an INI file
    hFile = FreeFile
    Open strINI_Name For Output As #hFile
    Print #hFile, ";" & strAsteriks
    Print #hFile, ";     " & UCase$(strFilename)
    Print #hFile, ";     Created on " & Format(Now, "d MMM yyyy  h:mm ampm")
    Print #hFile, "; "
    Print #hFile, ";     WARNING!  Do not modify this file manually."
    Print #hFile, ";" & strAsteriks
    Print #hFile, " "
    Close #hFile

    CreateINI = True

CreateINI_CleanUp:
    On Error GoTo 0
    Exit Function

CreateINI_Error:
    ErrorMsg MODULE_NAME, "CreateINI", Err.Description
    CreateINI = False
    Resume CreateINI_CleanUp

End Function

' ***************************************************************************
' Routine:       DoesSectionExist
'
' Description:   Determines if the name of a specific section exists in an
'                INI file.  If the section does not exist, you must create
'                it by calling the SaveSectionTitle() routine.

'
' Parameters:    strINI_Name     - Name of INI file to be updated.
'                strSectionTitle - Section name to be queried.
'
' Syntax:        DoesSectionExist "C:\Windows\Test.ini", "Colors"
'
'                       file:        Test.ini
'                       section:     [Colors]
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Public Function DoesSectionExist(ByVal strINI_Name As String, _
                                 ByVal strSectionTitle As String) As Boolean
Attribute DoesSectionExist.VB_Description = "Determines if the name of a specific section exists in an INI file."

    Dim strTemp As String
    
    On Error GoTo DoesSectionExist_Error
    
    DoesSectionExist = False   ' Preset to FALSE (Does not exist)
    strTemp = vbNullString
    
    ' Passing nulls can cause parts of the INI file to be deleted.  We don't
    ' wish to allow this here so we make a check to prevent this.
    If Len(Trim$(strINI_Name)) = 0 Or _
       Len(Trim$(strSectionTitle)) = 0 Then
    
        Exit Function
    End If
    
    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        Exit Function
    End If
    
    ' Check to see if section title exist
    strTemp = TrimStr(GetSectionTitle(strINI_Name, strSectionTitle))
    
    If Len(strTemp) > 0 Then
        DoesSectionExist = True    ' Found section title
    End If
    
DoesSectionExist_CleanUp:
    On Error GoTo 0
    Exit Function

DoesSectionExist_Error:
    ErrorMsg MODULE_NAME, "DoesSectionExist", Err.Description
    Resume DoesSectionExist_CleanUp

End Function

' ***************************************************************************
' Routine:       GetSectionTitle
'
' Description:   Obtains the name of a specific section in an INI file.
'
' Parameters:    strINI_Name     - Name of INI file to be updated.
'                strSectionTitle - Section name to be queried.
'
' Syntax:        GetSectionTitle "C:\Windows\Test.ini", "Colors"
'
'                       file:        Test.ini
'                       section:     [Colors]
'
' Returns:       Name of the section, if it exists.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  Testing an empty varialbe for existance of data.
' ***************************************************************************
Public Function GetSectionTitle(ByVal strINI_Name As String, _
                                ByVal strSectionTitle As String) As String
Attribute GetSectionTitle.VB_Description = "Obtains the name of a specific section in an INI file."

    Dim strData   As String
    Dim strTest   As String
    Dim lngLength As Long

    On Error GoTo GetSectionTitle_Error
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)
    strSectionTitle = TrimStr(strSectionTitle)
    strData = Space$(MAX_BYTES)

    ' Has required data been passed ?
    If Len(strINI_Name) = 0 Or _
       Len(strSectionTitle) = 0 Then

        GetSectionTitle = vbNullString
        Exit Function
    End If

    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        GetSectionTitle = vbNullString
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Capture the key value from the INI file after determining if the file is
    ' the WIN.INI or a privately named INI file.
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
        
        ' Update WIN.INI file
        lngLength = GetProfileSection(strSectionTitle, _
                                      strData, MAX_BYTES)
    Else
        ' Update privately named INI file
        lngLength = GetPrivateProfileSection(strSectionTitle, _
                                             strData, MAX_BYTES, _
                                             strINI_Name)
    End If

    ' Trim the variable and return the value
    GetSectionTitle = Left$(strData, lngLength)

GetSectionTitle_CleanUp:
    On Error GoTo 0
    Exit Function

GetSectionTitle_Error:
    ErrorMsg MODULE_NAME, "GetSectionTitle", Err.Description
    GetSectionTitle = vbNullString
    Resume GetSectionTitle_CleanUp

End Function

' ***************************************************************************
' Routine:       GetAllSectionTitles
'
' Description:   Queries an INI file and returns the names of all the
'                section titles.
'
' Parameters:    strINI_Name - Input - Name of INI file to be updated
'                astrAllSecTitles() - Output - Array to hold section titles
'
' Syntax:        GetAllSectionTitles "C:\Windows\Test.ini"
'
' Returns:       An array containg all the section titles
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed bug. Testing an empty varialbe for existance of data.
'              - Renamed return parameter to a more meaningful name.
' ***************************************************************************
Public Function GetAllSectionTitles(ByVal strINI_Name As String, _
                                    ByRef astrAllSecTitles() As String)

    Dim lngLength As Long
    Dim strTest   As String
    Dim strData   As String

    On Error GoTo GetAllSectionTitles_Error

    ' Make the string variable large enough
    ' (10k) to hold the returned values
    strData = Space$(KB_10)
    
    Erase astrAllSecTitles()  ' Always start with an empty array
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)

    ' Passing nulls can cause parts of the INI file
    ' to be deleted.  Therefore, we make a check to
    ' prevent this.
    If Len(strINI_Name) = 0 Then
        Exit Function
    End If

    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Read the INI file to find the Section title after we determine if we are
    ' reading the WIN.INI or a private INI file.
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
        
        ' Query WIN.INI file
        lngLength = GetProfileString(vbNullString, "", "", _
                                     strData, KB_10)
    Else
        ' Query privately named INI file
        lngLength = GetPrivateProfileString(vbNullString, "", "", _
                                            strData, KB_10, _
                                            strINI_Name)
    End If

    ' See if this read was succcessful
    If lngLength = 0 Then
        Exit Function
    End If
    
    strData = Left$(strData, lngLength)                          ' Capture returning data
    astrAllSecTitles() = Split(strData, Chr$(0), Len(strData))   ' Save data into an array

GetAllSectionTitles_CleanUp:
    On Error GoTo 0
    Exit Function

GetAllSectionTitles_Error:
    ErrorMsg MODULE_NAME, "GetAllSectionTitles", Err.Description
    Erase astrAllSecTitles()
    Resume GetAllSectionTitles_CleanUp

End Function

' ***************************************************************************
' Routine:       GetKeyValue
'
' Description:   Obtains the value of a specific key in an INI file.
'
' Parameters:    strINI_Name     - Name of INI file to be updated.
'                strSectionTitle - Section name to be created.
'                strKeyName      - Name of key to be queried.
'                StrDefault      - Return this value if key is empty
'
' Syntax:        GetKeyValue "C:\Windows\Test.ini", "Colors", "Red"
'
'                       file:        Test.ini
'                       section:     [Colors]
'                       key label:   Red=255   <--Returns this value
'
' Returns:       Value of a specific key. Max return length is 255.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  Testing an empty varialbe for existance of data.
' ***************************************************************************
Public Function GetKeyValue(ByVal strINI_Name As String, _
                            ByVal strSectionTitle As String, _
                            ByVal strKeyName As String, _
                   Optional ByVal strDefault As String = vbNullString) As String

    Dim strData   As String
    Dim strTest   As String
    Dim lngLength As Long

    On Error GoTo GetKeyValue_Error
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)
    strSectionTitle = TrimStr(strSectionTitle)
    strKeyName = TrimStr(strKeyName)
    strData = Space$(MAX_BYTES)

    ' Has required data been passed ?
    If Len(strINI_Name) = 0 Or _
       Len(strSectionTitle) = 0 Or _
       Len(strKeyName) = 0 Then

        GetKeyValue = vbNullString
        Exit Function
    End If

    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        GetKeyValue = vbNullString
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Capture the key value from the INI file
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
        
        ' Update WIN.INI file
        lngLength = GetProfileString(strSectionTitle, _
                                     strKeyName, "", _
                                     strData, MAX_BYTES)
    Else
        ' Update privately named INI file
        lngLength = GetPrivateProfileString(strSectionTitle, _
                                            strKeyName, "", _
                                            strData, MAX_BYTES, _
                                            strINI_Name)
    End If

    ' Trim the variable and return the value
    strTest = Left$(strData, lngLength)
    
    If Len(strTest) = 0 Then
        ' create a key with the default value
        SaveOneKeyValue strINI_Name, strSectionTitle, strKeyName, strDefault
        GetKeyValue = strDefault
    Else
        GetKeyValue = strTest
    End If
    
GetKeyValue_CleanUp:
    On Error GoTo 0
    Exit Function

GetKeyValue_Error:
    ErrorMsg MODULE_NAME, "GetKeyValue", Err.Description
    GetKeyValue = vbNullString
    Resume GetKeyValue_CleanUp

End Function

' ***************************************************************************
' Routine:       SaveSectionTitle
'
' Description:   Creates a section title in an INI file.
'
' Parameters:    strINI_Name     - Name of INI file to be updated.
'                strSectionTitle - Section name to be created.
'
' Syntax:        SaveSectionTitle "C:\Windows\Test.ini", "Colors"
'
'                       file:        Test.ini
'                       section:     [Colors]
'
' Returns:       Based on completion.  True or False
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  Testing an empty varialbe for existance of data.
' ***************************************************************************
Public Function SaveSectionTitle(ByVal strINI_Name As String, _
                                 ByVal strSectionTitle As String) As Boolean

    Dim strTest   As String
    Dim lngLength As Long

    On Error GoTo SaveSectionTitle_Error

    SaveSectionTitle = False   ' Preset to FALSE
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)
    strSectionTitle = TrimStr(strSectionTitle)

    ' Has required data been passed ?
    If Len(strINI_Name) = 0 Or _
       Len(strSectionTitle) = 0 Then

        Exit Function
    End If
    
    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Write to INI file after determining if the file is the WIN.INI or a
    ' privately named INI file.
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
    
        ' Update WIN.INI file
        lngLength = WriteProfileSection(strSectionTitle, "")
    Else
        ' Update privately named INI file
        lngLength = WritePrivateProfileSection(strSectionTitle, _
                                               "", strINI_Name)
    End If

SaveSectionTitle_CleanUp:
    SaveSectionTitle = CBool(lngLength)
    On Error GoTo 0
    Exit Function

SaveSectionTitle_Error:
    ErrorMsg MODULE_NAME, "SaveSectionTitle", Err.Description
    lngLength = 0
    Resume SaveSectionTitle_CleanUp

End Function

' ***************************************************************************
' Routine:       SaveCompleteSection
'
' Description:   Creates a section title in an INI file.
'
' Parameters:    strINI_Name     - Name of INI file to be updated.
'                strSectionTitle - Section name to be created.
'                strKeyData      - Keys and their values
'                Ex:  strKeyData = "Primary=Purple" & Chr$(0) & _
'                                  "Secondary=Green" & Chr$(0) & _
'                                  "Level=" & CStr(2) & Chr$(0) & Chr$(0)
'
' Syntax:        SaveCompleteSection "C:\Windows\Test.ini", "Colors", strKeyData
'
' Returns:       Based on completion.  True or False
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  Testing an empty varialbe for existance of data.
' ***************************************************************************
Public Function SaveCompleteSection(ByVal strINI_Name As String, _
                                    ByVal strSectionTitle As String, _
                                    ByVal strKeyData As String) As Boolean

    Dim strTest   As String
    Dim lngLength As Long

    On Error GoTo SaveCompleteSection_Error

    SaveCompleteSection = False   ' Preset to FALSE
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)
    strSectionTitle = TrimStr(strSectionTitle)
    strKeyData = TrimStr(strKeyData)

    ' Has required data been passed ?
    If Len(strINI_Name) = 0 Or _
       Len(strSectionTitle) = 0 Or _
       Len(strKeyData) = 0 Then

        Exit Function
    End If
    
    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Write to INI file after determining if the file is the WIN.INI or a
    ' privately named INI file.
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
    
        ' Update WIN.INI file
        lngLength = WriteProfileSection(strSectionTitle, _
                                        strKeyData)
    Else
        ' Update privately named INI file
        lngLength = WritePrivateProfileSection(strSectionTitle, _
                                               strKeyData, _
                                               strINI_Name)
    End If

SaveCompleteSection_CleanUp:
    SaveCompleteSection = CBool(lngLength)
    On Error GoTo 0
    Exit Function

SaveCompleteSection_Error:
    ErrorMsg MODULE_NAME, "SaveCompleteSection", Err.Description
    lngLength = 0
    Resume SaveCompleteSection_CleanUp

End Function

' ***************************************************************************
' Routine:       SaveOneKeyValue
'
' Description:   Saves a value to a specific key in an INI file.
'
' Parameters:    strINI_Name     - Name of INI file to be updated.
'                strSectionTitle - Section name to be updated.
'                strKeyName      - Specific key to be updated.
'                strKeyValue     - Value to be saved.
'
' Syntax:        SaveOneKeyValue "C:\Windows\Test.ini", "Colors", "Red", "255"
'
'                       file:          Test.ini
'                       section:       [Colors]
'                       key & value:   Red=255
'
' Returns:       Based on completion.  True or False
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  Testing an empty varialbe for existance of data.
' ***************************************************************************
Public Function SaveOneKeyValue(ByVal strINI_Name As String, _
                                ByVal strSectionTitle As String, _
                                ByVal strKeyName As String, _
                                ByVal strKeyValue As String) As Boolean
Attribute SaveOneKeyValue.VB_Description = "Saves a value to a specific key in an INI file."

    Dim strTest   As String
    Dim lngLength As Long

    On Error GoTo SaveOneKeyValue_Error

    SaveOneKeyValue = False   ' Preset to FALSE
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)
    strSectionTitle = TrimStr(strSectionTitle)
    strKeyValue = TrimStr(strKeyValue)

    ' Has required data been passed ?
    If Len(strINI_Name) = 0 Or _
       Len(strSectionTitle) = 0 Or _
       Len(strKeyValue) = 0 Then

        Exit Function
    End If

    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Write to INI file after we determine if we are updating the WIN.INI or
    ' a private INI file.
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
        
        ' Update WIN.INI file
        lngLength = WriteProfileString(strSectionTitle, _
                                       strKeyName, _
                                       strKeyValue)
    Else
        ' Update privately named INI file
        lngLength = WritePrivateProfileString(strSectionTitle, _
                                              strKeyName, _
                                              strKeyValue, _
                                              strINI_Name)
    End If

SaveOneKeyValue_CleanUp:
    SaveOneKeyValue = CBool(lngLength)
    On Error GoTo 0
    Exit Function

SaveOneKeyValue_Error:
    ErrorMsg MODULE_NAME, "SaveOneKeyValue", Err.Description
    lngLength = 0
    Resume SaveOneKeyValue_CleanUp

End Function

' ***************************************************************************
' Routine:       DeleteOneKey
'
' Description:   Deletes a particular key or a complete section within an
'                INI file.
'
' Parameters:    strINI_Name     - Name of INI file to be updated
'                strSectionTitle - Section name to be updated
'                strKeyName      - Specific key to be deleted.
'
' Syntax:        DeleteOneKey "C:\Windows\Test.ini", "Colors", "Red"
'
'                       file:        Test.ini
'                       section:     [Colors]
'                       key label:   Red=0
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  Testing an empty varialbe for existance of data.
' ***************************************************************************
Public Function DeleteOneKey(ByVal strINI_Name As String, _
                             ByVal strSectionTitle As String, _
                             ByVal strKeyName As String) As Boolean
Attribute DeleteOneKey.VB_Description = "Deletes a particular key or a complete section within an INI file."

    Dim strTest   As String
    Dim lngLength As Long

    On Error GoTo DeleteOneKey_Error

    DeleteOneKey = False   ' Preset to FALSE
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)
    strSectionTitle = TrimStr(strSectionTitle)

    ' Has required data been passed ?
    If Len(strINI_Name) = 0 Or _
       Len(strSectionTitle) = 0 Then

        Exit Function
    End If
    
    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Read the INI file to find the Section title after we determine if we are
    ' reading the WIN.INI or a private INI file.
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
        
        ' Make API call to perform the delete
        lngLength = WriteProfileString(strSectionTitle, _
                                        strKeyName, _
                                        vbNullString)
    Else
        ' Make API call to perform the delete
        lngLength = WritePrivateProfileString(strSectionTitle, _
                                              strKeyName, _
                                              vbNullString, _
                                              strINI_Name)
    End If

DeleteOneKey_CleanUp:
    DeleteOneKey = CBool(lngLength)
    On Error GoTo 0
    Exit Function

DeleteOneKey_Error:
    ErrorMsg MODULE_NAME, "DeleteOneKey", Err.Description
    lngLength = 0
    Resume DeleteOneKey_CleanUp

End Function

' ***************************************************************************
' Routine:       DeleteSection
'
' Description:   Deletes a complete section within an INI file.
'
' Parameters:    strINI_Name     - Name of INI file to be updated
'                strSectionTitle - Section name to be updated
'
' Syntax:        DeleteSection "C:\Windows\Test.ini", "Colors"
'
'                       file:        Test.ini
'                       section:     [Colors]
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  Testing an empty varialbe for existance of data.
' ***************************************************************************
Public Function DeleteSection(ByVal strINI_Name As String, _
                              ByVal strSectionTitle As String) As Boolean

    Dim strTest   As String
    Dim lngLength As Long

    On Error GoTo DeleteSection_Error

    DeleteSection = False   ' Preset to FALSE
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)
    strSectionTitle = TrimStr(strSectionTitle)

    ' Has required data been passed ?
    If Len(strINI_Name) = 0 Or _
       Len(strSectionTitle) = 0 Then

        Exit Function
    End If
    
    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Read the INI file to find the Section title after we determine if we are
    ' reading the WIN.INI or a private INI file.
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
        
        ' Make API call to perform the delete
        lngLength = WriteProfileString(strSectionTitle, _
                                        vbNullString, "")
    Else
        ' Make API call to perform the delete
        lngLength = WritePrivateProfileString(strSectionTitle, _
                                              vbNullString, "", _
                                              strINI_Name)
    End If

DeleteSection_CleanUp:
    DeleteSection = CBool(lngLength)
    On Error GoTo 0
    Exit Function

DeleteSection_Error:
    ErrorMsg MODULE_NAME, "DeleteSection", Err.Description
    lngLength = 0
    Resume DeleteSection_CleanUp

End Function

' ***************************************************************************
' Routine:       GetAllSectionData
'
' Description:   Queries a section and returns all of its values in an INI
'                file in the elements of an array.
'
' Parameters:    strINI_Name     - Input - Name of INI file to be updated
'                strSectionTitle - Input - Section name to be queried.
'                astrAllSecData()  - Output - Array of section data
'
' Syntax:        GetAllSectionData "C:\Windows\Test.ini", "Colors"
'
'                       file:        Test.ini
'                       section:     [Colors]
'
' Returns:       An array containg all the keys and their values.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2002  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 29-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed bug. Testing an empty varialbe for existance of data.
'              - Renamed return parameter to a more meaningful name.
'              - Moved new parameter from second to last position.
' ***************************************************************************
Public Function GetAllSectionData(ByVal strINI_Name As String, _
                                  ByVal strSectionTitle As String, _
                                  ByRef astrAllSecData() As String)

    Dim lngLength As Long
    Dim strTest   As String
    Dim strData   As String

    On Error GoTo GetAllSectionData_Error

    ' Make the string variable large enough
    ' (10k) to hold the returned values
    strData = Space$(KB_10)
    
    Erase astrAllSecData()   ' Always start with an empty array
    
    ' Passing nulls can cause parts of the INI file to be deleted
    strINI_Name = TrimStr(strINI_Name)
    strSectionTitle = TrimStr(strSectionTitle)

    ' Has required data been passed ?
    If Len(strINI_Name) = 0 Or _
       Len(strSectionTitle) = 0 Then

        Exit Function
    End If
    
    ' If there are no periods in strINI_Name then the
    ' filename has been supplied without an extension.
    ' Like the VB SaveSetting statement, we assume
    ' we are dealing with an INI file extension,
    ' therefore, so one is appended prior to continuing.
    If InStr(1, strINI_Name, ".") = 0 Then
        strINI_Name = strINI_Name & ".ini"
    End If

    ' Check to see if the INI file exist
    If Not IsPathValid(strINI_Name) Then
        Exit Function
    End If

    ' Capture just name of file
    strTest = GetFilePart(strINI_Name)

    ' Read the INI file to find the Section title after we determine if we are
    ' reading the WIN.INI or a private INI file.
    If StrComp(strTest, "win.ini", vbTextCompare) = 0 Then
        
        ' Query WIN.INI file
        lngLength = GetProfileSection(strSectionTitle, _
                                      strData, KB_10)
    Else
        ' Query privately named INI file
        lngLength = GetPrivateProfileSection(strSectionTitle, _
                                             strData, KB_10, _
                                             strINI_Name)
    End If

    ' See if this read was succcessful
    If lngLength = 0 Then
        Exit Function
    End If

    strData = Left$(strData, lngLength)                    ' Capture returning data
    astrAllSecData() = Split(strData, Chr$(0), lngLength)  ' Save data into an array

GetAllSectionData_CleanUp:
    On Error GoTo 0
    Exit Function

GetAllSectionData_Error:
    ErrorMsg MODULE_NAME, "GetAllSectionData", Err.Description
    Erase astrAllSecData()
    Resume GetAllSectionData_CleanUp

End Function



' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

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
  
' ***************************************************************************
' Routine:       GetFilePart
'
' Description:   Parses a path, ignoring the drive letter or UNC server/share
'                path parts.  Returns the full name of the file (drive letter
'                or UNC server/share).
'
'                Requires Version 5.00 and later of Shlwapi.dll
'
' Parameters:    strPath - Path to be parsed
'
' Returns:       Returns the full name of the file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function GetFilePart(ByVal strPath As String) As String

    GetFilePart = GetStrFromPtrA(PathFindFileName(strPath))
   
End Function

' ***************************************************************************
' Routine:       GetStrFromPtrA
'
' Description:   Strips a data string from memory
'
' Parameters:    lngTestData - Length of data to be returned
'
' Returns:       ASCII text string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function GetStrFromPtrA(ByVal lngTestData As Long) As String

    ' Given a pointer to a string, return the string
    GetStrFromPtrA = String$(lstrlenA(ByVal lngTestData), 0) ' preload with nulls
    lstrcpyA ByVal GetStrFromPtrA, ByVal lngTestData         ' Strip data
   
End Function


