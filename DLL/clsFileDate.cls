VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsFileDate
'
' Description:   Module with several routines to make use of the
'                File and System date/time stamps.
'
'                The national language support functions let applications set
'                the locale for the user, identifying the language in which
'                the user carries out work and retrieves strings,
'                representing times, dates, and other information, that are
'                correctly formatted for the given language and location of
'                the world. National language support also includes support
'                for keyboard layouts and language-specific fonts.
'
'                A locale is a collection of language-related,
'                user-preference information represented as a list of values.
'                Each system has at least one installed locale and often has
'                many locales from which the user may choose.
'
'                The system assigns a locale to each thread. Initially, the
'                system assigns the system default locale to the thread. This
'                default locale is set by the user when the system is
'                installed or through the Regional Settings program of
'                Control Panel. If a thread is run in a process belonging to
'                a user, the system assigns the user-default locale to the
'                thread. An application can override either default by using
'                the SetThreadLocale function to explicitly set the locale
'                for a thread.
'
'                There are predefined locale identifiers:
'
'                    LOCALE_SYSTEM_DEFAULT - which identifies the system
'                                            Default Locale
'                    LOCALE_USER_DEFAULT   - which identifies the locale
'                                            of the current user
'                    LOCALE_NEUTRAL        - which identifies the default
'                                            language-neutral locale
'
'                The LOCALE_NEUTRAL identifier is the same as
'                LOCALE_USER_DEFAULT. An application can retrieve the current
'                locale identifiers by using the GetSystemDefaultLCID and
'                GetUserDefaultLCID functions.
'
'                By using the GetSystemDefaultLCID (or GetThreadLocale)
'                together with GetLocaleInfo(), Visual Basic developers can
'                tap into the user's Regional Settings to assure that dates
'                and times are presented in the manner the user has
'                specified. This demo utilizes only a small subset of the
'                available LOCALE_* values possible, specifically those
'                dealing with returning Date information.
'
'                Above description by Randy Birch
'
' Reference:     Randy Birch   http://www.mvps.org/vbnet/index.html
'                SetFileTime: Modify the Date/Time of Folders and Files
'                Written:  21-Sep-2002
'                Updated:  05-Dec-2007
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-JUN-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated enumeration enumDateProperties to match API
'              SetFileTime() parameters.
'              Updated SetDateProperty() routine as to which date property
'              field(s) to be updated.
' 14-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added option to enumDateProperties to handle two of the
'              three date properties.
' 28-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Created read only properties for folder or file date property
'              fields.
' 28-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Renamed three routines, modified and documented
'                  GetSystemDateInfo()   -> SystemDateInfo()
'                  GetPathFileDateInfo() -> FileDateInfo()
'                  SetPathFileDate()     -> SetDateProperty()
' 01-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added additional input testing.  See FileDateInfo()
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced FileExists() and PathExists() routines with
'              IsPathValid() routine.
' 28-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated GetPropertyDates() routine.
' 21-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added new parameter to SetDateProperty() routine.
' 23-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Renamed two date enumeration members.  See enumDateProperties.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants Generic
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsFileDate"

' ***************************************************************************
' Module Constants that pertain to Locale Settings
' ***************************************************************************
  Private Const MAX_SIZE                   As Long = 260     'buffer area
  Private Const LOCALE_SDATE               As Long = &H1D    'date separator
  Private Const LOCALE_STIME               As Long = &H1E    'time separator
  Private Const LOCALE_SSHORTDATE          As Long = &H1F    'short date format string
  Private Const LOCALE_SLONGDATE           As Long = &H20    'long date format string
  Private Const LOCALE_STIMEFORMAT         As Long = &H1003  'time format string
  Private Const LOCALE_S1159               As Long = &H28    'AM designator
  Private Const LOCALE_S2359               As Long = &H29    'PM designator

  ' Constants defined for Date/Time API calls
  Private Const OPEN_EXISTING              As Long = 3
  Private Const FILE_SHARE_READ            As Long = &H1
  Private Const FILE_SHARE_DELETE          As Long = &H4
  Private Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
  Private Const FILE_FLAG_BACKUP_SEMANTICS As Long = &H2000000
  Private Const GENERIC_WRITE              As Long = &H40000000
  Private Const GENERIC_READ               As Long = &H80000000

' ***************************************************************************
' Types defined for Date/Time API calls
' ***************************************************************************
  ' The OFSTRUCT structure contains information about a file that the
  ' OpenFile function opened or attempted to open.
  Private Type OFSTRUCT
       cBytes               As Byte
       fFixedDisk           As Byte
       nErrCode             As Integer
       Reserved1            As Integer
       Reserved2            As Integer
       szPathName(MAX_SIZE) As Byte
  End Type

  ' The FILETIME structure is a 64-bit value representing the number of
  ' 100-nanosecond intervals since January 1, 1601.
  Private Type FILETIME
       dwLowDateTime  As Long
       dwHighDateTime As Long
  End Type

  ' The SYSTEMTIME structure represents a date and time using individual
  ' members for the month, day, year, weekday, hour, minute, second, and
  ' millisecond.
  Private Type SYSTEMTIME
       wYear         As Integer
       wMonth        As Integer
       wDayOfWeek    As Integer
       wDay          As Integer
       wHour         As Integer
       wMinute       As Integer
       wSecond       As Integer
       wMilliseconds As Long
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long
  
  ' The CreateFile function creates, opens, or truncates a file, pipe,
  ' communications resource, disk device, or console. It returns a handle
  ' that can be used to access the object. It can also open and return a
  ' handle to a directory.
  Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
          (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
          ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
          ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
          ByVal hTemplateFile As Long) As Long

  ' The GetLocalTime function retrieves the current local date and time.
  ' Loads the systemTIME structure.
  Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

  ' The GetFileTime function retrieves the date and time that a file was
  ' created, last accessed, and last modified.
  Private Declare Function GetFileTime Lib "kernel32" _
          (ByVal hFile As Long, lpmtypCreateDate As FILETIME, _
          lpmtypLastAccessed As FILETIME, lpmtypLastModified As FILETIME) As Long

  ' The SetFileTime function sets the date and time that a file was
  ' created, last accessed, or last modified.
  Private Declare Function SetFileTime Lib "kernel32" _
          (ByVal hFile As Long, lpmtypCreateDate As FILETIME, _
          lpmtypLastAccessed As FILETIME, lpmtypLastModified As FILETIME) As Long

  ' always release the file handle when finished
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hFile As Long) As Long

  ' The LocalFileTimeToFileTime function converts a local file time
  ' to a file time based on the Coordinated Universal Time (UTC).
  ' LocalFileTimeToFileTime uses the current settings for the time zone
  ' and daylight saving time. Therefore, if it is daylight saving time,
  ' this function will take daylight saving time into account, even if
  ' the time you are converting is in standard time.
  Private Declare Function LocalFileTimeToFileTime Lib "kernel32" _
          (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

  ' The FileTimeToSystemTime function converts a 64-bit file time to
  ' SYSTEMTIME format.
  Private Declare Function FileTimeToSystemTime Lib "kernel32" _
          (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

  ' The FileTimeToLocalFileTime function converts a file time based on
  ' the Coordinated Universal Time (UTC) to a local file time.
  ' FileTimeToLocalFileTime uses the current settings for the time zone
  ' and daylight saving time. Therefore, if it is daylight saving time,
  ' this function will take daylight saving time into account, even if
  ' the time you are converting is in standard time.
  Private Declare Function FileTimeToLocalFileTime Lib "kernel32" _
          (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

  ' The SystemTimeToFileTime function converts a SYSTEMTIME structure
  ' to a FILETIME format.
  Private Declare Function SystemTimeToFileTime Lib "kernel32" _
          (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

  ' A variant time is stored as an 8-byte real value (double), representing
  ' a date between January 1, 1753 and December 31, 2078, inclusive. The
  ' value 2.0 represents January 1, 1900; 3.0 represents January 2, 1900,
  ' and so on. Adding 1 to the value increments the date by a day. The
  ' fractional part of the value represents the time of day. Therefore,
  ' 2.5 represents noon on January 1, 1900; 3.25 represents 6:00 A.M. on
  ' January 2, 1900, and so on. Negative numbers represent the dates prior
  ' to December 30, 1899.
  '
  ' The SystemTimeToVariantTime function converts a system time to a
  ' variant date value.
  Private Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" _
          (lpSystemTime As SYSTEMTIME, vtime As Date) As Long

  ' ZeroMemory is used for clearing contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

  '==========================================================================
  ' Declares defined for Locale Setings API calls
  '
  ' The GetUserDefaultLCID function retrieves the user default locale
  ' identifier.
  Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

  ' This function retrieves information about a user's locale.  The
  ' GetLocaleInfo function always retrieves information in text format.
  ' If the information is a numeric value, the function converts the
  ' number to text using decimal notation.
  Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
          (ByVal Locale As Long, ByVal LCType As Long, _
          ByVal lpLCData As String, ByVal cchData As Long) As Long

  ' GetFileAttributes Function gets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Private Declare Function GetFileAttributes Lib "kernel32" _
          Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
  
  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Private Declare Function SetFileAttributes Lib "kernel32" _
          Alias "SetFileAttributesA" _
          (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

' ***************************************************************************
' Enumerations
' ***************************************************************************
  ' 01-Nov-2008 Reset order to match API SetFileTime()
  Public Enum enumDateProperties
      eCreateDate      ' 0
      eLastAccessed    ' 1
      eLastModified    ' 2
      eTwoDates        ' 3  Last Accessed, Last Modified (most common)
      eThreeDates      ' 4  Create Date, Last Accessed, Last Modified
  End Enum

' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Date)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m dat TimeStamp
' Variable name:     mdatTimeStamp
'
' ***************************************************************************
  Private mdatTimeStamp         As Date
  Private mdatCreateDate        As Date
  Private mdatLastModified      As Date
  Private mdatLastAccessed      As Date
  Private mstrPathFile          As String
  Private mstrTimeFormat        As String
  Private mstrShortDateFormat   As String
  Private mstrLongDateFormat    As String
  Private mstrDateSeparator     As String
  Private mstrTimeSeparator     As String
  Private mtypOFStruct          As OFSTRUCT
  Private mtypSystemTime        As SYSTEMTIME
  Private mtypConverted         As FILETIME
  Private mtypCreateDate        As FILETIME
  Private mtypLastModified      As FILETIME
  Private mtypLastAccessed      As FILETIME
  Private mlngSelectedDateField As enumDateProperties
  
  
' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let StopProcessing(ByVal blnData As Boolean)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents

End Property

Public Property Get StopProcessing() As Boolean
    
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents

End Property

Public Property Let SelectedDateField(ByVal lngData As enumDateProperties)
Attribute SelectedDateField.VB_Description = "Select which date field to be updated as per enumDateField."
    
    ' used when assigning a value to the property, on the left side of an
    ' assignment.     Syntax: x.SelectedDateField = eTwoDates
    
    Select Case lngData
           Case eCreateDate, eLastAccessed, eLastModified, eTwoDates, eThreeDates
                mlngSelectedDateField = lngData
           Case Else
                mlngSelectedDateField = eTwoDates
    End Select
    
End Property

Public Property Let PathFileName(ByVal strData As String)
Attribute PathFileName.VB_Description = "PathName of the file to be updated."

    ' used when assigning a value to the property, on the left side of an
    ' assignment.     Syntax: x.Filename = "C:\Program Files\Test.txt"
     
    Const ROUTINE_NAME As String = "PathFileName"

    mstrPathFile = Trim$(strData)

    If Not IsPathValid(mstrPathFile) Then
        InfoMsg "Cannot locate" & vbNewLine & mstrPathFile & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        mstrPathFile = vbNullString
    End If

End Property

Public Property Get PathFileName() As String
Attribute PathFileName.VB_Description = "PathName of the file to be updated."

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x."C:\Program Files\Test.txt"
    PathFileName = mstrPathFile

End Property

Public Property Let TimeStamp(ByRef datDate As Date)
Attribute TimeStamp.VB_Description = "DateTime to be used for updating."

    ' used when assigning a value to the property, on the left side of an
    ' assignment.     Syntax:  x.TimeStamp = Now()
    mdatTimeStamp = datDate
    
End Property

Public Property Get TimeStamp() As Date
Attribute TimeStamp.VB_Description = "DateTime to be used for updating."

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.TimeStamp
    TimeStamp = mdatTimeStamp

End Property

Public Property Get LongDateFormat() As String

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.LongDateFormat
    LongDateFormat = mstrLongDateFormat

End Property

Public Property Get ShortDateFormat() As String

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.ShortDateFormat
    ShortDateFormat = mstrShortDateFormat

End Property

Public Property Get TimeFormat() As String

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.TimeFormat
    TimeFormat = mstrTimeFormat

End Property

Public Property Get DateSeparator() As String

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.DateSeparator
    DateSeparator = mstrDateSeparator

End Property

Public Property Get TimeSeparator() As String

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.TimeSeparator
    TimeSeparator = mstrTimeSeparator

End Property

Public Property Get CreateDate() As Date

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.CreateDate
    CreateDate = mdatCreateDate

End Property

Public Property Get LastAccessed() As Date

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.LastAccessed
    LastAccessed = mdatLastAccessed

End Property

Public Property Get LastModified() As Date

    ' used when retrieving value of a property, on the right side of an
    ' assignment.     Syntax:  Debug.Print x.LastModified
    LastModified = mdatLastModified

End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       SystemDateInfo
'
' Description:   Format a date/time string based on the format desired by the
'                user obtained from the current system date/time settings.
'
' Parameters:    strDate - Local Date pattern [returned]
'                strTime - Local time pattern [returned]
'                blnUseShortDate - [Optional] Date format to use.
'                     TRUE = Use short date format (Default)
'                     FALSE = Use long date format
'
' Returns:       Formatted date/time stamp
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2002  Randy Birch   http://www.mvps.org/vbnet/index.html
'              Routine created
' 28-JUN-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 28-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Renamed routine
' ***************************************************************************
Public Function SystemDateInfo(ByRef strDate As String, _
                               ByRef strTime As String, _
                      Optional ByVal blnUseShortdate As Boolean = True) As Boolean
Attribute SystemDateInfo.VB_Description = "Format a date/time string based on the format desired by the user obtained from the current system date/time settings."

    Dim datDate As Date

    Const ROUTINE_NAME As String = "SystemDateInfo"

    On Error GoTo SystemDateInfo_Error

    datDate = Empty   ' Initialize date to "12:00:00 AM"
    
    ZeroMemory mtypSystemTime, Len(mtypSystemTime)  ' Clear type structure

    ' GetLocalTime API function retrieves
    ' the current local date and time
    GetLocalTime mtypSystemTime

    ' SystemTimeToVariantTime API function converts
    ' a system time to a variant date value
    SystemTimeToVariantTime mtypSystemTime, datDate
    
    ' Format date/time strings
    If blnUseShortdate Then
        strDate = Format$(datDate, mstrShortDateFormat)
    Else
        strDate = Format$(datDate, mstrLongDateFormat)
    End If

    strTime = Format$(datDate, mstrTimeFormat)
    SystemDateInfo = True

SystemDateInfo_CleanUp:
    datDate = Empty   ' Initialize date to "12:00:00 AM"
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

SystemDateInfo_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strDate = vbNullString
    strTime = vbNullString
    SystemDateInfo = False
    Resume SystemDateInfo_CleanUp

End Function

' ***************************************************************************
' Routine:       FileDateInfo
'
' Description:   Format a date/time string based on the format desired by the
'                user obtained from the passed filename date/time settings.
'                See property PathFileName().
'
' Parameters:    strDate - Local Date pattern returned
'                strTime - Local time pattern returned
'                blnUseShortDate - [Optional] Date format to use.
'                     TRUE = Use short date format (Default)
'                     FALSE = Use long date format
'                lngDateField - [Optional] Date property to be evaluated.
'                     Default = Last modified date
'
' Returns:       Formatted date/time stamp
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2002  Randy Birch   http://www.mvps.org/vbnet/index.html
'              Routine created
' 28-JUN-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 28-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Renamed routine and added Select..Case to determine date
'              selection
' 01-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added additional input testing
' ***************************************************************************
Public Function FileDateInfo(ByRef strDate As String, _
                             ByRef strTime As String, _
                    Optional ByVal blnUseShortdate As Boolean = True, _
                    Optional ByVal lngDateField As enumDateProperties = eLastModified) As Boolean

    Dim hFile   As Long  ' file handle
    Dim lngAttr As Long  ' attribute flags
    
    Const ROUTINE_NAME As String = "FileDateInfo"

    On Error GoTo FileDateInfo_Error

    ClearDateVariables  ' Clear property date variables
    strDate = vbNullString
    strTime = vbNullString
    
    ' Do we have a name to work with?
    If Len(Trim$(mstrPathFile)) = 0 Then
        ' Cannot identify folder or file name
        InfoMsg "Folder or file name not found." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    If lngDateField > 2 Then
        InfoMsg "Cannot identify which one of the date properties are " & _
                "to be returned." & vbNewLine & _
                vbNewLine & Space$(5) & "Creation date" & _
                vbNewLine & Space$(5) & "Last accessed date" & _
                vbNewLine & Space$(5) & "Last modified date" & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    ' capture file attributes
    lngAttr = GetFileAttributes(mstrPathFile)
    
    ' reset file attributes
    SetFileAttributes mstrPathFile, FILE_ATTRIBUTE_NORMAL
        
    hFile = GetHandleID()  ' Get folder/file handle identifier

    If hFile <= 0 Then
        ' Handle identifier is missing or access is denied.
        ' The folder or file may be write protected or is
        ' being accessed by another process.
        InfoMsg "Access denied!" & vbNewLine & mstrPathFile & vbNewLine & _
                "may be write protected or is being accessed by " & _
                "another process." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    GetPropertyDates hFile         ' Get property dates
    CloseHandle hFile              ' Always close file handle
    SetAttr mstrPathFile, lngAttr  ' Reset file attributes

    ' Return the completed formatted strings
    Select Case lngDateField
           
           Case eCreateDate     ' 0
                If blnUseShortdate Then
                    strDate = Format$(mdatCreateDate, mstrShortDateFormat)
                Else
                    strDate = Format$(mdatCreateDate, mstrLongDateFormat)
                End If
                
                strTime = Format$(mdatCreateDate, mstrTimeFormat)
           
           Case eLastAccessed   ' 1
                If blnUseShortdate Then
                    strDate = Format$(mdatLastAccessed, mstrShortDateFormat)
                Else
                    strDate = Format$(mdatLastAccessed, mstrLongDateFormat)
                End If
                
                strTime = Format$(mdatLastAccessed, mstrTimeFormat)
           
           Case eLastModified   ' 2
                If blnUseShortdate Then
                    strDate = Format$(mdatLastModified, mstrShortDateFormat)
                Else
                    strDate = Format$(mdatLastModified, mstrLongDateFormat)
                End If
                
                strTime = Format$(mdatLastModified, mstrTimeFormat)
                
           Case Else
                InfoMsg "Cannot identify which one of the date properties are " & _
                        "to be returned." & vbNewLine & _
                        vbNewLine & Space$(5) & "Creation date" & _
                        vbNewLine & Space$(5) & "Last accessed date" & _
                        vbNewLine & Space$(5) & "Last modified date" & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                        
                ClearDateVariables   ' Clear property date variables
                GoTo FileDateInfo_CleanUp
    End Select
    
    FileDateInfo = True

FileDateInfo_CleanUp:
    CloseAllFiles     ' close any open files
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

FileDateInfo_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strDate = vbNullString
    strTime = vbNullString
    ClearDateVariables   ' Clear property date variables
    FileDateInfo = False
    Resume FileDateInfo_CleanUp

End Function

' ***************************************************************************
' Routine:       SetDateProperty
'
' Description:   Change a file's date/time stamp based on date stored
'                in property TimeStamp().
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2002  Randy Birch   http://www.mvps.org/vbnet/index.html
'              Routine created
' 28-JUN-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated which date property field(s) to be updated
' 28-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Renamed routine and updated documentation
' 21-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added new parameter
' ***************************************************************************
Public Function SetDateProperty(Optional ByVal blnDisplayMsgs As Boolean = True) As Boolean

    Dim hFile       As Long
    Dim typSysTime  As SYSTEMTIME
    Dim typFileTime As FILETIME
    
    Const ROUTINE_NAME As String = "SetDateProperty"

    On Error Resume Next

    SetDateProperty = False                   ' Preset to FALSE
    ClearDateVariables                        ' Clear property date variables
    ZeroMemory typSysTime, Len(typSysTime)    ' Clear type structures
    ZeroMemory typFileTime, Len(typFileTime)

    ' Do we have a name to work with?
    If Len(Trim$(mstrPathFile)) = 0 Then
        ' Cannot identify folder or file name
        InfoMsg "Folder or file name not found." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    SetAttr mstrPathFile, vbNormal   ' reset folder/file attributes
    hFile = GetHandleID()            ' Get folder/file handle identifier
    
    If hFile <= 0 Then
        ' Handle identifier is missing or access is denied.
        ' The folder or file may be write protected or is
        ' being accessed by another process.
        If blnDisplayMsgs Then
            InfoMsg "Access denied!" & vbNewLine & mstrPathFile & vbNewLine & _
                    "may be write protected or is being accessed by " & _
                    "another process." & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        End If
        
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo SetDateProperty_Error

    GetPropertyDates hFile   ' Get property dates
    
    ' If mdatTimeStamp has not been properly set
    ' then use folder/file "Last accessed date"
    If mdatTimeStamp = "12:00:00 AM" Then
        mdatTimeStamp = mdatLastAccessed  ' Use last accessed date
    End If
    
    ' set the day, month and year, and
    ' the hour, minute and second to the
    ' values representing the desired date/time
    With typSysTime
        .wDay = Day(mdatTimeStamp)
        .wMonth = Month(mdatTimeStamp)
        .wYear = Year(mdatTimeStamp)
        .wHour = Hour(mdatTimeStamp)
        .wMinute = Minute(mdatTimeStamp)
        .wSecond = Second(mdatTimeStamp)
    End With
    
    ' Convert the system time (typSysTime)
    ' to a file time (typFileTime) format
    DoEvents
    If CBool(SystemTimeToFileTime(typSysTime, typFileTime)) Then
    
        ' Call LocalFileTimeToFileTime to convert the
        ' local file time to a file time based on the
        ' Coordinated Universal Time (UTC). Conveniently,
        ' the same FILETIME can be used as the in/out
        ' parameters.
        DoEvents
        If CBool(LocalFileTimeToFileTime(typFileTime, typFileTime)) Then
        
            ' Select which date property to update.
            Select Case mlngSelectedDateField
            
                   ' individual date properties
                   Case eCreateDate:   mtypCreateDate = typFileTime
                   Case eLastAccessed: mtypLastAccessed = typFileTime
                   Case eLastModified: mtypLastModified = typFileTime
                   
                   ' Most common dates changed
                   Case eTwoDates
                        mtypLastAccessed = typFileTime
                        mtypLastModified = typFileTime
            
                   ' All three date properties
                   Case eThreeDates
                        mtypCreateDate = typFileTime
                        mtypLastAccessed = typFileTime
                        mtypLastModified = typFileTime
            
                   Case Else
                        mtypLastAccessed = typFileTime
                        mtypLastModified = typFileTime
            End Select
            
            ' Call SetFileTime() to set the date and time
            ' that a file was created, last accessed,
            ' and/or last modified.  Since SetFileTime
            ' returns 1 if successful, cast to return
            ' a Boolean indicating failure or success.
            DoEvents
            If CBool(SetFileTime(hFile, _
                                 mtypCreateDate, _
                                 mtypLastAccessed, _
                                 mtypLastModified)) Then
                                 
                GetPropertyDates hFile   ' Get property dates
    
                ' Compare requested time stamp
                ' with current time stamp
                If mdatTimeStamp = mdatLastAccessed Then
                    SetDateProperty = True    ' Successful update
                End If
            Else
                InfoMsg "Failed to update path/file date properties." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
            End If
        Else
            InfoMsg "Failed to convert local time to file time." & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        End If
    Else
        InfoMsg "Failed to convert system time to file time." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If
    
SetDateProperty_CleanUp:
    ZeroMemory typSysTime, Len(typSysTime)    ' Clear type structures
    ZeroMemory typFileTime, Len(typFileTime)
    Call CloseHandle(hFile)                   ' Always close file handle
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

SetDateProperty_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ClearDateVariables   ' Clear property date variables
    SetDateProperty = False
    Resume SetDateProperty_CleanUp

End Function


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       GetHandleID
'
' Description:   Capture the folder/file unique identifier (handle).  By
'                using CreateFile's special FILE_FLAG_BACKUP_SEMANTICS flag
'                we can obtain a handle to a folder or file and, with
'                SetFileTime, change its created, accessed or modified dates.
'
' Returns:       Numeric folder or file identifier
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2002  Randy Birch   http://www.mvps.org/vbnet/index.html
'              Routine created
' 28-JUN-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified
' ***************************************************************************
Private Function GetHandleID() As Long

    ' Called by FileDateInfo()
    '           SetDateProperty()
    '
    ' Capture path/file handle
     GetHandleID = CreateFile(mstrPathFile, _
                              GENERIC_READ Or GENERIC_WRITE, _
                              FILE_SHARE_READ Or FILE_SHARE_DELETE, _
                              0&, OPEN_EXISTING, _
                              FILE_FLAG_BACKUP_SEMANTICS, 0&)
                                
End Function

' ***************************************************************************
' Routine:       GetPropertyDates
'
' Description:   Capture the file date properties for date created, date
'                last accessed, and date last modified.
'
' Parameters:    hFile - File handle identifier
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2002  Randy Birch   http://www.mvps.org/vbnet/index.html
'              Routine created
' 28-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 28-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated looping to match date property enumeration
' ***************************************************************************
Private Sub GetPropertyDates(ByVal hFile As Long)

    ' Called by FileDateInfo()
    '           SetDateProperty()
    
    Dim datDate      As Date
    Dim typFileTime  As FILETIME
    Dim typLocalTime As FILETIME
    Dim typSysTime   As SYSTEMTIME
    Dim lngDateProp  As enumDateProperties
   
    Const ROUTINE_NAME As String = "GetPropertyDates"

    On Error GoTo GetPropertyDates_Error

    ClearDateVariables   ' Clear property date variables
    
    ' Load property date type structures
    ' with folder or file information
    If CBool(GetFileTime(hFile, _
                         mtypCreateDate, _
                         mtypLastAccessed, _
                         mtypLastModified)) Then
    
        For lngDateProp = 0 To 2
            
            ' Clear local type structures
            ZeroMemory typLocalTime, Len(typLocalTime)
            ZeroMemory typFileTime, Len(typFileTime)
            ZeroMemory typSysTime, Len(typSysTime)
            datDate = Empty
            
            Select Case lngDateProp
                   Case eCreateDate:   typFileTime = mtypCreateDate
                   Case eLastAccessed: typFileTime = mtypLastAccessed
                   Case eLastModified: typFileTime = mtypLastModified
            End Select
            
            ' Convert file time to a local file time
            '    typFileTime -> typLocalTime
            If FileTimeToLocalFileTime(typFileTime, typLocalTime) Then
            
                ' Convert local file time to system time format
                '    typLocalTime -> typSysTime
                If FileTimeToSystemTime(typLocalTime, typSysTime) Then
                    
                    ' Format date using new system time
                    ' ex: 5/9/2002 9:15:51 AM
                    With typSysTime
                        datDate = DateSerial(.wYear, .wMonth, .wDay) & _
                                  Chr$(32) & _
                                  TimeSerial(.wHour, .wMinute, .wSecond)
                    End With
        
                    ' Store newly created date in appropriate date variable
                    Select Case lngDateProp
                           Case eCreateDate:   mdatCreateDate = datDate
                           Case eLastAccessed: mdatLastAccessed = datDate
                           Case eLastModified: mdatLastModified = datDate
                    End Select
            
                End If
            End If
        
        Next lngDateProp
    
    Else
        InfoMsg "Failed to capture Path/File date property information." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If

GetPropertyDates_CleanUp:
    ZeroMemory typLocalTime, Len(typLocalTime)  ' Clear type structures
    ZeroMemory typFileTime, Len(typFileTime)
    ZeroMemory typSysTime, Len(typSysTime)
    datDate = Empty
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetPropertyDates_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ClearDateVariables   ' Clear property date variables
    Resume GetPropertyDates_CleanUp
   
End Sub

' ***************************************************************************
' Routine:       GetDateFormat
'
' Description:   Capture the date formatting for this particular locale
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-JUN-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 06-JUL-2002  Kenneth Ives  kenaso@tx.rr.com
'              Corrected obtaining Locale Country Code ID by capturing the
'              user's personal PC locale settings.
' ***************************************************************************
Private Sub GetDateFormat()

    ' Called by Class_Initialize()
    
    Dim LCID As Long   ' local country code identifier

    On Error GoTo GetDateFormat_Error

    mstrShortDateFormat = vbNullString  ' local short date formatting
    mstrLongDateFormat = vbNullString   ' local long date formatting
    mstrDateSeparator = vbNullString    ' local date separator. Some use a backslash
                              ' while others may use something else.
                              ' Ex:  dash (-), period (.), etc.

    LCID = GetUserDefaultLCID()   ' capture local country code

    ' Get date formats used by this PC
    mstrShortDateFormat = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)
    mstrLongDateFormat = GetUserLocaleInfo(LCID, LOCALE_SLONGDATE)
    mstrDateSeparator = GetUserLocaleInfo(LCID, LOCALE_SDATE)

GetDateFormat_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetDateFormat_Error:
    Err.Clear
    mstrShortDateFormat = vbNullString
    mstrLongDateFormat = vbNullString
    mstrDateSeparator = vbNullString
    Resume GetDateFormat_CleanUp

End Sub

' ***************************************************************************
' Routine:       GetTimeFormat
'
' Description:   Capture the time formatting for this particular locale
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-JUN-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 06-JUL-2002  Kenneth Ives  kenaso@tx.rr.com
'              Corrected obtaining Locale Country Code ID by capturing the
'              user's personal PC locale settings.
' ***************************************************************************
Private Sub GetTimeFormat()

    ' Called by Class_Initialize()
    
    Dim intPosition  As Integer  ' String pointer
    Dim LCID         As Long     ' local country code identifier
    Dim strAM        As String   ' AM/PM designators
    Dim strPM        As String

    On Error GoTo GetTimeFormat_Error

    mstrTimeFormat = vbNullString      ' local time formatting
    mstrTimeSeparator = vbNullString   ' local time separator. Some use a colon
                             ' while others may use something else.
                             ' Ex:  Period (.), Slash (/), etc.

    LCID = GetUserDefaultLCID()  ' capture local country code

    ' get time format used by this PC
    mstrTimeFormat = GetUserLocaleInfo(LCID, LOCALE_STIMEFORMAT)
    mstrTimeSeparator = GetUserLocaleInfo(LCID, LOCALE_STIME)

    ' If there is a trailing time indicator then capture
    ' what is used for AM/PM within this locale time format
    intPosition = InStr(1, Trim$(mstrTimeFormat), " ")
    
    If intPosition > 0 Then
        
        strAM = GetUserLocaleInfo(LCID, LOCALE_S1159)  ' AM indicator
        strPM = GetUserLocaleInfo(LCID, LOCALE_S2359)  ' PM indicator
        
        ' Format time display
        mstrTimeFormat = Trim$(Left$(mstrTimeFormat, intPosition)) & _
                         Chr$(32) & _
                         Trim$(strAM) & Trim$(strPM)
    End If

GetTimeFormat_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetTimeFormat_Error:
    Err.Clear
    mstrTimeFormat = vbNullString
    mstrTimeSeparator = vbNullString
    Resume GetTimeFormat_CleanUp

End Sub

' ***************************************************************************
' Routine:       GetUserLocaleInfo
'
' Description:   Capture the format settings for this particular locale
'
' Parameters:    lngLocaleID - Local country code identifier
'                lngLocaleType - One of the constants listed above that deal
'                       a particular format setting
'
' Returns:       dtea string representing a particular format setting
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-MAY-1999  Randy Birch   http://www.mvps.org/vbnet/index.html
'              Routine created
' 28-JUN-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Private Function GetUserLocaleInfo(ByVal lngLocaleID As Long, _
                                   ByVal lngLocaleType As Long) As String

    ' Called by GetTimeFormat()
    '           GetDateFormat()
    
    Dim strBuffer     As String
    Dim lngDataLength As Long

    strBuffer = Space$(MAX_SIZE)    ' pad the buffer with spaces

    ' Call API function by passing Locale ID of this country
    ' (lngLocaleID) and type of data needed (lngLocaleType)
    lngDataLength = GetLocaleInfo(lngLocaleID, lngLocaleType, strBuffer, Len(strBuffer))

    If lngDataLength > 0 Then
        
        ' Safety measure.  Make a second call.
        '
        strBuffer = Space$(MAX_SIZE)    ' pad the buffer with spaces
    
        ' Call API function again
        lngDataLength = GetLocaleInfo(lngLocaleID, lngLocaleType, strBuffer, Len(strBuffer))
    
        If lngDataLength > 0 Then
    
            ' lngDataLength holds the size of the string
            ' including the terminating null
            GetUserLocaleInfo = Left$(strBuffer, lngDataLength - 1)
        Else
            GetUserLocaleInfo = vbNullString
        End If
    Else
        GetUserLocaleInfo = vbNullString
    End If

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
Private Function CloseAllFiles() As Boolean

    While FreeFile > 1
        Close #FreeFile - 1
    Wend
    
End Function

Private Sub ResetVariables()

    ' Empty type structures
    ZeroMemory mtypOFStruct, Len(mtypOFStruct)
    ZeroMemory mtypSystemTime, Len(mtypSystemTime)
    ZeroMemory mtypConverted, Len(mtypConverted)
        
    mstrPathFile = vbNullString
    mdatTimeStamp = Empty   ' Initialize date to "12:00:00 AM"
    ClearDateVariables      ' Clear property date variables
    
End Sub

Private Sub ClearDateVariables()

    ' Empty type structures
    ZeroMemory mtypCreateDate, Len(mtypCreateDate)
    ZeroMemory mtypLastModified, Len(mtypLastModified)
    ZeroMemory mtypLastAccessed, Len(mtypLastAccessed)
    
    mdatCreateDate = Empty    ' Initialize dates to "12:00:00 AM"
    mdatLastAccessed = Empty
    mdatLastModified = Empty
    
End Sub

Private Sub Class_Initialize()
    
    ' When this class is instantiated capture
    ' date and time formats in use by this PC
    
    ResetVariables
    GetTimeFormat
    GetDateFormat
    
    ' Preset property values
    SelectedDateField = eTwoDates   ' 3 - Last Accessed and Last Modified date
    StopProcessing = False

End Sub

Private Sub Class_Terminate()
    ResetVariables
End Sub
