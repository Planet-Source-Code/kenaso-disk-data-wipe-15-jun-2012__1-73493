VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsWipe.cls
'
' Description:   Routines that perform the wiping processes.  Excessive
'                DoEvents is to capture keystrokes.  In particular, if the
'                user opts to stop processing.  Todays faster CPU's will
'                sometimes fly past a boolean IF statement if a DoEvents
'                is not present to allow a momentary pause to check on
'                external processes.
'
' Important:      This module must have access to modTrimStr.bas
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
' 11-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Added functionality to completely wipe a USB drive.  Created
'                two new routines, ProcessUSB() and FillWithTempFiles().
'              - ProcessUSB() and WipeTheFreeSpace() routines now call the
'                FillWithTempFiles() routine.
' 15-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Optimized *Pattern() routines.  Removed some obsolete
'                variables and updated some documentation.
'              - Updated percent logic in CalculateProgress() routine since
'                I am using currency based values.
' 20-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added new event to set a flag for displaying the name of the
'              folder that is in the process of being removed.
' 01-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Removed Get/Put references for API Read/Write.
'              - Added a new property to delete the files in a folder structure
'                but leave the folder structure in tact.
' 09-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified FillWithTempFiles() routine for a maximum file size
'              of 1 gb for fixed drives.
' 11-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Consolidated cipher and hash selection to a single routine
'                when using encrypted overwrite option.
'              - Replaced UpdateLogData() event with CurrentPattern() event.
'              - Added new property CurrentPattern().
' 12-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified WipeTheFreeSpace() routine with the capability to
'              toggle between three passes using blank spaces or one pass
'              using an alternate wiping method pattern.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated SetPathFileDate() to update all three date property
'              fields of a file or folder.
' 24-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Created property named AlternateMethod to determine if to use
'                US DoD Short method to wipe free space.
'              - Fixed bug to keep current progress bar empty when not using
'                an alternate method for wiping.
' 26-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Joe for noticing the erroneous way of performing
'              the selection of date and time in the SetPathFileDate()
'              routine.
' 13-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Speeded up data processing:
'              - Added GetMemorySize() routine to capture total physical
'                memory of current PC.
'              - Updated GetBlockSize() routine by adding new block size
'                selections based on physical memory
'              - Precalculate amount of random data to be created if it
'                is to be accessed.
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated CreateNewName(), WipeFile(), WipeFolder() routines.
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Updated WipeTheFreeSpace() routine to display properly
'                when wiping freespace with the alternate pattern.
'              - Updated ChangeAlgorithms() routine.
'              - Removed GetProvider() and ReleaseProvider() routines.
'              - Combined EncryptData() with EncryptPattern() routine.
' 02-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Encryption options now include:
'                  Rijndael [US Gov Advanced Encryption Standard (AES)]
'                  Serpent  [Finalist in NIST encryption contest]
'                  Twofish  [Finalist in NIST encryption contest]
'                  ArcFour  [Strong encryption algorithm]
' 10-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed bottleneck in WipeFile() routine by removing call to
'                WAIT module.  Speeded up the processing.
'              - Added CalcTempFiles() routine to determine number and size
'                of temp files to be created when wiping freespace.  Called
'                by FillWithTempFiles().
' 14-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated CalcTempFiles() routine by adding 8GB file size.
'              - Updated FillWithTempFiles() routine.
' 16-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated year selection in SetPathFileDate() routine.
' 20-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed an overflow bug while calculating number of files to be
'              created in CalcTempFiles() routine.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed progress display when using encryption.  See bottom of
'                this module.
'              - Replaced Serpent encryption with Blowfish. Strong and fast.
'                See ChangeAlgorithms(), EncryptPatterns() routines.
' ***************************************************************************
Option Explicit

' ***************************************************************************
'  Size definitions used by various disk manufacturers
'
'    Bit        0 or 1
'    Nibble     4 Bits
'    Byte       8 Bits
'    Kibibit    1,024 bits
'    Kilobit    1,000 bits
'    Kibibyte   1,024 bytes
'    Kilobyte   1,000 bytes
'    Mebibit    1,048,576 bits
'    Megabit    1,000,000 bits
'    Mebibyte   1,048,576 bytes
'    Megabyte   1,000,000 bytes
'    Gibibit    1,073,741,824 bits
'    Gigabit    1,000,000,000 bits
'    Gibibyte   1,073,741,824 bytes
'    Gigabyte   1,000,000,000 bytes
'    Tebibit    1,099,511,627,776 bits
'    Terabit    1,000,000,000,000 bits
'    Tebibyte   1,099,511,627,776 bytes
'    Terabyte   1,000,000,000,000 bytes
'    Pebibit    1,125,899,906,842,624 bits
'    Petabit    1,000,000,000,000,000 bits
'    Pebibyte   1,125,899,906,842,624 bytes
'    Petabyte   1,000,000,000,000,000 bytes
' ***************************************************************************

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME                 As String = "clsWipe"
  Private Const TEMP_FOLDER                 As String = "DD_Temp"
  Private Const ABEND_MSG                   As String = "User may have opted to stop processing"
  Private Const WAIT_MSG                    As String = "Calculating folders and files"
  Private Const RE_EVAL_DR                  As String = " ... Re-Evaluating drive"
  Private Const FREE_SPACE_PATTERN          As String = "Binary 0's  [ Ignore space messages ]"
  Private Const CLEANUP_MSG                 As String = "REMOVING WORK FILES - Please wait"
  Private Const FILE_ATTRIBUTE_HIDDEN       As Long = &H2
  Private Const FILE_ATTRIBUTE_NORMAL       As Long = &H80
  Private Const INVALID_HANDLE_VALUE        As Long = -1
  Private Const MAX_SIZE                    As Long = 260
  Private Const MAX_INT                     As Long = &H7FFF       ' 32767
  Private Const KB_4                        As Long = &H1000&      ' 4096
  Private Const KB_16                       As Long = &H4000&      ' 16384
  Private Const MB_5                        As Long = &H500000     ' 5242880
  Private Const GB_1                        As Long = &H40000000   ' 1073741824
  
' ***************************************************************************
' Type structures
' ***************************************************************************
  ' The FILETIME structure is a 64-bit value representing the number of
  ' 100-nanosecond intervals since January 1, 1601.
  Private Type FILETIME
       dwLowDateTime  As Long
       dwHighDateTime As Long
  End Type

  ' The WIN32_FIND_DATA structure describes a file found by the FindFirstFile
  ' or FindNextFile function. If a file has a long filename, the complete
  ' name appears in the cFileName field, and the 8.3 format truncated version
  ' of the name appears in the cAlternate field. Otherwise, cAlternate is empty.
  Private Type WIN32_FIND_DATA
       dwFileAttributes As Long          ' file attributes
       ftCreationTime   As FILETIME      ' File creation date and time
       ftLastAccessTime As FILETIME      ' File last accessed date and time
       ftLastWriteTime  As FILETIME      ' File last modified date and time
       nFileSizeHigh    As Long          ' file sizes over 2GB (2,147,483,647)
       nFileSizeLow     As Long          ' file sizes under 2GB (2,147,483,647)
       dwReserved0      As Long
       dwReserved1      As Long
       cFilename        As String * MAX_SIZE  ' full file name w/o path
       cAlternate       As String * 14        ' short (8.3) file name w/o path
  End Type

  ' User defined type structures
  Private Type FILE_PARAMS
      sFileRoot   As String
      sFileList() As String
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The FindFirstFile function searches a directory for a file whose name
  ' matches the specified filename. FindFirstFile examines subdirectory names
  ' as well as filenames.  The FindFirstFile function opens a search handle
  ' and returns information about the first file whose name matches the
  ' specified pattern. Once the search handle is established, you can use the
  ' FindNextFile function to search for other files that match the same
  ' pattern. When the search handle is no longer needed, close it by using
  ' the FindClose function.  The FindFirstFile function searches for files by
  ' name only; it cannot be used for attribute-based searches.
  Private Declare Function FindFirstFile Lib "kernel32" _
          Alias "FindFirstFileA" (ByVal lpFileName As String, _
          lpFindFileData As WIN32_FIND_DATA) As Long

  ' The FindNextFile function continues a file search from a previous call
  ' to the FindFirstFile function.  The FindNextFile function searches
  ' for files by name only; it cannot be used for attribute-based searches.
  Private Declare Function FindNextFile Lib "kernel32" _
          Alias "FindNextFileA" (ByVal hFindFile As Long, _
          lpFindFileData As WIN32_FIND_DATA) As Long

  ' The FindClose function closes the specified search handle. The
  ' FindFirstFile and FindNextFile functions use the search handle
  ' to locate files with names that match a given name.  Always close
  ' file handle when FINISHED
  Private Declare Function FindClose Lib "kernel32" _
          (ByVal hFindFile As Long) As Long

  ' The CloseHandle function closes an open object handle. CloseHandle invalidates
  ' the specified object handle, decrements the object’s handle count, and performs
  ' object retention checks. Once the last handle to an object is closed, the
  ' object is removed from the operating system.  This function does not close
  ' module objects.
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long

  ' The DeleteFile function deletes an existing file.  If the function
  ' succeeds, the return value is nonzero.
  Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
          (ByVal lpFileName As String) As Long
  
' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumWipeMethod
      eWipe_AlternateMethod   '  0 - Use alternate method to wipe freespace
      eWipe_BinaryZeroes      '  1 - Binary zeroes (0x00)
      eWipe_RandomData        '  2 - Random genereated data
      eWipe_US_DoD_Short      '  3 - US DoD 5220.22-M E
      eWipe_US_DoD_Long       '  4 - US DoD 5220.22-M ECE
      eWipe_NATO              '  5 - North Atlantic Treaty Organization (NATO)
      eWipe_GermanVISTR       '  6 - Germany BSI Verschlusssachen-IT-Richtlinien (VSITR)
      eWipe_BruceSchneier     '  7 - Bruce Schneier (security technologist, cryptoanalyst)
      eWipe_PeterGutmann      '  8 - Peter Gutmann (Computer Science professor)
      eWipe_Rijndael          '  9 - Encryption [Rijndael - NIST winner] AES [Advanced Encryption Standard]
      eWipe_Blowfish          ' 10 - Encryption [Blowfish - Strong, reliable and fast]
      eWipe_Twofish           ' 11 - Encryption [Twofish  - NIST runner up]
      eWipe_ArcFour           ' 12 - Encryption [ArcFour  - Strong, reliable and fast]
      eWipe_CustomPattern     ' 13 - User defined pattern
  End Enum

  Public Enum enumEncryptAlgo
      eAlgo_Rijndael   ' 0 - Rijndael algorithm (AES)
      eAlgo_Blowfish   ' 1 - Blowfish algorithm
      eAlgo_Twofish    ' 2 - Twofish algorithm
      eAlgo_ArcFour    ' 3 - ArcFour algorithm
  End Enum
  
' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Collection)
'                    |  |     |----- Variable subname
'                    - --- --------------
' Naming standard:   m col Folders
' Variable name:     mcolFolders
'
' ***************************************************************************
  Private mlngPause          As Long                ' Delay time
  Private mlngStartTime      As Long                ' when the process starts
  Private mcurFilesize       As Currency            ' Current original file size (EncryptPAttern only)
  Private mcurBlockSize      As Currency            ' Current position in file (EncryptPAttern only)
  Private mcurByteCount      As Currency            ' Total byte count
  Private mcurFileCount      As Currency            ' Total file count
  Private mcurFolderCount    As Currency            ' Total folder count
  Private mcurCurrPerPass    As Currency            ' Current amount for this pass
  Private mcurOA_MaxAmt      As Currency            ' Total amount for progress bar
  Private mcurOA_MaxPerFile  As Currency            ' Total byte count per file
  Private mcurOA_CurrentAmt  As Currency            ' Track total progress
  Private mstrStatistics     As String              ' last entry to the log if successful
  Private mstrTopLevelFolder As String              ' name of top level folder
  Private mblnDiskFull       As Boolean             ' Used when filling disk with null files
  Private mblnMultiple       As Boolean             ' Used only in custom patterns
  Private mblnProcessUSB     As Boolean             ' True if called by ProcessUSB routine
  Private mtypFP             As FILE_PARAMS         ' holds search parameters
  Private mcolFolders        As Collection          ' collection of folder names (32kb limit)
  Private mobjRandom         As cPrng               ' Generate random data
  Private mobjFileDate       As cFileDate           ' Process file property dates
  Private mobjBigFiles       As cBigFiles           ' Manipulate file data

' ***************************************************************************
' Module variables to hold property values
' ***************************************************************************
  Private mlngWipeMethod       As enumWipeMethod    ' type of wipe method to use
  Private mlngEncryptAlgo      As enumEncryptAlgo   ' Encryption algorithm
  Private mlngPasses           As Long              ' number of passes over the drive
  Private mlngTypeTarget       As Long              ' type of drive
  Private mlngProtectiveItems  As Long              ' Number of protected wipe options
  Private mblnLogData          As Boolean           ' flag denotes whether to update a log file
  Private mblnDoFolders        As Boolean           ' flag denotes process a folder
  Private mblnUSB_Drive        As Boolean           ' flag denotes a USB drive (Greater than 8gb)
  Private mblnFlashDrive       As Boolean           ' flag denotes a flash drive (8gb or less)
  Private mblnVerifyData       As Boolean           ' flag denote to verify the last write
  Private mblnDisplayMsgs      As Boolean           ' flag denotes displaying the verify messages
  Private mblnDoSubFolders     As Boolean           ' flag denotes process subfolders
  Private mblnZeroLastWrite    As Boolean           ' flag designates to zero the last pass
  Private mblnWipeFreeSpace    As Boolean           ' flag denotes wipe free space only
  Private mblnLogEncryptParms  As Boolean           ' Flag denotes logging AES encrypt parameters
  Private mblnAlternateMethod  As Boolean           ' flag denotes using DoD Short for wiping free space
  Private mblnRemoveTopFolder  As Boolean           ' flag denotes wipe top level folder also
  Private mblnKeepDirStructure As Boolean           ' flag denotes removing files only
  Private mstrWipePath         As String            ' wipe free space path
  Private mstrCurrentPattern   As String            ' current wipe pattern in use
  Private mastrCustom()        As String            ' Hold custom patterns
  
' ***************************************************************************
' Progress display calculations.  See bottom of this module.
' ***************************************************************************
  Private WithEvents mobjArcFour  As cArcFour
Attribute mobjArcFour.VB_VarHelpID = -1
  Private WithEvents mobjRijndael As cRijndael
Attribute mobjRijndael.VB_VarHelpID = -1
  Private WithEvents mobjBlowfish As cBlowFish
Attribute mobjBlowfish.VB_VarHelpID = -1
  Private WithEvents mobjTwofish  As cTwofish
Attribute mobjTwofish.VB_VarHelpID = -1


' ***************************************************************************
' ****                         Events                                    ****
' ***************************************************************************
  ' Raises an accurate count of the passes performed
  Public Event CountPasses(ByVal lngCurrentPass As Long, _
                           ByVal lngMaxPasses As Long)

  ' Raises accumulation information during the wipe process.
  ' 12-Jul-2008 Removed curFolderCount as obsolete.  Added
  ' flag for tracking application temp files only.
  Public Event CountTotals(ByVal strPathFile As String, _
                           ByVal curFileCount As Currency, _
                           ByVal curByteCount As Currency, _
                           ByVal curFileSize As Currency, _
                           ByVal blnCountOnly As Boolean, _
                           ByVal blnTempFiles As Boolean)

  ' 11-Aug-2008 Added event
  ' Raises the name of the current pattern being used.
  Public Event CurrentPattern(ByVal strCurrentPattern As String)

  ' Raises an accurate count of the bytes written
  Public Event CurrentProgress(ByVal curByteCount As Currency, _
                               ByVal curMaxAmount As Currency)

  ' Raises the Elapsed time since the process started.
  Public Event ElapsedTime(ByVal strElapsedTime As String)
Attribute ElapsedTime.VB_Description = "Raises the Elapsed time snce the process started."

  ' Raises the current percentage for the individual progress bar.
  Public Event FileProgress(ByVal lngFileProgress As Long)

  ' Raises an accurate count of the overall progress
  Public Event OverallProgress(ByVal curByteCount As Currency, _
                               ByVal curMaxAmount As Currency)

  ' Raises the time remaining and transfer rate.
  Public Event TimeRemaining(ByVal strTimeRemaining As String, _
                             ByVal strTransferRate As String)
Attribute TimeRemaining.VB_Description = "Raises the time remaining and transfer rate."

  ' Raises the current percentage for the overall progress bar.
  Public Event TotalProgress(ByVal lngTotalProgress As Long)

  ' Raises the data string to be written to the log file.
  Public Event UpdateLogData(ByVal strLogRecord As String)
Attribute UpdateLogData.VB_Description = "Raises the data string to be written to the log file."

  ' Display a Stand-by message
  Public Event WaitMsg(ByVal strMsg As String, ByVal lngColor As Long)

  ' 10-Feb-2010  Added event
  ' Raises a flag to start wiping files.
  Public Event WipingFiles(ByVal blnStarted As Boolean)

  ' 20-Jul-2008  Added event
  ' Raises a flag to start wiping folders.
  Public Event WipingFolders(ByVal blnStarted As Boolean)

  ' 20-Aug-2010  Added event
  ' Raises a flag to start wiping free space
  Public Event WipingFreeSpace(ByVal blnStarted As Boolean)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let AlternateMethod(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.     Syntax: X.AlternateMethod = TRUE
    mblnAlternateMethod = blnData
End Property

Public Property Get AlternateMethod() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.        Syntax: Debug.Print X.AlternateMethod
    AlternateMethod = mblnAlternateMethod
End Property

Public Property Let CurrentPattern(ByVal strData As String)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.
    ' Syntax: X.CurrentPattern = "5 writes per sector [ (Random * 4), 0x00 ]"
    mstrCurrentPattern = strData
End Property

Public Property Let DisplayMsgs(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.DisplayMsgs = FALSE
    mblnDisplayMsgs = blnData
End Property

Public Property Let DoFolders(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.     Syntax: X.DoFolders = TRUE
    mblnDoFolders = blnData
End Property

Public Property Get DoFolders() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.        Syntax: Debug.Print X.DoFolders
    DoFolders = mblnDoFolders
End Property

Public Property Let DoSubFolders(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax: X.DoSubFolders = TRUE
    mblnDoSubFolders = blnData
End Property

Public Property Let EncryptAlgo(ByVal lngData As enumEncryptAlgo)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.EncryptAlgo = eWipe_Rijndael
    mlngEncryptAlgo = lngData
End Property

Public Property Let FlashDrive(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.FlashDrive = FALSE
    mblnFlashDrive = blnData
End Property

Public Property Get FlashDrive() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax: Debug.Print X.FlashDrive
    FlashDrive = mblnFlashDrive
End Property

Public Property Let KeepDirStructure(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax: X.mblnKeepDirStructure = TRUE
    mblnKeepDirStructure = blnData
End Property

Public Property Let LogData(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.LogData = True
    mblnLogData = blnData
End Property

Public Property Let LogEncryptParms(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.LogData = True
    mblnLogEncryptParms = blnData
End Property

Public Property Let Passes(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.Passes = 5
    mlngPasses = lngData
End Property

Public Property Let ProtectiveItems(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.ProtectiveItems = 5
    mlngProtectiveItems = lngData
End Property

Public Property Let RemoveTopFolder(ByVal blnData As Boolean)
Attribute RemoveTopFolder.VB_Description = "Flag that designates if the top level folder is also to be deleted."
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax: X.DoSubFolders = TRUE
    mblnRemoveTopFolder = blnData
End Property

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

Public Property Let TypeTarget(ByVal lngData As Long)
Attribute TypeTarget.VB_Description = "Numeric value indicating the type of target to be processed."
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.TypeTarget = 3
    mlngTypeTarget = lngData
End Property

Public Property Get TypeTarget() As Long
Attribute TypeTarget.VB_Description = "Numeric value indicating the type of target to be processed."
    ' used when retrieving value of a property, on the right
    ' side of an assignment.        Syntax: Debug.Print X.TypeTarget
    TypeTarget = mlngTypeTarget
End Property

Public Property Let USB_Drive(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.USB_Drive = FALSE
    mblnUSB_Drive = blnData
End Property

Public Property Get USB_Drive() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax: Debug.Print X.USB_Drive
    USB_Drive = mblnUSB_Drive
End Property

Public Property Let VerifyData(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.VerifyData = FALSE
    mblnVerifyData = blnData
End Property

' Comment out Version Property if not used in this application
Public Property Get Version() As String
    ' Version of the DLL this class is associated with.
    Version = DLL_NAME & ".dll (tm) v" & App.Major & "." & App.Minor & "." & App.Revision & _
              " Copyright (c) 2004-2012 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Let WipeMethod(ByVal lngData As enumWipeMethod)
Attribute WipeMethod.VB_Description = "Numeric value indicating the wiping method to be used."

    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.WipeMethod = WIPE_ENCRYPT
    mlngWipeMethod = lngData
End Property

Public Property Get WipeMethod() As Long
Attribute WipeMethod.VB_Description = "Numeric value indicating the wiping method to be used."
    ' used when retrieving value of a property, on the right
    ' side of an assignment.        Syntax: Debug.Print X.WipeMethod
    WipeMethod = mlngWipeMethod
End Property

Public Property Let WipeFreeSpace(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.     Syntax: X.WipeFreeSpace = TRUE
    mblnWipeFreeSpace = blnData
End Property

Public Property Get WipeFreeSpace() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.        Syntax: Debug.Print X.WipeFreeSpace
    WipeFreeSpace = mblnWipeFreeSpace
End Property

Public Property Let WipePath(ByVal strData As String)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.WipePath = "Wipe drive C:"
    mstrWipePath = strData
End Property

Public Property Get WipePath() As String
    ' used when retrieving value of a property, on the right
    ' side of an assignment.    Syntax: Debug.Print X.WipePath
    WipePath = mstrWipePath
End Property

Public Property Let WipePatterns(ByRef astrData() As String)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.
    ' Syntax: X.WipePatterns = array of data
    Dim lngIndex As Long
    Dim lngMax   As Long
    
    lngMax = UBound(astrData)
    mblnMultiple = False
    
    Erase mastrCustom()        ' Always start with an empty array
    ReDim mastrCustom(lngMax)  ' Size array
    
    For lngIndex = 0 To lngMax - 1
        
        mastrCustom(lngIndex) = astrData(lngIndex)
        
        If StrComp("Multiple", astrData(lngIndex), vbTextCompare) = 0 Then
            mblnMultiple = True
        End If
        
    Next lngIndex
    
End Property

Public Property Let ZeroLastWrite(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.  Syntax: X.ZeroLastWrite = FALSE
    ' Only valid with options 2-7
    mblnZeroLastWrite = blnData
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       BeginProcessing
'
' Description:   Begin processing the request to either wipe a file, files
'                within a folder, or a complete folder and it's sub-foldrs
'
' Parameters:    avntList - an array of folders or file names
'
' Returns:       TRUE - Successful completion
'                FALSE - Something wrong happened
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 12-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added option for wiping a USB drive
' 02-Oct-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added test for possible unicode folder or file names.  By the
'              time this routine is called, the folders or files have been
'              selected via Windows Explorer.
' ***************************************************************************
Public Function BeginProcessing(ByVal avntList As Variant) As Boolean
Attribute BeginProcessing.VB_Description = "Begin processing the request to either wipe a file, files within a folder, or a complete folder and it's subfolders."

    Dim strPath     As String     ' name of folder
    Dim strMsg      As String     ' msgbox message
    Dim strDrive    As String     ' drive letter
    Dim strRecord   As String     ' log file message
    Dim curFileSize As Currency
    Dim blnFinished As Boolean
    Dim objDI       As cDiskInfo

    Const ROUTINE_NAME As String = "BeginProcessing"
    
    ' initialize variables
    RaiseEvent WaitMsg(WAIT_MSG, &HE0E0E0)
    RaiseEvent CurrentPattern(mstrCurrentPattern)
    RaiseEvent WipingFiles(False)
    RaiseEvent WipingFolders(False)
    ResetVariables
    RaiseEvent CountPasses(0, 0)
    RaiseEvent CountTotals("", 0, 0, 0, True, False)

    mblnProcessUSB = False
    mstrTopLevelFolder = vbNullString
    Set mcolFolders = New Collection   ' Instantiate new collection
        
    Rnd -1                             ' Reset VB random number generator
    Randomize mobjRandom.RndSeed       ' Reseed VB random number generator
        
    ' make sure an array was passed
    If IsArrayInitialized(avntList) Then

        strDrive = QualifyPath(Left$(avntList(0), 2))
        
        Select Case mlngTypeTarget

               Case 1   ' Files

                    strPath = avntList(0)  ' Save the first element
                    
                    ' process multiple files if more than one element in array
                    If UBound(avntList) > 1 Then
                    
                        ' If there are problems with identifying
                        ' a target at this point then display an
                        ' appropriate message and leave.
                        If Not IsPathValid(strPath) Then
                            InfoMsg "Cannot identify folder." & vbNewLine & vbNewLine & _
                                    "It may have unreadable characters in the" & vbNewLine & _
                                    "path or is being held by another process." & _
                                    vbNewLine & vbNewLine & strPath & _
                                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                            
                            If mblnLogData Then
                                strRecord = "ABENDED:" & vbTab & "Cannot access:  " & strPath
                                RaiseEvent UpdateLogData(strRecord)
                            End If

                            gblnStopProcessing = True
                        End If

                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo BeginProcessing_CleanUp
                        End If
                        
                        ProcessFolders avntList, True
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo BeginProcessing_CleanUp
                        End If

                    Else
                    
                        If Not IsPathValid(strPath) Then
                    
                            ' If there are problems with identifying
                            ' a target at this point then display an
                            ' appropriate message and leave.
                            InfoMsg "Cannot identify file." & vbNewLine & vbNewLine & _
                                    "It may have unreadable characters in the" & vbNewLine & _
                                    "path or is being held by another process." & _
                                    vbNewLine & vbNewLine & strPath & _
                                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                            
                            If mblnLogData Then
                                strRecord = "ABENDED:" & vbTab & "Cannot access:  " & strPath
                                RaiseEvent UpdateLogData(strRecord)
                            End If

                            gblnStopProcessing = True
                        End If

                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo BeginProcessing_CleanUp
                        End If

                        mobjBigFiles.CalcFileSize strPath, curFileSize   ' get length of the file
                        
                        If curFileSize < 0 Then
                            gblnStopProcessing = True     ' Time to leave
                        Else
                            mcurByteCount = curFileSize   ' Save the data
                            mcurFileCount = 1@            ' file count is 1
                        End If
                        
                    End If

                    If mcurFileCount = 0 And _
                       mcurByteCount = 0 And _
                       mcurFolderCount = 0 Then
                       
                          gblnStopProcessing = True
                    End If

                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo BeginProcessing_CleanUp
                    End If

                    strMsg = "Are you sure you want to wipe " & vbNewLine
                    strMsg = strMsg & strPath & vbNewLine & vbNewLine
                    strMsg = strMsg & "Consisting of:" & vbNewLine
                    strMsg = strMsg & Space$(10) & Format$(mcurFileCount, "#,##0") & " files" & vbNewLine
                    strMsg = strMsg & Space$(10) & "Totaling " & Format$(mcurByteCount, "#,##0") & " bytes"

                    If ResponseMsg(strMsg, vbQuestion Or vbOKCancel Or vbDefaultButton1, PGM_NAME) = vbOK Then

                        RaiseEvent CountTotals("", mcurFileCount, mcurByteCount, 0, True, False)

                        mstrStatistics = Format$(mcurFileCount, "#,##0") & " files (" & _
                                         Format$(mcurByteCount, "#,##0") & " bytes (" & _
                                         DisplayNumber(mcurByteCount) & "))"

                        mcurOA_MaxAmt = mcurByteCount * mlngPasses
                        mcurOA_MaxPerFile = mcurByteCount
                        ResetCounters
                        
                        ' process the selected files within this folder
                        If UBound(avntList) > 1 Then
                            
                            ProcessFolders avntList, False
                        
                            ' An error occurred or user opted to STOP processing
                            DoEvents
                            If gblnStopProcessing Then
                                GoTo BeginProcessing_CleanUp
                            End If
    
                            ' Update log file, if requested
                            If mblnLogData Then
                                strRecord = "FINISHED:" & vbTab & mstrStatistics
                                RaiseEvent UpdateLogData(strRecord)
                            End If

                        Else
                            
                            ' Update log file, if requested
                            If mblnLogData Then
                                strRecord = "STARTED:" & vbTab & CStr(mlngPasses) & _
                                            IIf(mlngPasses = 1, " Pass using ", " Passes using ") & _
                                            mstrCurrentPattern

                                If mblnVerifyData Then
                                    strRecord = strRecord & " with Verfication"
                                End If
                                
                                RaiseEvent UpdateLogData(strRecord)
                            End If
                            
                            mcurFileCount = mcurFileCount + 1
                            WipeFile strPath

                            ' An error occurred or user opted to STOP processing
                            DoEvents
                            If gblnStopProcessing Then
                                GoTo BeginProcessing_CleanUp
                            End If

                            ' Update log file, if requested
                            If mblnLogData Then
                                strRecord = "FINISHED:" & vbTab & mstrStatistics
                                RaiseEvent UpdateLogData(strRecord)
                            End If

                        End If

                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo BeginProcessing_CleanUp
                        End If

                        RaiseEvent OverallProgress(mcurOA_MaxPerFile, mcurOA_MaxPerFile) ' maintain overall totals

                    Else
                        gblnStopProcessing = True
                    End If

               Case 2  ' folders

                    strPath = Left$(avntList(0), 2)
                    mstrTopLevelFolder = UnQualifyPath(avntList(0))

                    If Not IsPathValid(mstrTopLevelFolder) Then
                    
                        ' If there are problems with identifying
                        ' a target at this point then display an
                        ' appropriate message and leave.
                        InfoMsg "Cannot identify folder." & vbNewLine & vbNewLine & _
                                "It may have unreadable characters in the" & vbNewLine & _
                                "path or is being held by another process." & _
                                vbNewLine & vbNewLine & mstrTopLevelFolder & _
                                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                            
                        If mblnLogData Then
                            strRecord = "ABENDED:" & vbTab & "Cannot access:  " & mstrTopLevelFolder
                            RaiseEvent UpdateLogData(strRecord)
                        End If

                        gblnStopProcessing = True
                    End If

                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo BeginProcessing_CleanUp
                    End If

                    ProcessFolders avntList, True
                    
                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo BeginProcessing_CleanUp
                    End If

                    strMsg = "Are you sure you want to wipe " & vbNewLine
                    strMsg = strMsg & avntList(0) & vbNewLine & vbNewLine
                    strMsg = strMsg & "Consisting of:" & vbNewLine
                    
                    If mblnRemoveTopFolder Then
                        strMsg = strMsg & Space$(10) & Format$(mcurFolderCount, "#,##0") & " folders" & vbNewLine
                    Else
                        ' do not delete top level folder, only
                        ' remove the files and subfolders
                        If mcurFolderCount > 1 Then
                            strMsg = strMsg & Space$(10) & Format$(mcurFolderCount - 1, "#,##0") & " folders" & vbNewLine
                        Else
                            strMsg = strMsg & Space$(10) & Format$(mcurFolderCount, "#,##0") & " folders" & vbNewLine
                        End If
                    End If
                    
                    strMsg = strMsg & Space$(10) & Format$(mcurFileCount, "#,##0") & " files" & vbNewLine
                    strMsg = strMsg & Space$(10) & "Totaling " & Format$(mcurByteCount, "#,##0") & " bytes"

                    If ResponseMsg(strMsg, vbQuestion Or vbOKCancel Or vbDefaultButton1, PGM_NAME) = vbOK Then

                        mstrStatistics = Format$(mcurFolderCount, "#,##0") & " folders, " & _
                                         Format$(mcurFileCount, "#,##0") & " files (" & _
                                         Format$(mcurByteCount, "#,##0") & " bytes (" & _
                                         DisplayNumber(mcurByteCount) & "))"

                        RaiseEvent CountTotals("", mcurFileCount, mcurByteCount, 0, True, False)

                        mcurOA_MaxAmt = mcurByteCount * mlngPasses
                        mcurOA_MaxPerFile = mcurByteCount
                        ResetCounters

                        ' process the selected files within this folder
                        ProcessFolders avntList, False

                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo BeginProcessing_CleanUp
                        End If

                        ' Update log file, if requested
                        If mblnLogData Then
                            strRecord = "FINISHED:" & vbTab & mstrStatistics
                            RaiseEvent UpdateLogData(strRecord)
                        End If

                        RaiseEvent OverallProgress(0, 0) ' maintain overall totals
                        CalculateProgress 0, 0, True

                    Else
                        gblnStopProcessing = True
                    End If
               
               Case 3   ' USB drive
               
                    strPath = QualifyPath(Left$(avntList(0), 2))
                    mstrTopLevelFolder = UnQualifyPath(avntList(0))

                    If Not IsPathValid(mstrTopLevelFolder) Then
                    
                        ' If there are problems with identifying
                        ' a target at this point then display an
                        ' appropriate message and leave.
                        InfoMsg "Cannot identify folder." & vbNewLine & vbNewLine & _
                                "It may have unreadable characters in the" & vbNewLine & _
                                "path or is being held by another process." & _
                                vbNewLine & vbNewLine & mstrTopLevelFolder & _
                                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                            
                        If mblnLogData Then
                            strRecord = "ABENDED:" & vbTab & "Cannot access:  " & mstrTopLevelFolder
                            RaiseEvent UpdateLogData(strRecord)
                        End If

                        gblnStopProcessing = True
                    End If

                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo BeginProcessing_CleanUp
                    End If

                    ProcessFolders avntList, True
                    
                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo BeginProcessing_CleanUp
                    End If

                    mcurFolderCount = IIf(mcolFolders.Count > 1, mcolFolders.Count - 1, 0@)
                    
                    Set objDI = New cDiskInfo
                    mcurByteCount = objDI.GetDiskSpaceInfo(strDrive, eUsedSpace)
                    mcurByteCount = mcurByteCount - KB_16
                    Set objDI = Nothing
                    
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo BeginProcessing_CleanUp
                    End If
                    
                    strMsg = "Are you sure you want to wipe USB drive " & vbNewLine
                    strMsg = strMsg & mstrTopLevelFolder & vbNewLine & vbNewLine
                    strMsg = strMsg & "Consisting of:" & vbNewLine
                    strMsg = strMsg & Space$(10) & Format$(mcurFolderCount, "#,##0") & " root level folders" & vbNewLine
                    strMsg = strMsg & Space$(10) & Format$(mcurFileCount, "#,##0") & " root level files" & vbNewLine
                    strMsg = strMsg & Space$(10) & "Totaling " & Format$(mcurByteCount, "#,##0") & " bytes" & vbNewLine
                    strMsg = strMsg & Space$(20) & "including subfolder files"

                    If ResponseMsg(strMsg, vbQuestion Or vbOKCancel Or vbDefaultButton1, PGM_NAME) = vbOK Then

                        mstrStatistics = Format$(mcurFolderCount, "#,##0") & " folders, " & _
                                         Format$(mcurFileCount, "#,##0") & " files (" & _
                                         Format$(mcurByteCount, "#,##0") & " bytes (" & _
                                         DisplayNumber(mcurByteCount) & "))"

                        RaiseEvent CountTotals("", mcurFileCount, mcurByteCount, 0, True, False)

                        ' proceed to overwrite USB drive
                        If Not ProcessUSB(strPath) Then
                            gblnStopProcessing = True
                            GoTo BeginProcessing_CleanUp
                        End If

                        ' Update log file, if requested
                        If mblnLogData Then
                            strRecord = "FINISHED:" & vbTab & mstrStatistics
                            RaiseEvent UpdateLogData(strRecord)
                        End If

                        RaiseEvent OverallProgress(mcurOA_MaxPerFile, mcurOA_MaxPerFile) ' maintain overall totals

                    Else
                        gblnStopProcessing = True
                    End If
               
               Case Else  ' Safety net
                    gblnStopProcessing = True
        End Select
    Else
        gblnStopProcessing = True
    End If

BeginProcessing_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        blnFinished = False
    Else
        blnFinished = True
    End If

    EmptyCollection mcolFolders       ' Empty collection object
    Set objDI = Nothing               ' Free objects from memory
    CloseAllFiles                     ' Close any opened files
    RaiseEvent WaitMsg("", &HE0E0E0)  ' Reset to gray background
    
    If mblnLogData Then
        If Not blnFinished Then
            strRecord = "ABENDED:" & vbTab & "Processing aborted"
            RaiseEvent UpdateLogData(strRecord)
        End If
    End If

    BeginProcessing = blnFinished

End Function

' **************************************************************************
' Routine:       WipeTheFreeSpace
'
' Description:   Creates a temp folder in the root directory of the drive
'                if one does not exist.  Within this folder, a series of
'                hidden files are created and filled with null (0x00)
'                values.  Maximum size of each file is 1GB.  This is
'                designed for speed.  Visual Basic's Get() and PUT()
'                functions are limited to long integers and will not work
'                on files greater than 2GB.  When finished, the folder and
'                files within are emptied and deleted.
'
'                Because this is so fast, I have decided to make a minimum
'                of three (3) passes on each drive.  This can be raised,
'                but not lowered, by changing a value on the option screen.
'
'                The exception being a flash drive.  The number of passes
'                will be set to one (1) pass only because of the lack of
'                overlapping magnetic write areas as found on a fixed disk.
'
' Parameters:    strDrive - Drive to have free space overwritten
'
' Returns:       TRUE - Successful completion
'                FALSE - something went wrong
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 03-May-2008  Kenneth Ives  kenaso@tx.rr.com
'              Set the write process to create a series of files filled with
'              null values.
' 11-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine.  Now calls FillWithTempFiles() to fill drive
'              space with binary zero files.
' 12-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added capability to toggle between three passes with blank
'              spaces or one pass using an alternate method
' 24-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Created property named AlternateMethod to determine if to use
'              an alternate method to wipe free space.
'              Fixed bug to keep current progress bar empty when not using
'              an alternate method for wiping.
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated WipeTheFreeSpace() routine to display properly
'              when wiping freespace with the alternate pattern.
' ***************************************************************************
Public Function WipeTheFreeSpace(ByVal strDrive As String) As Boolean

    Dim lngPass     As Long      ' pass counter
    Dim lngPassesH  As Long      ' Temp hold of pass counter
    Dim lngMethodH  As Long      ' Temp hold of wipe method
    Dim lngDelay    As Long      ' Delay for display purposes only
    Dim strPath     As String    ' temp Path
    Dim strRecord   As String    ' log record
    Dim astrFiles() As String    ' array of temp file names (Alternate method only)
    
    Const ROUTINE_NAME As String = "WipeTheFreeSpace"
    Const WIPE_MSG     As String = "Wipe free space on drive "

    On Error GoTo WipeTheFreeSpace_Error

    DoEvents
    WipeTheFreeSpace = False           ' Preset flag to FALSE
    RaiseEvent WipingFreeSpace(True)
    DoEvents
    
    Erase astrFiles()                  ' Always start with empty arrays
    ResetVariables                     ' reset accumulation variables
    strDrive = QualifyPath(strDrive)   ' add a backslash, if needed
    strPath = strDrive & TEMP_FOLDER   ' Format the path
    lngPass = 1                        ' Init first pass count
    
    If mblnLogData Then
        strRecord = "STARTED:" & vbTab & WIPE_MSG & strDrive
        RaiseEvent UpdateLogData(strRecord)
    End If
    
    RaiseEvent WaitMsg(WIPE_MSG & strDrive, &HE0E0E0)
    
    ' See if accessing a flash drive or
    ' user opted to use US DoD Short pattern
    If mblnAlternateMethod Then
        lngPassesH = mlngPasses   ' Save current number of passes
        Passes = 1
        mstrCurrentPattern = "DoD Short  [ Ignore space messages ]"
    Else
        mstrCurrentPattern = FREE_SPACE_PATTERN
    End If
        
    DoEvents
    RaiseEvent CurrentPattern(mstrCurrentPattern)  ' Update pattern being used
    RaiseEvent CountPasses(lngPass, mlngPasses)    ' update number of passes
    RaiseEvent CurrentProgress(0, 100)             ' Do not show current progress
    CalculateProgress 0, 100, False                ' Do not show current progress
    DoEvents
    
    mlngStartTime = GetTickCount  ' Setting the timers
    mlngPause = GetTickCount      ' Used in CalculateProgress()
          
    ' Fill DD_Temp folder with null filled files
    FillWithTempFiles strDrive, astrFiles()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo WipeTheFreeSpace_CleanUp
    End If
    
    ' See if user wants to wipe free
    ' space using US DoD Short algorithm
    If mblnAlternateMethod Then
        
        mobjRandom.EmptyCollection mcolFolders
        Set mcolFolders = New Collection
        
        ' Temporarily reset wipe method to
        ' use the alternate wiping pattern
        lngMethodH = mlngWipeMethod          ' Save current wipe method
        WipeMethod = eWipe_AlternateMethod   ' update property with alternate wipe method
        
        If mblnLogData Then
            strRecord = vbTab & vbTab & "Using pattern " & mstrCurrentPattern
            RaiseEvent UpdateLogData(strRecord)
        End If
    
        ' Reset display counters
        ResetCounters
        mblnDoFolders = True
'        mblnRemoveTopFolder = True
        mstrTopLevelFolder = QualifyPath(strPath)
        
        ' Get an accurate count of the
        ' temp files to be overwritten
        ProcessFolders astrFiles(), True
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo WipeTheFreeSpace_CleanUp
        End If
    
        RaiseEvent CountTotals("", mcurFileCount, mcurByteCount, 0, True, False)
        
        ' Calculate progress display
        mcurOA_MaxAmt = mcurByteCount * mlngPasses
        mcurOA_MaxPerFile = mcurByteCount
        
        ' Reset display counters
        ResetCounters

        ' process the temp files within the DD_Temp folder
        ProcessFolders astrFiles(), False

    Else
        
        Erase astrFiles()   ' Always empty arrays when not needed
        
        If mblnLogData Then
            strRecord = vbTab & vbTab & "Using pattern " & mstrCurrentPattern
            RaiseEvent UpdateLogData(strRecord)
        End If
    
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo WipeTheFreeSpace_CleanUp
    End If
    
    ' See if there is a DD_Temp folder
    If IsPathValid(strPath) Then
        FolderCleanUp strPath, False    ' Delete folder and contents
    End If
            
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo WipeTheFreeSpace_CleanUp
    End If
    
    ' update the log file, if requested
    If mblnLogData Then
        strRecord = vbTab & vbTab & "Removed " & Format$(mcurFileCount, "#,##0") & _
                    " temp files from drive " & strDrive
        RaiseEvent UpdateLogData(strRecord)
    End If
    
    ' Slight delay for screen viewing only
    lngDelay = GetTickCount + 500
    Do While lngDelay > GetTickCount
        DoEvents
    Loop
            
    If mblnAlternateMethod Or mblnFlashDrive Then
        ' If opting to use the alternate method
        ' or this is a flash drive then only
        ' one pass will be used.
        mobjRandom.EmptyCollection mcolFolders
    Else
    
        ' Calc ovarall total for progressbar
        mcurOA_MaxAmt = mcurOA_MaxPerFile * mlngPasses
    
        For lngPass = 2 To mlngPasses
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
            
            DoEvents
            RaiseEvent CountPasses(lngPass, mlngPasses)   ' update number of passes on progress screen
            DoEvents
                
            ' update the log file, if requested
            If mblnLogData Then
                strRecord = vbTab & vbTab & "Pass " & CStr(lngPass) & _
                            " of " & CStr(mlngPasses) & _
                            " using pattern " & mstrCurrentPattern
                RaiseEvent UpdateLogData(strRecord)
            End If
            
            ' Fill DD_Temp folder with null filled files
            If Not FillWithTempFiles(strDrive, astrFiles()) Then
                Exit For    ' exit For..Next loop
            End If
    
            Erase astrFiles()   ' Always empty arrays when not needed
    
            ' See if there is a DD_Temp folder
            If IsPathValid(strPath) Then
                FolderCleanUp strPath, False    ' Delete folder and contents
            End If
                    
            ' update the log file, if requested
            If mblnLogData Then
                strRecord = vbTab & vbTab & "Removed " & Format$(mcurFileCount, "#,##0") & _
                            " temp files from drive " & strDrive
                RaiseEvent UpdateLogData(strRecord)
            End If
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            RaiseEvent WaitMsg(WIPE_MSG & strDrive & RE_EVAL_DR, &HE0E0E0)
    
            ' Slight delay for screen viewing only
            lngDelay = GetTickCount + 500
            Do While lngDelay > GetTickCount
                DoEvents
            Loop
            
        Next lngPass
    
    End If
    
WipeTheFreeSpace_CleanUp:
    DoEvents
    RaiseEvent WaitMsg(CLEANUP_MSG, vbYellow)  ' Display message in yellow
    Erase astrFiles()                          ' Empty arrays when not needed
    
    DoEvents
    mobjRandom.EmptyCollection mcolFolders  ' Verify collectionis empty
    FolderCleanUp strPath, True             ' Remove DD_Temp folder
    
    DoEvents
    RaiseEvent CurrentProgress(0, 100)
    CalculateProgress 0, 100, False
    RaiseEvent WipingFreeSpace(False)   ' Hide disc scrubbing image
    DoEvents
    
    mstrStatistics = "Amount overwritten:  " & _
                     Format$(mcurByteCount, "#,##0") & " bytes (" & _
                     DisplayNumber(mcurByteCount) & ")"
    
    If gblnStopProcessing Then
        
        ' if logging data then reformat msg
        If mblnLogData Then
            strRecord = "ABENDED:" & vbTab & WIPE_MSG & strDrive & vbNewLine & _
                        Space$(32) & ABEND_MSG
            RaiseEvent UpdateLogData(strRecord)
        End If
        
        strRecord = "ABENDED:  " & WIPE_MSG & strDrive & vbNewLine & ABEND_MSG & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        
        InfoMsg strRecord
        WipeTheFreeSpace = False
    Else
        strRecord = "FINISHED:" & vbTab & WIPE_MSG & strDrive & _
                    Space$(2) & mstrStatistics
        
        ' Update log file, if requested
        If mblnLogData Then
            RaiseEvent UpdateLogData(strRecord)
        End If
    
        WipeTheFreeSpace = True
    End If

    ' update the log file, if requested
    If mblnLogData Then
        RaiseEvent UpdateLogData(strRecord)
    End If

    ' Reset original values
    If mblnFlashDrive Then
        Passes = lngPassesH
    End If
    
    ' Reset pass count and wipe method
    ' back to its original values
    If mblnAlternateMethod Then
        Passes = lngPassesH
        WipeMethod = lngMethodH
    End If
    
    DoEvents
    mobjRandom.EmptyCollection mcolFolders  ' empty collection
    RaiseEvent WaitMsg("", &HE0E0E0)        ' Reset to gray background
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

WipeTheFreeSpace_Error:
    Err.Clear
    DoEvents
    gblnStopProcessing = True
    Resume WipeTheFreeSpace_CleanUp

End Function

Public Sub FolderCleanUp(ByVal strPath As String, _
                Optional ByVal blnRemoveFolder As Boolean = False)
                    
    ' NOTE:  This code requires that you add a reference to the
    '        Microsoft Scripting Runtime type library
    
    Dim hFile         As Long
    Dim strDummy      As String
    Dim strTempFolder As String   ' System Temp folder path
    Dim objFSO        As Scripting.FileSystemObject
    
    CloseAllFiles  ' close any open files
    
    ' is this the root of a drive?
    If IsPathARoot(strPath) Then
    
        ' If this is the root folder of a
        ' drive, it will not be deleted!
        Exit Sub
    End If
    
    SetFileAttributes strPath, FILE_ATTRIBUTE_NORMAL  ' Reset folder attributes
    strPath = UnQualifyPath(strPath)                  ' Remove trailing backslash
    
    ' Release the grip on any currently
    ' accessed files by creating an empty
    ' temp file in the system temp folder
    strDummy = CreateTempFile()
    hFile = FreeFile
    Open strDummy For Output As #hFile
    Close #hFile

    ' Save path for later cleanup
    strTempFolder = QualifyPath(GetPath(strDummy))
 
    ' If complete folder needs to be removed
    If blnRemoveFolder Then
        
        If IsPathValid(strPath) Then
            On Error Resume Next
            Set objFSO = New Scripting.FileSystemObject  ' Instantiate scripting object
            DoEvents
            objFSO.DeleteFolder strPath, True            ' Remove folder and its contents
            DoEvents
            Wait 3000                                    ' Three second delay
            Set objFSO = Nothing                         ' Free scripting object from memory
            On Error GoTo 0                              ' Nullify this error trap
        End If
        
    Else
        FilesAndFolders_Delete strPath, "*.tmp"   ' Remove only folder contents
        Wait 2000                                 ' Two second delay
    End If
            
    ' Remove all files in Windows temp folder prefixed
    ' with "~ki"
    '
    ' If running in the VB IDE, sometimes the temp work
    ' folders will not be removed because the IDE itself
    ' is still recognizing them.
    FilesAndFolders_Delete strTempFolder, TMP_FILE_PREFIX & "*.tmp"
    
    Wait 2000   ' Two second delay
    
End Sub



' ***************************************************************************
' ****                Internal Procedures and Functions                  ****
' ***************************************************************************

' ***************************************************************************
' Routine:       ProcessFolders
'
' Description:   Set up the starting data, Begin the search.
'
' Parameters:    strFolder - Start the search at this folder
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch  http://vbnet.mvps.org/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' 12-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added reference for wiping a USB drive
' 02-Oct-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added test for possible unicode folder or file names
' ***************************************************************************
Private Sub ProcessFolders(ByVal vntFileList As Variant, _
                           ByVal blnCountOnly As Boolean)

    Dim lngIndex    As Long
    Dim curFileSize As Currency  ' File size
    Dim strPath     As String    ' Path to TEMP folder
    Dim strRecord   As String    ' log file message
    Dim strBase     As String

    Const ROUTINE_NAME As String = "ProcessFolders"
    
    On Error Resume Next         ' ignore any errors

    ReDim mtypFP.sFileList(UBound(vntFileList))
    
    ' Initialize file parameter structure
    With mtypFP
         .sFileRoot = QualifyPath(vntFileList(0))  ' start path
         .sFileList() = vntFileList                ' Array of files
    End With

    ' Update log file, if requested
    If Not mblnAlternateMethod Then
        If mblnLogData Then
            If Not blnCountOnly Then
                strRecord = "STARTED:" & vbTab & CStr(mlngPasses) & _
                            IIf(mlngPasses = 1, " Pass using ", " Pass(es) using ") & mstrCurrentPattern
    
                If mblnVerifyData Then
                    strRecord = strRecord & " with Verfication"
                End If
    
                RaiseEvent UpdateLogData(strRecord)
            End If
        End If
    End If
    
    strPath = GetLongName(mtypFP.sFileRoot)  ' get long path name
    strPath = QualifyPath(strPath)           ' Add trailing backslash

    Select Case mlngTypeTarget
           
           Case 1
                If blnCountOnly Then
                    
                    For lngIndex = 1 To UBound(mtypFP.sFileList) - 1
                        
                        If IsPathValid(strPath & mtypFP.sFileList(lngIndex)) Then
                            
                            mobjBigFiles.CalcFileSize strPath & mtypFP.sFileList(lngIndex), curFileSize
                            
                            If curFileSize < 0 Then
                                gblnStopProcessing = True     ' Time to leave
                            Else
                                mcurByteCount = mcurByteCount + curFileSize
                                mcurFileCount = mcurFileCount + 1
                            
                                ' Count the data only.  Do not wipe anything.
                                RaiseEvent CountTotals(strPath, mcurFileCount, mcurByteCount, 0, True, False)
                            
                            End If
                        
                        Else
                
                            ' If there are problems with identifying
                            ' a target at this point then display an
                            ' appropriate message and leave.
                            InfoMsg "Cannot identify file." & vbNewLine & vbNewLine & _
                                    "It may have unreadable characters in the" & vbNewLine & _
                                    "path or is being held by another process." & vbNewLine & vbNewLine & _
                                    strPath & mtypFP.sFileList(lngIndex) & _
                                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME

                        
                            If mblnLogData Then
                                strRecord = "ABENDED" & vbTab & "Cannot access:  " & _
                                            strPath & mtypFP.sFileList(lngIndex)
                                RaiseEvent UpdateLogData(strRecord)
                            End If
                        
                            gblnStopProcessing = True
                        End If
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
    
                    Next lngIndex
                
                Else
                    
                    For lngIndex = 1 To UBound(mtypFP.sFileList) - 1
                        
                        If IsPathValid(strPath & mtypFP.sFileList(lngIndex)) Then
    
                            mcurFileCount = mcurFileCount + 1
                            RaiseEvent CountTotals(strPath, mcurFileCount, mcurByteCount, 0, False, False)
    
                            ' Delete the file
                            WipeFile strPath & mtypFP.sFileList(lngIndex)
                            DoEvents
                        
                        Else
                
                            ' If there are problems with identifying
                            ' a target at this point then display an
                            ' appropriate message and leave.
                            InfoMsg "Cannot identify file." & vbNewLine & vbNewLine & _
                                    "It may have unreadable characters in the" & vbNewLine & _
                                    "path or is being held by another process." & vbNewLine & vbNewLine & _
                                    strPath & mtypFP.sFileList(lngIndex) & _
                                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                        
                            If mblnLogData Then
                                strRecord = "ABENDED:" & vbTab & "Cannot access:  " & _
                                            strPath & mtypFP.sFileList(lngIndex)
                                RaiseEvent UpdateLogData(strRecord)
                            End If
                        
                            gblnStopProcessing = True
                        End If
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
    
                    Next lngIndex
                
                End If

       Case Else
            If blnCountOnly Then
                
                mcurFolderCount = mcurFolderCount + 1  ' increment folder count
                mcolFolders.Add strPath                ' Save the path in the collection
                FilesAndFolders_Count strPath          ' count the number of files and folders
                            
                If Not mblnRemoveTopFolder Then
                    mcurFolderCount = mcurFolderCount - 1  ' decrement folder count
                End If
                
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    GoTo ProcessFolders_CleanUp
                End If
    
            Else
                If mlngWipeMethod <> eWipe_AlternateMethod Then
                    RaiseEvent WipingFiles(True)   ' Set files animation to TRUE
                End If
                
                ' Wipe out all files in the
                ' directory structure first
                FilesAndFolders_Wipe strPath
                
                If mlngWipeMethod <> eWipe_AlternateMethod Then
                    RaiseEvent WipingFiles(False)   ' Set files animation to FALSE
                End If
                
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    GoTo ProcessFolders_CleanUp
                End If
    
                ' See if removing files only and
                ' keeping the folder structure
                If mblnKeepDirStructure Then
                    GoTo ProcessFolders_CleanUp
                End If
                
                ' See if wiping a USB drive
                If Not mblnUSB_Drive Then
                    
                    ' Clear the entries for the
                    ' individual progress and file sizes
                    CalculateProgress 0, 0, False
                    Wait 1000   ' Pause
                    mcurFolderCount = mcolFolders.Count
                    
                    If mcurFolderCount = 1 And Not mblnRemoveTopFolder Then
                        GoTo ProcessFolders_CleanUp
                    End If
                    
                    strBase = "Wiping folder - "     ' Set up display prefix
                    RaiseEvent WipingFolders(True)   ' Set folder animation to TRUE
                    
                    ' Multiple folders to process
                    If (mcurFolderCount > 1) And mblnDoSubFolders Then
                        
                        ' See if the top level
                        ' folder is to be removed
                        If mblnRemoveTopFolder Then
                            mcurFolderCount = mcolFolders.Count     ' Include top level
                        Else
                            mcurFolderCount = mcolFolders.Count - 1 ' Just the subfolders
                        End If
                        
                        ' Move folder count to the display.  Will not show at this time.
                        RaiseEvent CountTotals("", mcurFolderCount, 0, 0, True, False)
                        
                        SortFolderList         ' sort list of folder names
                        mcurFolderCount = 0@   ' reset folder counter
                        
                        ' Display the folder count.  (Ex: Number of items:  0 of 99 )
                        RaiseEvent CountTotals("Wiping folders", mcurFolderCount, 0, 0, False, False)
                        Wait 1000
                        
                        ' Loop thru folder collection and
                        ' delete folder structures
                        For lngIndex = mcolFolders.Count To 1 Step -1
    
                            If lngIndex = 1 Then
                                If mblnRemoveTopFolder Then
                            
                                    strPath = mcolFolders.Item(lngIndex)            ' Get path name
                                    strRecord = strBase & ShrinkToFit(strPath, 40)  ' format display record
                                    mcurFolderCount = mcurFolderCount + 1           ' increment folder count
                                    
                                    ' Display folder count of current folder.   (Ex: Number of items:  1 of 99 )
                                    RaiseEvent CountTotals(strRecord, mcurFolderCount, 0, 0, False, False)
                                            
                                    mcolFolders.Remove lngIndex   ' remove name from collection
                                    WipeFolder strPath            ' delete the folder
                                End If
                            Else
                                strPath = mcolFolders.Item(lngIndex)            ' Get path name
                                strRecord = strBase & ShrinkToFit(strPath, 40)  ' format display record
                                mcurFolderCount = mcurFolderCount + 1           ' increment folder count
                                
                                ' Display folder count of current folder.   (Ex: Number of items:  1 of 99 )
                                RaiseEvent CountTotals(strRecord, mcurFolderCount, 0, 0, False, False)
                                        
                                mcolFolders.Remove lngIndex   ' remove name from collection
                                WipeFolder strPath            ' delete the folder
                            End If
                            
                            ' An error occurred or user opted to STOP processing
                            DoEvents
                            If gblnStopProcessing Then
                                Exit For    ' exit For..Next loop
                            End If
    
                        Next lngIndex
                    Else
                        ' only one folder to process
                        If mblnDoFolders Then
                            
                            If mcurFolderCount = 1 And Not mblnRemoveTopFolder Then
                                GoTo ProcessFolders_CleanUp
                            End If
                    
                            ' Move folder count to the display.
                            RaiseEvent CountTotals("Wiping folders", mcurFolderCount, 0, 0, True, False)
                        
                            ' Display the folder count.  (Ex: Number of items:  0 of 1 )
                            RaiseEvent CountTotals("Wiping folders", mcurFolderCount, 0, 0, False, False)
                            Wait 1000
                        
                            strRecord = strBase & ShrinkToFit(strPath, 40)  ' format display record
                            
                            ' Display folder count of current folder.   (Ex: Number of items:  1 of 1 )
                            RaiseEvent CountTotals(strRecord, mcurFolderCount, 0, 0, False, False)
                                    
                            WipeFolder strPath   ' delete the folder
                            
                            ' An error occurred or user opted to STOP processing
                            DoEvents
                            If gblnStopProcessing Then
                                GoTo ProcessFolders_CleanUp
                            End If
    
                            RaiseEvent CountTotals(strRecord, mcurFolderCount, 0, 0, False, False)
                        End If
                    End If
                    
                    RaiseEvent WipingFolders(False)   ' Set folder animation to FALSE
                    Wait 1000
                    
                End If
            End If
    End Select
    
ProcessFolders_CleanUp:
    On Error GoTo 0   ' Nullify this error trap

End Sub

' ***************************************************************************
' Routine:       FilesAndFolders_Count
'
' Description:   This routine is for counting all files in a given folder
'                (*.*), or for determining hidden or system files, or
'                directories on a given path.  Since this is a dedicated
'                routine, it is faster.
'
' Parameters:    strPath - Current folder being processed.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch  http://vbnet.mvps.org/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Private Sub FilesAndFolders_Count(ByRef strPath As String)

    Dim strPathFile  As String           ' current folder or filename
    Dim strSubFolder As String           ' current folder or filename
    Dim hFile        As Long             ' folder or file handle
    Dim curFileSize  As Currency         ' File size
    Dim typWFD       As WIN32_FIND_DATA  ' Folder or file data structure

    On Error Resume Next   ' ignore all errors

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    typWFD.cFilename = vbNullString     ' Verify variables are empty
    strPathFile = vbNullString

    strPath = QualifyPath(strPath)      ' Add a trailing backslash if missing
    strPathFile = strPath & "*.*"       ' set full search parameters
    hFile = FindFirstFile(strPathFile, typWFD)  ' Make API call

    ' If we have a valid handle then continue processing
    If hFile <> INVALID_HANDLE_VALUE Then
        Do
            ' Collect all the folder names
            If (typWFD.dwFileAttributes And vbDirectory) Then

                strSubFolder = TrimStr(typWFD.cFilename)

                ' make sure this is not a VTOC identifier
                If strSubFolder <> "." And strSubFolder <> ".." Then
                     
                    mcurFolderCount = mcurFolderCount + 1  ' increment folder count
                    
                    If mblnUSB_Drive Then
                        ' need root level folder names only
                        mcolFolders.Add strPath & strSubFolder
                    End If
                    
                    ' process all subfolders
                    If mblnDoSubFolders Then
                        
                        ' add path to the collection
                        mcolFolders.Add strPath & strSubFolder
                        
                        ' Count the data only.  Do not wipe anything.
                        RaiseEvent CountTotals(strPath, mcurFileCount, mcurByteCount, 0, True, False)

                        FilesAndFolders_Count strPath & strSubFolder
                    End If
                End If
            
            Else
                
                ' This must be a file.
                strPathFile = strPath & TrimStr(typWFD.cFilename)
                    
                ' Calc file size
                mobjBigFiles.CalcFileSize strPathFile, curFileSize
                                
                If curFileSize < 0 Then
                    gblnStopProcessing = True     ' Time to leave
                Else
                    mcurByteCount = mcurByteCount + curFileSize
                    mcurFileCount = mcurFileCount + 1
                
                    ' Count the data only.  Do not wipe anything.
                    RaiseEvent CountTotals(strPath, mcurFileCount, mcurByteCount, 0, True, False)
                End If

            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

            typWFD.cFilename = vbNullString     ' Verify variables are empty
            strPathFile = vbNullString

        Loop While FindNextFile(hFile, typWFD)  ' Make API call

    End If

CleanUp:
    FindClose hFile   ' always close file handles when not in use
    On Error GoTo 0   ' Nullify this error trap

End Sub

' ***************************************************************************
' Routine:       FilesAndFolders_Wipe
'
' Description:   This routine Parses a path and processes the selected
'                files and folders.  Since this is a dedicated routine,
'                it is faster.
'
' Parameters:    strPath - Current folder being processed.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch  http://vbnet.mvps.org/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' 12-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added reference for wiping a USB drive
' ***************************************************************************
Private Sub FilesAndFolders_Wipe(ByVal strPath As String)

    Dim strPathFile  As String           ' current folder or filename
    Dim strSubFolder As String           ' current folder or filename
    Dim hFile        As Long             ' folder or file handle
    Dim curFileSize  As Currency         ' File size
    Dim typWFD       As WIN32_FIND_DATA  ' Folder or file data structure

    On Error Resume Next   ' ignore all errors

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    typWFD.cFilename = vbNullString     ' Verify variables are empty
    strPathFile = vbNullString

    strPath = QualifyPath(strPath)    ' Add a trailing backslash if missing
    strPathFile = strPath & "*.*"
    hFile = FindFirstFile(strPathFile, typWFD)  ' Make API call

    ' If we have a valid handle then continue processing
    If hFile <> INVALID_HANDLE_VALUE Then
        Do
            ' Collect all the folder names
            If (typWFD.dwFileAttributes And vbDirectory) Then

                strSubFolder = TrimStr(typWFD.cFilename)
                
                ' make sure this is not a VTOC identifier
                If strSubFolder <> "." And strSubFolder <> ".." Then
                    
                    ' process all subfolders
                    If mblnDoSubFolders Then

                        mcurFolderCount = mcurFolderCount + 1    ' update folder count
                        
                        If Not mblnProcessUSB Then
                            RaiseEvent CountTotals(strPath, mcurFileCount, mcurByteCount, 0, False, False)
                        End If
                        
                        FilesAndFolders_Wipe strPath & strSubFolder

                    End If
                End If
            Else
                ' This must be a file.
                strPathFile = strPath & TrimStr(typWFD.cFilename)
                
                ' Calc file size
                mobjBigFiles.CalcFileSize strPathFile, curFileSize
                                
                mcurByteCount = mcurByteCount + curFileSize
                mcurFileCount = mcurFileCount + 1
                        
                If Not mblnProcessUSB Then
                    RaiseEvent CountTotals(strPathFile, mcurFileCount, 0, curFileSize, False, False)
                End If
                
                WipeFile strPathFile   ' delete the file

            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

            typWFD.cFilename = vbNullString     ' Verify variables are empty
            strPathFile = vbNullString

        Loop While FindNextFile(hFile, typWFD)  ' Make API call

    End If

    FindClose hFile     ' always close file handles when not in use

CleanUp:
    On Error GoTo 0   ' Nullify this error trap

End Sub

' ***************************************************************************
' Routine:       FilesAndFolders_Delete
'
' Description:   This routine is for deleting all files in a given folder
'                (*.*), or for determining hidden or system files, or
'                directories on a given path.  Since this is a dedicated
'                routine, it is faster.
'
' Parameters:    strPath - Current folder being processed.
'                strCriteria - [Optional] Type of file to search for.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch  http://vbnet.mvps.org/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' 12-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added reference for wiping a USB drive
' ***************************************************************************
Private Sub FilesAndFolders_Delete(ByVal strPath As String, _
                          Optional ByVal strCriteria As String = vbNullString)

    Dim strFile      As String           ' current file name
    Dim strDrive     As String           ' drive letter
    Dim strPathFile  As String           ' current folder or filename
    Dim strSubFolder As String           ' current folder or filename
    Dim hFile        As Long             ' folder or file handle
    Dim curFileSize  As Currency
    Dim typWFD       As WIN32_FIND_DATA  ' Folder or file data structure
    
    On Error Resume Next

    typWFD.cFilename = vbNullString   ' Verify variables are empty
    strPathFile = vbNullString

    strPath = QualifyPath(strPath)    ' Add trailing backslashes
    
    If mblnProcessUSB Then
        strDrive = QualifyPath(Left$(strPath, 2))
    End If
    
    ' set full search parameters
    If Len(strCriteria) = 0 Then
        strPathFile = strPath & "*.*"         ' All files
    Else
        strPathFile = strPath & strCriteria   ' specific files
    End If
    
    hFile = FindFirstFile(strPathFile, typWFD)  ' Make API call

    ' If we have a valid handle then continue processing
    If hFile <> INVALID_HANDLE_VALUE Then
        Do
            curFileSize = 0@
            
            ' Collect all the folder names
            If (typWFD.dwFileAttributes And vbDirectory) Then

                strSubFolder = TrimStr(typWFD.cFilename)

                ' make sure this is not a VTOC identifier
                If strSubFolder <> "." And strSubFolder <> ".." Then
                    
                    ' Make a recursive call for a subfolder
                    FilesAndFolders_Delete strPath & strSubFolder, strCriteria
                End If
                
            Else
                ' This must be a file.
                strFile = TrimStr(typWFD.cFilename)
                strPathFile = strPath & strFile
                
                ' Looking for a specific type of file
                If Len(Trim$(strCriteria)) > 0 Then
                     
                    ' Does this file meet search criteria?
                    If strFile Like strCriteria Then
                         
                        DoEvents
                        CloseHandle hFile  ' close API file handle
                        
                        ' Reset file attributes
                        SetFileAttributes strPathFile, FILE_ATTRIBUTE_NORMAL
                        DeleteFile strPathFile  ' Delete file
                        
                    End If
                Else
                    
                    DoEvents
                    CloseHandle hFile  ' close API file handle
                    
                    ' Reset file attributes
                    SetFileAttributes strPathFile, FILE_ATTRIBUTE_NORMAL
                                         
                    ' get file size
                    If mblnProcessUSB Then
                        mobjBigFiles.CalcFileSize strPathFile, curFileSize
                    End If
                    
                    DeleteFile strPathFile  ' Delete file
                                        
                End If
                
            End If

            DoEvents
            If mblnProcessUSB Then
                mcurCurrPerPass = mcurCurrPerPass + curFileSize                  ' Calc amount deleted
                RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxPerFile)   ' maintain overall totals
                CalculateProgress mcurCurrPerPass, mcurOA_MaxAmt, True           ' overall progress
            End If
                
            typWFD.cFilename = vbNullString     ' Verify variables are empty
            strPathFile = vbNullString

        Loop While FindNextFile(hFile, typWFD)  ' Make API call

    End If

    FindClose hFile     ' always close file handles when not in use

CleanUp:
    On Error GoTo 0   ' Nullify this error trap

End Sub

' **************************************************************************
' Routine:       WipeFile
'
' Description:   First determines which wipe method has been selected.  The
'                file's contents are completely overwritten.  It is then
'                emptied and renamed 13 times.  Finally, it is renamed to a
'                single character.  The date/time stamps for the file are
'                updated to 01-01-1980 12:00:00 PM  The file is finally
'                deleted.
'
' Parameters:    strFilename - Path/filename to be wiped
'
' Returns:       NONE
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 03-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to update file date properties with a
'              randomized timestamp
' 12-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added reference for wiping a USB drive
' 12-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added reference for wiping free space with an alternate method
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added test for possible unicode folder or file names
' 02-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added four encryption options.
' ***************************************************************************
Private Sub WipeFile(ByVal strFileName As String)

    Dim hFile         As Long       ' free file handle
    Dim lngIndex      As Long       ' Loop counter
    Dim curFileSize   As Currency   ' original size of the file
    Dim strMsg        As String     ' Displayed message
    Dim strDrive      As String
    Dim strDummy      As String
    Dim strRecord     As String     ' Log file message
    Dim strNewName    As String     ' Unique filename
    Dim strTempFolder As String   ' System Temp folder path
    
    Const ROUTINE_NAME As String = "WipeFile"

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    On Error GoTo WipeFile_Error
    
    mobjBigFiles.CalcFileSize strFileName, curFileSize    ' Get the size of the file
     
    If mblnProcessUSB Then
        strDrive = QualifyPath(Left$(strFileName, 2))
        strMsg = "Wipe USB drive" & strDrive
        RaiseEvent CountTotals(strMsg, mcurFileCount, 0, curFileSize, False, False)
    Else
        RaiseEvent CountTotals(strFileName, mcurFileCount, 0, curFileSize, False, False)
    End If
    
    CalculateProgress 0, curFileSize, False   ' individual progress

    If curFileSize > 0 Then
        Err.Clear
        Select Case mlngWipeMethod

               Case eWipe_AlternateMethod, eWipe_BinaryZeroes, _
                    eWipe_RandomData, eWipe_BruceSchneier
                    
                    If Not MiscPatterns(strFileName) Then
                        GoTo WipeFile_CleanUp
                    End If
                    
               Case eWipe_US_DoD_Short, eWipe_US_DoD_Long
                    If Not DoDPatterns(strFileName) Then
                        GoTo WipeFile_CleanUp
                    End If
                    
               Case eWipe_GermanVISTR, eWipe_NATO
                    If Not EuroPatterns(strFileName) Then
                        GoTo WipeFile_CleanUp
                    End If
                    
               Case eWipe_PeterGutmann
                    If Not GutmannPattern(strFileName) Then
                        GoTo WipeFile_CleanUp
                    End If
                    
               Case eWipe_Rijndael, eWipe_Blowfish, eWipe_Twofish, eWipe_ArcFour
                    If Not EncryptPattern(strFileName) Then
                        GoTo WipeFile_CleanUp
                    End If
                    
               Case eWipe_CustomPattern      ' Customized method
                    If Not CustomPattern(strFileName) Then
                        GoTo WipeFile_CleanUp
                    End If
                    
               Case Else    ' Error
                    InfoMsg "Invalid wipe method selected." & _
                            vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                    GoTo WipeFile_CleanUp
        End Select

    End If
        
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo WipeFile_CleanUp
    End If
    
    ' Verify we are not wiping a USB Drive
    If mblnProcessUSB Then
    
        ' Convert the file to a two-byte file length.
        hFile = FreeFile
        Open strFileName For Output As #hFile
        Print #hFile, ""
        Close #hFile
        DoEvents
    
    Else
        If mblnLogData Then
            strRecord = vbTab & vbTab & "File:  " & strFileName & " (" & _
                        Format$(curFileSize, "#,##0") & " bytes)"
            RaiseEvent UpdateLogData(strRecord)
        End If
    
        If curFileSize > 0 Then
    
            ' rename the file 26 times
            strNewName = CreateNewName(strFileName)
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo WipeFile_CleanUp
            End If
    
            ' Convert file to four-byte length
            hFile = FreeFile
            Open strNewName For Output As #hFile
            Print #hFile, vbNewLine
            Close #hFile
            DoEvents
                
        Else
    
            ' Convert file to four-byte length
            hFile = FreeFile
            Open strFileName For Output As #hFile
            Print #hFile, vbNewLine
            Close #hFile
            DoEvents
        
            ' rename the file 26 times
            strNewName = CreateNewName(strFileName)
        
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo WipeFile_CleanUp
            End If
    
        End If
        
        ' reset file properties with
        ' random date and time stamp
        SetPathFileDate strNewName
    
    End If
    
    ' An error occurred or user opted to STOP processing
    ' or the file disappeared
    DoEvents
    If gblnStopProcessing Then
        GoTo WipeFile_CleanUp
    End If

    On Error Resume Next
    
    ' Release the grip on any currently
    ' accessed files by creating an empty
    ' temp file in the windows temp folder
    strDummy = CreateTempFile()
    hFile = FreeFile
    Open strDummy For Output As #hFile
    Close #hFile

    ' Save path for later cleanup
    strTempFolder = QualifyPath(GetPath(strDummy))
 
    ' make three attempts to verify
    ' the file no longer exist
    For lngIndex = 1 To 3
        
        ' See if original file is available
        DoEvents
        If IsPathValid(strFileName) Then
        
            ' delete original file
            If CBool(DeleteFile(strFileName)) Then
                Exit For  ' exit For..Next loop
            End If
        
        Else
            Exit For  ' exit For..Next loop
        End If
    
    Next lngIndex
    
    ' make three attempts to verify
    ' the file no longer exist
    For lngIndex = 1 To 3
        
        ' See if new file is available
        DoEvents
        If IsPathValid(strNewName) Then
        
            ' delete new file
            If CBool(DeleteFile(strNewName)) Then
                Exit For  ' exit For..Next loop
            End If
        
        Else
            Exit For  ' exit For..Next loop
        End If
    
    Next lngIndex
    
    ' Remove all files in Windows temp folder prefixed
    ' with "~ki"
    '
    ' If running in the VB IDE, sometimes the temp work
    ' folders will not be removed because the IDE itself
    ' is still recognizing them.
    FilesAndFolders_Delete strTempFolder, TMP_FILE_PREFIX & "*.tmp"
    DoEvents
    
WipeFile_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    
    If mblnProcessUSB Then
        RaiseEvent CountTotals(strMsg, mcurFileCount, 0, curFileSize, False, False)
    Else
        RaiseEvent CountTotals(strFileName, mcurFileCount, curFileSize, curFileSize, False, False)
    End If
    
    DoEvents
    Exit Sub

WipeFile_Error:
    strRecord = vbTab & vbTab & "File:  " & strFileName & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & _
                "ERROR:  " & CStr(Err.Number) & "  " & Err.Description & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & _
                "Location:  " & MODULE_NAME & "." & ROUTINE_NAME

    strMsg = "File:  " & strFileName & vbNewLine & _
             "Error:  " & Err.Number & vbNewLine & Err.Description & _
             vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    Err.Clear
    ErrorMsg MODULE_NAME, ROUTINE_NAME, strMsg

    ' Update log file, if requested
    If mblnLogData Then
        RaiseEvent UpdateLogData(strRecord)
    End If

    Resume WipeFile_CleanUp

End Sub

' **************************************************************************
' Routine:       WipeFolder
'
' Description:   Renames folders 26 times and changes the folder date
'                properites to a random timestamp, then deletes the new
'                folder.
'
' Parameters:    strOldFolder - Folder to be removed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 03-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to update folder date properties with a
'              randomized timestamp
' 12-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic flow when removing folders
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added test for possible unicode folder or file names
' ***************************************************************************
Private Sub WipeFolder(ByVal strOldFolder As String)

    Dim strRecord    As String
    Dim strNewFolder As String
    
    On Error Resume Next

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Do not process USB Drive
    ' folders in this routine.
    ' They are handled in the
    ' ProcessUSB() routine.
    If mblnProcessUSB Then
        Exit Sub
    End If
                
    ' Cannot delete a root drive
    ' (Ex:  "D:\")
    If IsPathARoot(strOldFolder) Then
        Exit Sub
    End If
    
    ' Is this is a top level folder
    If StrComp(strOldFolder, mstrTopLevelFolder, vbTextCompare) = 0 Then

        If mblnRemoveTopFolder Then
            
            ' Remove top level folder
            strOldFolder = UnQualifyPath(strOldFolder)             ' remove trailing backslash
            SetFileAttributes strOldFolder, FILE_ATTRIBUTE_NORMAL  ' Reset folder attributes
            strNewFolder = CreateNewName(strOldFolder)                ' rename folder 26 times
        
        Else
            ' Do not remove top level folder
            GoTo WipeFolder_CleanUp
        End If

    Else
        
        ' Process subfolders
        strOldFolder = UnQualifyPath(strOldFolder)             ' remove trailing backslash
        SetFileAttributes strOldFolder, FILE_ATTRIBUTE_NORMAL  ' Reset folder attributes
        strNewFolder = CreateNewName(strOldFolder)                ' rename folder 26 times
        
    End If

    ' If old folder exists, delete it
    If IsPathValid(strOldFolder) Then
        
        ' reset old folder date properties
        ' with random date and time stamp
        SetPathFileDate strOldFolder
    
        FolderCleanUp strOldFolder, True
    
    End If
    
    ' If new folder exists, delete it
    If IsPathValid(strNewFolder) Then
        
        ' reset new folder date properties
        ' with random date and time stamp
        SetPathFileDate strNewFolder
    
        FolderCleanUp strNewFolder, True
    
    End If
    
WipeFolder_CleanUp:
    Err.Clear   ' Clear any error return codes
    DoEvents
          
    ' See if this folder still exist
    If IsPathValid(strOldFolder) Then
        mcurFolderCount = mcurFolderCount - 1    ' decrement folder count
        strRecord = "ERROR:" & vbTab & strOldFolder & " not removed"
    Else
        strRecord = vbTab & vbTab & "Folder:  " & strOldFolder & "  [DELETED]"
    End If

    If mblnLogData Then
        RaiseEvent UpdateLogData(strRecord)  ' Update log file
    End If
    
    On Error GoTo 0   ' Nullify this error trap

End Sub

' **************************************************************************
' Routine:       ProcessUSB
'
' Description:   This routine is designed for the user that feels that the
'                old data on their USB drive must be completely obliterated.
'
'                Designed to completely erase and overwrite a USB drive.
'                First a list of all folders at the root level are
'                collected, sorted and deleted.  Then any files left in the
'                root will be deleted.  Next, a series of temporary files
'                will be created and overwritten with a pattern selected by
'                the user.  When finished, all temp files will be deleted.
'
'                Busting the Biggest PC Myths
'                http://pcworld.about.com/magazine/2208p107id116572.htm
'
'                "Want to erase data from a hard drive you plan to toss?
'                Don't bother with a magnet. Overwrite the data that is
'                stored on the media instead. For a flash drive, fill up
'                the drive with anything, like pictures of your beloved
'                dachshund. Unlike with magnetic media, from which experts
'                can usually recover at least some overwritten data, once
'                new data is written to flash media, the old data is gone
'                forever."
'
' Note:          If you want to clean a flash drive then use Windows
'                Explorer to delete all the files and folders first.
'                Now select the DiscDataWipe option to wipe the free
'                space on the flash drive.
'
' Parameters:    strDrive - USB drive to be completely cleaned
'
' Returns:       TRUE - Successful completion
'                FALSE - something went wrong
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 11-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function ProcessUSB(ByVal strDrive As String) As Boolean

    Dim lngIndex    As Long      ' loop counter
    Dim strPath     As String    ' temp Path
    Dim strMsg      As String
    Dim strRecord   As String    ' log record
    Dim astrFiles() As String    ' array of temp file names
    Dim blnError    As Boolean
    Dim objDI       As cDiskInfo
    
    Const ROUTINE_NAME As String = "ProcessUSB"
    
    On Error GoTo ProcessUSB_Error

    strPath = vbNullString
    Erase astrFiles()                  ' Start with an empty array
    ProcessUSB = False
    mblnProcessUSB = False
    blnError = False
    strDrive = QualifyPath(strDrive)   ' Add a backslash to drive
    
    strMsg = "Removing folders and files on USB drive ( " & strDrive & " )"
    RaiseEvent WaitMsg(strMsg, &HE0E0E0)
    
    ' Update log file, if requested
    If mblnLogData Then
        strRecord = "STARTED:" & vbTab & "Wiping all data from USB drive ( " & strDrive & " )"
        RaiseEvent UpdateLogData(strRecord)
        RaiseEvent UpdateLogData(vbTab & strMsg)
    End If
    
    mcurOA_MaxAmt = mcurByteCount
    mcurOA_MaxPerFile = mcurByteCount
    mcurFileCount = 0@
    mcurByteCount = 0@
    mcurFolderCount = 0@
    mcurOA_CurrentAmt = 0@
    mcurCurrPerPass = 0@
    mlngStartTime = GetTickCount
    mlngPause = GetTickCount
    
    DoEvents
    RaiseEvent OverallProgress(0, mcurOA_MaxPerFile)  ' maintain overall totals
    CalculateProgress 0, mcurOA_MaxAmt, True          ' overall progress

    If mcolFolders.Count > 1 Then
                
        ' Sort the collection of folders listed at the root level
        ' of the drive. Since a collection is limited to 32,767
        ' entries, I do not believe we will have any problems.
        SortFolderList

        On Error Resume Next

        ' Loop thru and remove all files and folders
        For lngIndex = mcolFolders.Count To 1 Step -1

            ' Ignore the root itself but get the other folders
            DoEvents
            If Len(mcolFolders.Item(lngIndex)) > 3 Then

                mblnProcessUSB = True
                strPath = UnQualifyPath(mcolFolders.Item(lngIndex))

                ' Reset folder attributes
                SetFileAttributes strPath, FILE_ATTRIBUTE_NORMAL

                ' Delete the folder's contents
                FolderCleanUp strPath, True
                DoEvents

                ' See if this folder still exist
                If IsPathValid(strPath) Then
                    DoEvents
                    FolderCleanUp strPath, True
                    DoEvents
                End If

                ' See if this folder still exist
                If IsPathValid(strPath) Then

                    strRecord = "ERROR:" & vbTab & strPath & " not removed"
                    blnError = True

                    If mblnLogData Then
                        RaiseEvent UpdateLogData(strRecord)  ' Update log file
                    End If

                End If

            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If

        Next lngIndex
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo ProcessUSB_CleanUp
    End If
            
    On Error GoTo 0   ' Nullify this error trap
    On Error GoTo ProcessUSB_Error

    strPath = vbNullString
    mblnProcessUSB = True
    FilesAndFolders_Delete strDrive  ' Delete files in root drive
    ResetVariables                   ' Reset module accumulators
    EmptyCollection mcolFolders      ' empty folder collection
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo ProcessUSB_CleanUp
    End If
            
    ' Reset display
    RaiseEvent OverallProgress(0, 0)   ' maintain overall totals
    CalculateProgress 0, 0, True       ' overall progress
    
    If blnError Then
    
        strRecord = vbNullString
        strRecord = strRecord & "Some of the data on USB drive " & strDrive & " could not be removed."
        strRecord = strRecord & vbNewLine & vbNewLine
        strRecord = strRecord & "Stop this application and manually remove all data from USB drive " & _
        strRecord = strRecord & strDrive & " ." & vbNewLine
        strRecord = strRecord & "When finished, restart this application and try again.  Thank you."
            
        InfoMsg strRecord & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
    
    End If
        
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo ProcessUSB_CleanUp
    End If

    WipeTheFreeSpace strDrive
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo ProcessUSB_CleanUp
    End If

    ProcessUSB = True  ' Successful finish

ProcessUSB_CleanUp:
    DoEvents
    RaiseEvent WaitMsg(CLEANUP_MSG, vbYellow)
    strPath = strDrive & TEMP_FOLDER
    
    ' If the temp folder still exists then delete it
    If IsPathValid(strPath) Then
        SetFileAttributes strPath, FILE_ATTRIBUTE_NORMAL   ' Reset folder attributes
        FolderCleanUp strPath, True                        ' Remove the temp folder
    End If
    
    mblnProcessUSB = False  ' reset flag
    Erase astrFiles()       ' empty arrays when not needed
    Set objDI = Nothing     ' free objects from memory
    ResetVariables          ' Reset module accumulators
    
    ' Update log file, if requested
    If mblnLogData Then
        
        ' Check for a successful finish
        If ProcessUSB Then
            strRecord = "FINISHED:" & vbTab & "Wipe USB drive " & strDrive
            RaiseEvent UpdateLogData(strRecord)
        Else
            strRecord = "ABENDED:" & vbTab & "Wipe USB drive " & strDrive & vbNewLine & _
                        Space$(32) & ABEND_MSG
            RaiseEvent UpdateLogData(strRecord)
        End If
    
    End If

    RaiseEvent WaitMsg("", &HE0E0E0)  ' Reset to gray background
    On Error GoTo 0                   ' Nullify this error trap
    Exit Function

ProcessUSB_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
        
    If mblnLogData Then
        strRecord = "ABENDED:" & vbTab & "Wipe USB drive " & strDrive & vbNewLine & _
                    vbTab & vbTab & vbTab & vbTab & _
                    CStr(Err.Number) & "  " & Err.Description
        RaiseEvent UpdateLogData(strRecord)
    End If
    
    ProcessUSB = False
    Resume ProcessUSB_CleanUp
                        
End Function

' **************************************************************************
' Routine:       FillWithTempFiles
'
' Description:   This routine will create temp files in the DD_Temp folder
'                that are filled with null values.
'
' Parameters:    strDrive - Target drive
'                astrFiles() - Array of temp file names
'
' Returns:       astrFiles() array will be filled with temp file names.
'                Used by alternate wiping method only.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 11-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 19-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Converted routine from Sub to Function.  Added logic to stop
'              processing if less than 16kb of space left.
' 01-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated the file creation using the basBigFile module.
' ***************************************************************************
Private Function FillWithTempFiles(ByVal strDrive As String, _
                                   ByRef astrFiles() As String) As Boolean

    Dim lngIdx         As Long
    Dim lngStart       As Long       ' Starting index when creating extra files
    Dim lngIndex       As Long       ' loop counter
    Dim lngFileCount   As Long       ' Number of temp files to create
    Dim lngTmpFileCnt  As Long
    Dim blnNTFS        As Boolean    ' Flag designating NTFS file system
    Dim curFileSize    As Currency   ' Temp file size
    Dim curFreeSpace   As Currency   ' amount of freespace
    Dim strMsg         As String
    Dim strRecord      As String
    Dim strFileSys     As String     ' Partition file system (FAT32, NTFS)
    Dim strFileName    As String     ' temp filename
    Dim strTempFolder  As String     ' temp Path
    Dim adblTemp()     As Double     ' 2D array to hold file sizes and qty
    Dim objDI          As cDiskInfo
    
    Const ROUTINE_NAME As String = "FillWithTempFiles"

    On Error GoTo FillWithTempFiles_CleanUp
    
    mblnDiskFull = False
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo FillWithTempFiles_CleanUp
    End If

    If mblnProcessUSB Then
        strMsg = "Preparing USB drive " & strDrive
    Else
        strMsg = "Wipe free space on drive " & strDrive
    End If
    
    RaiseEvent WaitMsg(strMsg, &HE0E0E0)
    
    Erase astrFiles()  ' Start with empty arrays
    Erase adblTemp()
    
    ResetCounters
    mcurFolderCount = 1@
    
    ' update the log file, if requested
    If mblnLogData Then
        strRecord = Space$(5) & "Creating files filled with binary 0's on drive " & strDrive
        RaiseEvent UpdateLogData(strRecord)
    End If
    
    ' Format path to temp folder
    strTempFolder = UnQualifyPath(strDrive & TEMP_FOLDER)
    
    ' See if the DD_Temp folder already exist
    If IsPathValid(strTempFolder) Then
        ' Reset folder attributes
        SetFileAttributes strTempFolder, FILE_ATTRIBUTE_NORMAL
        FolderCleanUp strTempFolder, True  ' Delete DD_Temp folder
    End If
    
    ' Verify DD_Temp folder does not exist
    If Not IsPathValid(strTempFolder) Then
        MkDir strTempFolder  ' Create DD_Temp folder
    End If
    
    SetFileAttributes strTempFolder, FILE_ATTRIBUTE_HIDDEN       ' Hide DD_Temp folder
    strTempFolder = QualifyPath(strTempFolder)                   ' Add a backslash
    
    Set objDI = New cDiskInfo                                    ' Instantsiate class object
    With objDI
        curFreeSpace = .GetDiskSpaceInfo(strDrive, eFreespace)   ' Get amount of freespace on drive
        .GetVolumeInfo strDrive, strFileSys                      ' Determine drive partition file system
        blnNTFS = IIf(StrComp("NTFS", strFileSys, vbTextCompare) = 0, True, False)
    End With
    Set objDI = Nothing                                          ' Free class object from memory
    
    ' If the amount of free space is less than
    ' five megabytes then display a message and
    ' exit procedure
    If curFreeSpace <= MB_5 Then
        
        InfoMsg "Not enough free space on drive " & strDrive & vbNewLine & _
                "Clean up this drive first and then try again." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo FillWithTempFiles_CleanUp
        
    End If
    
    ' Determine maximum size of temp files
    mcurOA_MaxPerFile = curFreeSpace
    
    ' Determine number and size of files to be created
    adblTemp() = CalcTempFiles(curFreeSpace, blnNTFS)
    
    ReDim astrFiles(1)            ' Size return array with names of temp files
    astrFiles(0) = strTempFolder  ' Add path as first entry in return array
    
    DoEvents
    RaiseEvent CurrentProgress(0, 100)   ' Do not show current progress
    CalculateProgress 0, 100, False
    RaiseEvent CountTotals(strMsg, lngFileCount, mcurOA_MaxPerFile, 0, True, True)
    
    lngStart = 1       ' Set starting counters
    lngFileCount = 0
    
    For lngIndex = 0 To UBound(adblTemp)
    
        lngTmpFileCnt = CLng(adblTemp(lngIndex, 0))  ' number of files to create
        
        If lngTmpFileCnt > 0 Then
            
            curFileSize = CCur(adblTemp(lngIndex, 1))    ' Size of new files
            lngFileCount = lngFileCount + lngTmpFileCnt  ' Increment overall file count
            ReDim Preserve astrFiles(lngFileCount)       ' Increment array size
            
            For lngIdx = lngStart To lngFileCount
            
                ' format temp file name (ex:  ~0000001.tmp)
                DoEvents
                astrFiles(lngIdx) = "~" & Format$(lngIdx, "0000000") & ".tmp"
                
                ' Format fully qualified path and file name
                strFileName = strTempFolder & astrFiles(lngIdx)
            
                ' Create a temp file filled with null values
                CreateNullFile strFileName, curFileSize
                
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit For    ' exit For..Next loop
                End If
                                  
                ' See if we have run out of disk space
                If mblnDiskFull Then
                    Exit For
                End If
                
                DoEvents
                mcurFileCount = CCur(lngIdx)                          ' Track number of temp files
                mcurByteCount = mcurByteCount + curFileSize           ' Calc the byte count
                mcurOA_CurrentAmt = mcurOA_CurrentAmt + curFileSize   ' calc overall progress for pbar
                
                DoEvents
                RaiseEvent CurrentProgress(0, 100)                                ' Do not show current progress
                RaiseEvent CountTotals(strMsg, mcurFileCount, 0, 0, False, True)
                
                DoEvents
                CalculateProgress 0, 100, False
                RaiseEvent OverallProgress(mcurOA_CurrentAmt, mcurOA_MaxPerFile)  ' maintain overall totals
                CalculateProgress mcurOA_CurrentAmt, mcurOA_MaxPerFile, True      ' overall progress
    
            Next lngIdx
            
            lngStart = lngIdx  ' Save new starting position
        
        End If
    
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
                  
        ' See if we have run out of disk space
        If mblnDiskFull Then
            Exit For
        End If
                
    Next lngIndex

    FillWithTempFiles = True    ' Successful completion
        
FillWithTempFiles_CleanUp:
    Set objDI = Nothing   ' Verify object free from memory
    Erase adblTemp()      ' Always empty arrays when not needed
    mblnDiskFull = False
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        FillWithTempFiles = False
        Exit Function
    End If
    
    ' update the log file, if requested
    If mblnLogData Then
        strRecord = Space$(5) & "Finished creating " & Format$(mcurFileCount, "#,##0") & _
                    " temp files on drive " & strDrive
        RaiseEvent UpdateLogData(strRecord)
    End If
    
    On Error GoTo 0   ' Nullify this error trap
    
End Function

Private Sub CreateNullFile(ByVal strFileName As String, _
                           ByVal curFileSize As Currency)

    ' Called by FillWithTempFiles()
        
    Dim hFile As Long
            
    On Error GoTo CreateNullFile_Error
    
    ' Create a temp file filled with null values
    DoEvents
    If curFileSize < GB_1 Then
                             
        ' If file already exist then
        ' verify it is empty
        If IsPathValid(strFileName) Then
            hFile = FreeFile                        ' Get first free file handle
            Open strFileName For Output As #hFile   ' Create empty file
            Close #hFile                            ' Close file
        End If
            
        DoEvents
        hFile = FreeFile                                     ' Get first free file handle
        Open strFileName For Binary Access Write As #hFile   ' Open for writing
        Put #hFile, curFileSize, Chr$(0)                     ' Fill with binary 0's
        Close #hFile                                         ' Close file
    
    Else
        mobjBigFiles.CreateBigFile strFileName, curFileSize, False
    End If
    
CreateNullFile_CleanUp:
    DoEvents
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

CreateNullFile_Error:
    If Err.Number = 61 Then
        mblnDiskFull = True
    End If
    
    Err.Clear      ' reset error code
    Close #hFile   ' Close file
    GoTo CreateNullFile_CleanUp
    
End Sub

' **************************************************************************
' Routine:       MiscPatterns
'
' Description:   Loops thru an open file and overwrites each byte of data
'                with either all zeroes (Null values (ASCII decimal 0)) or
'                random generated data.
'
' Important:     This is good security for up thru "SECRET" documents only if
'                the overwrite is multiple (five or more) passes of
'                eWipe_RandomData or one pass using eWipe_BruceSchneier.
'
'                Internationally-renowned security technologist and author
'                Bruce Schneier recommends wiping a drive seven times. The
'                first pass overwrites the drive with Null values (ASCII
'                decimal 0), the second with Binary 1's (ASCII decimal 255),
'                and the next five with a random generated bit pattern.
'
'                Bruce Schneier   Write 1 - Null values (ASCII decimal 0)
'                                 Write 2 - Binary 1's (ASCII decimal 255)
'                                 Write 3 - Random data stream (Perform 5 times)
'                                 (7 overwrites equal one pass)
'
'                These processes are performed from 1 to 99 times as defined
'                by the user.
'
'                If verification is requested then the final pass will
'                be compared with the data that was supposed to be written.
'
'                If opting to wipe the free space using the alternate method,
'                the pattern is similar to US DoD 5220.22-M as described in
'                National Industrial Security Program Operating Manual (NISPOM)
'                DoD 5220.22-M, dtd February 28, 2006.  Chapter 8 -Information
'                System Security, Section 3-Common Requirements, 8-301 Clearing
'                and Sanitization:
'
'                "Overwrite all addressable locations with a character, its
'                complement, then a random character and verify."
'                http://www.dss.mil/isp/fac_clear/download_nispom.html
'
'                US DoD 5220.22-M (3 overwrites) without verify
'                Alternate Method  Write 1 - One random character
'                                  Write 2 - Complement of previous character
'                                  Write 3 - One random character
'                                  (3 overwrites equal one pass)
'
' Disposal:      If the hard disk has had data classified greater than
'                "CONFIDENTIAL", then the disk should be replaced with a new
'                one.  Since the cost of a hard disk has dropped so
'                dramatically, this should not be a factor.  You should be
'                considering the question, "What is my information worth to
'                someone else?".
'
'                Steps to follow to dispose of the old hard drive:
'                1. Overwrite multiple times with random data (Min 5 times).
'                   I recommend the Dban web site and creating a bootable CD
'                   or USB device that will wipe every sector on a disk. This
'                   is freeware and several governments approve its use.
'                   http://www.dban.org/
'                2. Remove disk from the old desktop or laptop and record the
'                   manufacturer, model, serial number, date of destruction and
'                   name of individual performing this process.
'                3. Plate area should be drilled in several places using a
'                   1/2 inch drill bit.
'                4. Disintegrate, incinerate, pulverize, shred, or melt the
'                   hard drive.
'
'                All of the above should be witnessed by at least two
'                additional persons and documented.
'
'                Ref:  Guidelines for Media Sanitization (SP800-88 dtd Sep-2006)
'                      http://csrc.nist.gov/publications/nistpubs/800-88/NISTSP800-88_rev1.pdf
'
' Parameters:    strFilename - Name of the file to be processed
'
' Returns:       TRUE - Successful completion
'                FALSE - something went wrong
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-FEB-2007  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 14-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              The write process has been speeded up by 50% or more by
'              adjusting the record length based on amount of data left to
'              write.  This process calls the routine GetBlockSize().
' 24-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added option to write zeroes as the last write
' 14-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Simplified logic for when to create random data
' 12-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added wiping free space using an alternate method.
'                 1. Select a random byte and create a record
'                 2. Get complement of previous byte and create a record
'                 3. Select a random byte and create a record
' ***************************************************************************
Private Function MiscPatterns(ByVal strFileName As String) As Boolean

    Dim hFile         As Long      ' file handle
    Dim lngIdx        As Long
    Dim lngIndex      As Long      ' loop counter
    Dim lngPause      As Long
    Dim lngWriteCnt   As Long      ' Number of overwrites to a file
    Dim lngMaxDataAmt As Long      ' Max amount of random data to create
    Dim lngNoOfWrites As Long      ' number of writes to a file to equal one overwrite
    Dim lngPointer    As Long      ' Rnd data string pointer
    Dim lngFilePos    As Long      ' Starting position in the file
    Dim lngRecordSize As Long      ' length of data record
    Dim blnLastPass   As Boolean   ' Flag designating last pass being performed
    Dim curFileSize   As Currency  ' original file size
    Dim curAmtLeft    As Currency  ' used to determine how much is left to write
    Dim curFilePos    As Currency  ' Starting position in the file
    Dim strData       As String    ' data string to hold output data
    Dim strRecord     As String    ' log file message
    Dim strPattern    As String    ' log info designating the pattern in use
    Dim astrData(8)   As String
    Dim bytTemp1      As Byte
    Dim bytTemp2      As Byte
    Dim bytTemp3      As Byte
    Dim abytRnd()     As Byte      ' random data
    Dim abytTemp()    As Byte
    Dim abytVerify()  As Byte      ' used to verify written data
    
    Const ROUTINE_NAME As String = "MiscPatterns"

    On Error GoTo MiscPatterns_Error

    Erase abytRnd()    ' Always start with empty arrays
    Erase abytTemp()
    Erase abytVerify()
    Erase astrData()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    MiscPatterns = False
    blnLastPass = False                                   ' Init last pass flag
    lngWriteCnt = MAX_INT                                 ' Init number of overwrites
    SetFileAttributes strFileName, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes
    mobjBigFiles.CalcFileSize strFileName, curFileSize    ' Get the size of the file
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo MiscPatterns_CleanUp
    End If
    
    With mobjRandom
        
        Rnd -1               ' Reset VB random number generator
        Randomize .RndSeed   ' Reseed VB random number generator
        
        Select Case mlngWipeMethod
               Case eWipe_BruceSchneier
                    ' Determine if the last write is to be all zeroes
                    lngNoOfWrites = IIf(mblnZeroLastWrite, 7, 6)
        End Select
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo MiscPatterns_CleanUp
        End If
    
        For lngIndex = 1 To mlngPasses
    
            If lngIndex = mlngPasses Then
                blnLastPass = True
            End If
            
            lngFilePos = 1                               ' Set pointer to beginning of file
            curFilePos = 0                               ' Set pointer to beginning of file
            curAmtLeft = curFileSize                     ' Work with copy of file size
            lngRecordSize = GetBlockSize(curAmtLeft)     ' determine record size of output record
            lngMaxDataAmt = KB_4 + lngRecordSize         ' Max amount of random data to create
                    
            RaiseEvent CountPasses(lngIndex, mlngPasses)
            RaiseEvent CurrentProgress(0, curFileSize)   ' maintain ongoing totals
            CalculateProgress 0, curFileSize, False      ' individual progress
            
            ' Open target file
            If Not mobjBigFiles.OpenReadWrite(strFileName, hFile) Then
                gblnStopProcessing = True
            End If
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo MiscPatterns_CleanUp
            End If
    
            Do
                ' An error occurred or user opted to STOP processing
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
        
                If lngRecordSize > 1 Then
                    
                    Select Case mlngWipeMethod
                                        
                           Case eWipe_AlternateMethod    ' Alternate method of wiping freespace
                                
                                ' After 1000 overwrites, generate new data
                                If lngWriteCnt > 1000 Then
                                 
                                    lngWriteCnt = 0                      ' reset overwrite counter
                                    bytTemp1 = Int(Rnd() * 256)          ' Create one byte of random data (0-255) (Ex:  15 = 0x0F = 00001111)
                                    bytTemp2 = GetComplement(bytTemp1)   ' Get complement of previous byte        (Ex: 240 = 0xF0 = 11110000)
                                                                         
                                    ' All three bytes should be unique
                                    Do
                                       bytTemp3 = Int(Rnd() * 256)       ' Create one byte of random data (0-255) (Ex: 153 = 0x99 = 10011001)
                                    Loop While bytTemp3 = bytTemp1 Or _
                                               bytTemp3 = bytTemp2
                                End If
                                
                                Erase astrData()                                ' Always start with an empty array
                                astrData(0) = String$(lngRecordSize, bytTemp1)  ' Create string using random byte
                                astrData(1) = String$(lngRecordSize, bytTemp2)  ' Create string using compement of previous byte
                                astrData(2) = String$(lngRecordSize, bytTemp3)  ' Create string using random byte
                                
                                ' overwrite this position three times
                                For lngIdx = 0 To 2
                                
                                    Erase abytTemp()
                                    abytTemp() = StringToByteArray(astrData(lngIdx))
                                    
                                    ' Excessive number of DoEvents is to allow the read and
                                    ' write process to stay in sync with the application speed.
                                    '
                                    ' Write data to file
                                    If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                                        gblnStopProcessing = True
                                    End If
                   
                                Next lngIdx
    
                           Case eWipe_BinaryZeroes   ' binary 0's (Hex 0x00)
                                Erase abytTemp()                          ' Always start with an empty array
                                strData = String$(lngRecordSize, 0)       ' Stream of binary 0's (ASCII decimal 0)
                                abytTemp() = StringToByteArray(strData)   ' Convert string to byte array
                                
                                ' Excessive number of DoEvents is to allow the read and
                                ' write process to stay in sync with the application speed.
                                '
                                ' Write data to file
                                If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                                    gblnStopProcessing = True
                                    Exit Do   ' exit Do..Loop
                                End If
                   
                                ' If verification has been requested then read the record
                                ' into a buffer and verify it's contents against what was
                                ' supposed to be written.
                                If mblnVerifyData Then
                                    If blnLastPass Then
                                        
                                        strPattern = "0x00"
                                        ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                                        
                                        ' Capture the data that was just written to the file
                                        DoEvents
                                        If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                            gblnStopProcessing = True
                                            Exit Do   ' exit Do..Loop
                                        End If
                                    
                                        If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                               strFileName, curFilePos, strPattern) Then
                                                               
                                            gblnStopProcessing = True
                                            Exit Do   ' exit Do..Loop
                                        End If
                                    
                                    End If
                                End If
                                
                           Case eWipe_RandomData   ' random data
                                ' Every 100 overwrites to
                                ' file generate new data
                                If lngWriteCnt > 100 Then
                                    lngWriteCnt = 0                                     ' reset overwrite counter
                                    Erase abytRnd()                                     ' Always start with an empty array
                                    abytRnd() = .BuildRndData(lngMaxDataAmt, , False)   ' Fill a byte array with random data
                                End If
                                
                                lngPointer = Int(Rnd() * KB_4)               ' Create a starting point between 0-4095
                                ReDim abytTemp(lngRecordSize)                ' Size temp array
                                CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                                ReDim Preserve abytTemp(lngRecordSize - 1)   ' Re-Size temp array
                                DoEvents
                    
                                ' Excessive number of DoEvents is to allow the read and
                                ' write process to stay in sync with the application speed.
                                '
                                ' Write data to file
                                If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                                    gblnStopProcessing = True
                                    Exit Do   ' exit Do..Loop
                                End If
                   
                                ' If verification has been requested then read the record
                                ' into a buffer and verify it's contents against what was
                                ' supposed to be written.
                                If mblnVerifyData Then
                                    If blnLastPass Then
                                        
                                        strPattern = "Random"                ' Type of data used for error reporting
                                        ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                                        
                                        ' Capture the data that was just written to the file
                                        DoEvents
                                        If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                            gblnStopProcessing = True
                                            Exit Do   ' exit Do..Loop
                                        End If
                                    
                                        If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                               strFileName, curFilePos, strPattern) Then
                                                               
                                            gblnStopProcessing = True
                                            Exit Do   ' exit Do..Loop
                                        End If
                                    
                                    End If
                                End If
                                
                           Case eWipe_BruceSchneier   ' Bruce Schneier's Algorithm (7 overwrites)
                                ' Every 50 overwrites to
                                ' file generate new data
                                If lngWriteCnt >= 50 Then
                                    lngWriteCnt = 0                                     ' reset overwrite counter
                                    Erase abytRnd()                                     ' Always start with an empty array
                                    abytRnd() = .BuildRndData(lngMaxDataAmt, , False)   ' Fill a byte array with random data
                                End If
                                
                                astrData(0) = String$(lngRecordSize, 0)    ' Stream of binary 0's (ASCII decimal 0)
                                astrData(1) = String$(lngRecordSize, 255)  ' Stream of binary 1's (ASCII decimal 255)
                                
                                ' Fill with random generated data
                                For lngIdx = 2 To 6
                                    
                                    lngPointer = Int(Rnd() * KB_4)                     ' Create a starting point between 0-4095
                                    ReDim abytTemp(lngRecordSize)                      ' Size temp array
                                    CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                                    ReDim Preserve abytTemp(lngRecordSize - 1)         ' Re-Size temp array
                                    astrData(lngIdx) = ByteArrayToString(abytTemp())   ' Convert temp array to string data
                                    
                                Next lngIdx
                                
                                ' Optional last write using zeroes
                                If mblnZeroLastWrite Then
                                    astrData(7) = String$(lngRecordSize, 0)   ' Stream of binary 0's (ASCII decimal 0)
                                End If
                                
                                ' overwrite this position multiple times
                                For lngIdx = 0 To lngNoOfWrites
    
                                    Erase abytTemp()
                                    abytTemp() = StringToByteArray(astrData(lngIdx))
                                                            
                                    ' Excessive number of DoEvents is to allow the read and
                                    ' write process to stay in sync with the application speed.
                                    '
                                    ' Write data to file
                                    If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                                        gblnStopProcessing = True
                                    End If
                   
                                    ' If verification has been requested then read the record
                                    ' into a buffer and verify it's contents against what was
                                    ' supposed to be written.
                                    If mblnVerifyData Then
                                        If blnLastPass Then
                                        
                                            ' Type of data used for error reporting
                                            Select Case lngIdx
                                                   Case 0: strPattern = "0x00"
                                                   Case 1: strPattern = "0xFF"
                                                   Case 2 To 6: strPattern = "Random"
                                                   Case Else: strPattern = "0x00"
                                            End Select
                                        
                                            ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                                            
                                            ' Capture the data that was just written to the file
                                            DoEvents
                                            If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                                gblnStopProcessing = True
                                                Exit For    ' exit For..Next loop
                                            End If
                                            
                                            If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                                   strFileName, curFilePos, strPattern) Then
                                                                   
                                                gblnStopProcessing = True
                                                Exit For    ' exit For..Next loop
                                            End If
                                    
                                        End If
                                    End If
    
                                    ' An error occurred or user opted to STOP processing
                                    DoEvents
                                    If gblnStopProcessing Then
                                        Exit For    ' exit For..Next loop
                                    End If
    
                                Next lngIdx
                    End Select
                
                End If
    
                Erase abytTemp()   ' Empty temp array
                
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
                    
                curFilePos = curFilePos + CCur(lngRecordSize)          ' Adjust pointers accordingly
                lngWriteCnt = lngWriteCnt + 1                          ' update overwrite counter
                curAmtLeft = curAmtLeft - lngRecordSize                ' calc how much is left
                mcurCurrPerPass = mcurCurrPerPass + lngRecordSize      ' Current amt written this pass
                mcurOA_CurrentAmt = mcurOA_CurrentAmt + lngRecordSize  ' calc overall progress for pbar
                lngRecordSize = GetBlockSize(curAmtLeft)               ' determine record size of output record
    
                RaiseEvent CurrentProgress(curFilePos, curFileSize)    ' individual ongoing totals
                CalculateProgress curFilePos, curFileSize, False       ' individual progress
    
                RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxPerFile)  ' maintain overall totals
                CalculateProgress mcurOA_CurrentAmt, mcurOA_MaxAmt, True        ' overall progress
                DoEvents
    
                If curAmtLeft < 1 Then
                    Exit Do   ' exit Do..Loop
                End If
                    
            Loop
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            ' Verify file handle is released
            mobjBigFiles.API_CloseFile hFile
    
            ' The pause is here to allow
            ' the write process to sync up
            ' with the speed of the CPU
            DoEvents
            lngPause = GetTickCount + 200
            Do While lngPause > GetTickCount
                DoEvents
            Loop
                    
            Erase abytTemp()   ' Empty the temp arrays
            Erase astrData()
            
            If lngIndex < mlngPasses Then
                mcurCurrPerPass = mcurCurrPerPass - curFileSize
            End If
    
        Next lngIndex
    End With
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo MiscPatterns_CleanUp
    End If
    
    MiscPatterns = True
    
MiscPatterns_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        If mlngWipeMethod = eWipe_AlternateMethod Then
            RaiseEvent WaitMsg(CLEANUP_MSG, vbYellow)   ' Display yellow message
        End If
    End If
    
    DoEvents
    mobjBigFiles.API_CloseFile hFile                       ' Verify file handle is released
    RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
    CalculateProgress 100, 100, False                      ' individual progress
    DoEvents
    
    ' Update log file, if requested
    If gblnStopProcessing Then
        If mblnLogData Then
            strRecord = "ABENDED:" & vbTab & "Misc file wipe on " & strFileName & vbNewLine & _
                        Space$(32) & ABEND_MSG
            RaiseEvent UpdateLogData(strRecord)
        End If
    End If
    
    Erase astrData()   ' Always empty arrays when not needed
    Erase abytRnd()
    Erase abytTemp()
    Erase abytVerify()
    
    strData = vbNullString    ' Empty data strings
    strRecord = vbNullString
    
    RaiseEvent WaitMsg("", &HE0E0E0)  ' Reset to gray background
    On Error GoTo 0                   ' Nullify this error trap
    Exit Function

MiscPatterns_Error:
    ' Update log file, if requested
    If mblnLogData Then
        strRecord = vbTab & vbTab & "File:  " & strFileName & vbNewLine & _
                    vbTab & vbTab & vbTab & vbTab & _
                    CStr(Err.Number) & "  " & Err.Description
        RaiseEvent UpdateLogData(strRecord)
    End If
    
    strRecord = "ERROR File:  " & strFileName & vbNewLine & _
                CStr(Err.Number) & "  " & Err.Description
    ErrorMsg MODULE_NAME, ROUTINE_NAME, strRecord

    MiscPatterns = False
    Resume MiscPatterns_CleanUp

End Function

' **************************************************************************
' Routine:       DoDPatterns
'
' Description:   Loops thru an open file and overwrites each sector with a
'                pre-defined pattern.
'
' Reference:     National Industrial Security Program Operating Manual
'                (NISPOM) DoD 5220.22-M, dtd February 28, 2006.
'                Chapter 8 -Information System Security, Section 3-Common
'                Requirements, 8-301 Clearing and Sanitization
'                http://www.dss.mil/isp/fac_clear/download_nispom.html
'
'                "Overwrite all addressable locations with a character, its
'                complement, then a random character and verify.  THIS
'                METHOD IS NOT APPROVED FOR SANITIZING MEDIA THAT CONTAINS
'                TOP SECRET INFORMATION."
'
'                Dban Forum
'                http://sourceforge.net/projects/dban/forums/forum/208932
'
'                -----------------------------------------------------------
'                U.S. Standard, DoD 5220.22-M (E)  [US DoD Short]
'                Three overwrites
'
'                A standard was developed by the Defence Security Service
'                (DSS) which should solve the problem of the permanent
'                removal of data for some time. This was used by many
'                commercial enterprises. Under the National Industrial
'                Security Program (NISP) representatives of the Industrial
'                Security presented their security programs. As a part of
'                these NISP the DSS developed the DoD 5220.22-M standard
'                (National Industrial Security Program Operating Manual -
'                NISPOM). Which is used in the meantime in almost every
'                deletion tool.
'
'                In this manual is beside other procedures the description
'                of a method for the removal of data from magnetic media.
'                The NISP defines a technique for the overwriting of data
'                so that all information is removed from the medium.
'
'                US DoD Short    Write 1 - One random character
'                                Write 2 - Complement of previous character
'                                Write 3 - Random data stream
'                                (3 overwrites equal one pass)
'
'                In the regulations of the US American Ministry of Defense
'                people refer expressly to the deletion of information on
'                media with the military classification "Secret" or
'                "TOP-Secret" which is not allowed with this method.
'
'                -----------------------------------------------------------
'                U.S. Standard, DoD 5220.22-M (ECE)  [US DoD Long]
'                Seven overwrites
'
'                This method is an extended variant of the DoD 5220.22-M.
'                This variant of the DoD Standard uses for overwriting the
'                data seven runs. Here the data is overwritten two times by
'                using the DoD 5220.22-M (E) standard and one time with
'                random value DoD 5220.22-M (C).
'
'                Steps 1-3 overwrite the data with the DoD 5220.22-M (E)
'                          Standard (US DoD Short)
'                Step 4    overwrite the data with one pseudo random value,
'                          DoD 5220.22-M (C)
'                Steps 5-7 overwrite the data with the DoD 5220.22-M (E)
'                          Standard (US DoD Short) again
'
'                US DoD Long     Write 1 - One random character
'                                Write 2 - Complement of previous character
'                                Write 3 - Random data stream
'                                Write 4 - One random character
'                                Write 5 - One random character
'                                Write 6 - Complement of previous character
'                                Write 7 - Random data stream
'                                (7 overwrites equal one pass)
'
'                Verification will be checked to give you a true security
'                overwrite.  The final pass will be compared with the data
'                that was supposed to be written.
'
' Important:     This is good security for up thru "SECRET" documents only.
'
' Disposal:      If the hard disk has had data classified greater than
'                "CONFIDENTIAL", then the disk should be replaced with a new
'                one.  Since the cost of a hard disk has dropped so
'                dramatically, this should not be a factor.  You should be
'                considering the question, "What is my information worth to
'                someone else?".
'
'                Steps to follow to dispose of the old hard drive:
'                1. Overwrite multiple times with random data (Min 5 times).
'                   I recommend the Dban web site and creating a bootable CD
'                   or USB device that will wipe every sector on a disk. This
'                   is freeware and several governments approve its use.
'                   http://www.dban.org/
'                2. Remove disk from the old desktop or laptop and record the
'                   manufacturer, model, serial number, date of destruction and
'                   name of individual performing this process.
'                3. Plate area should be drilled in several places using a
'                   1/2 inch drill bit.
'                4. Disintegrate, incinerate, pulverize, shred, or melt the
'                   hard drive.
'
'                All of the above should be witnessed by at least two
'                additional persons and documented.
'
'                Ref:  Guidelines for Media Sanitization (SP800-88 dtd Sep-2006)
'                      http://csrc.nist.gov/publications/nistpubs/800-88/NISTSP800-88_rev1.pdf
'
' Parameters:    strFilename - Name of the file to be processed
'
' Returns:       TRUE - Successful completion
'                FALSE - something went wrong
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 25-FEB-2005  Kenneth Ives  kenaso@tx.rr.com
'              Updated how random data was to be determined for the output
'              record.  This speeded up the process considerably.
' 04-JAN-2006  Kenneth Ives  kenaso@tx.rr.com
'              Added US DoD Short & Long algorithms. I found these on the
'              Dban forum.
' 13-MAR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified the pointer for gathering the random data so as to
'              speed up this process.
' 14-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              The write process has been speeded up by 50% or more by
'              adjusting the record length based on amount of data left to
'              write.  This process calls the routine GetBlockSize().
' 24-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added option to write zeroes as the last write
' 14-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Simplified logic for when to create random data
' 15-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Force verification of last pass to meet DoD (Dept of Defense)
'              standards
' ***************************************************************************
Private Function DoDPatterns(ByVal strFileName As String) As Boolean

    Dim hFile         As Long      ' file handle
    Dim lngIdx        As Long      ' loop counter
    Dim lngIndex      As Long      ' loop counter
    Dim lngPause      As Long
    Dim lngPointer    As Long      ' Rnd data string pointer
    Dim lngWriteCnt   As Long      ' Number of overwrites to a file
    Dim lngRecordSize As Long      ' length of data record
    Dim lngMaxDataAmt As Long      ' Max amount of random data to create
    Dim lngNoOfWrites As Long      ' number of writes to a file to equal one overwrite
    Dim blnLastPass   As Boolean   ' Flag designating last pass being performed
    Dim curAmtLeft    As Currency  ' used to determine how much is left to write
    Dim curFilePos    As Currency  ' Starting position in the file
    Dim curFileSize   As Currency
    Dim strRecord     As String    ' log file message
    Dim strPattern    As String    ' log info designating the pattern in use
    Dim astrData(8)   As String    ' data array to hold overwrite data
    Dim bytTemp1      As Byte
    Dim bytTemp2      As Byte
    Dim bytTemp3      As Byte
    Dim bytTemp4      As Byte
    Dim bytTemp5      As Byte
    Dim abytRnd()     As Byte      ' random data
    Dim abytTemp()    As Byte      ' Temp data
    Dim abytVerify()  As Byte      ' used to verify written data

    Const ROUTINE_NAME As String = "DoDPatterns"

    On Error GoTo DoDPatterns_Error

    Erase abytRnd()  ' make sure we start with empty arrays
    Erase abytTemp()
    Erase abytVerify()
    Erase astrData()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Determine if the last write is to be all zeroes
    Select Case mlngWipeMethod
           Case eWipe_US_DoD_Short: lngNoOfWrites = IIf(mblnZeroLastWrite, 3, 2)
           Case eWipe_US_DoD_Long:  lngNoOfWrites = IIf(mblnZeroLastWrite, 7, 6)
    End Select
    
    lngWriteCnt = MAX_INT   ' Init file overwrite counter
    blnLastPass = False     ' Init last pass flag

    ' Reset file attributes
    SetFileAttributes strFileName, FILE_ATTRIBUTE_NORMAL
    
    mobjBigFiles.CalcFileSize strFileName, curFileSize   ' Get file size

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo DoDPatterns_CleanUp
    End If

    With mobjRandom
        
        Rnd -1               ' Reset VB random number generator
        Randomize .RndSeed   ' Reseed VB random number generator
        
        For lngIndex = 1 To mlngPasses
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            DoEvents
            RaiseEvent CountPasses(lngIndex, mlngPasses)
            RaiseEvent CurrentProgress(0, curFileSize)     ' maintain the ongoing totals
            CalculateProgress 0, curFileSize, False        ' individual progress
            DoEvents
    
            If lngIndex = mlngPasses Then
                blnLastPass = True
            End If
            
            curFilePos = 0                             ' Set pointer to beginning of file
            curAmtLeft = curFileSize                   ' Work with copy of file size
            lngRecordSize = GetBlockSize(curAmtLeft)   ' determine record size of output record
            lngMaxDataAmt = KB_4 + lngRecordSize       ' Max amount of random data to create
                    
            ' Open target file
            If Not mobjBigFiles.OpenReadWrite(strFileName, hFile) Then
                gblnStopProcessing = True
            End If
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            Do
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
        
                Erase abytTemp()   ' Always start with an empty arrays
                Erase astrData()
                                
                If lngRecordSize > 1 Then
    
                    ' After 50 overwrites, generate new data
                    If lngWriteCnt >= 50 Then
                        lngWriteCnt = 0                                     ' reset overwrite counter
                        Erase abytRnd()                                     ' Always start with an empty array
                        abytRnd() = .BuildRndData(lngMaxDataAmt, , False)   ' Fill a byte array with random data
                    End If
                    
                    Select Case mlngWipeMethod
                           
                           Case eWipe_US_DoD_Short   ' US DoD Short (3 overwrites)
                                lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                                bytTemp1 = abytRnd(lngPointer)                  ' Capture one byte of random data (Ex:  85 = 0x55 = 01010101)
                                bytTemp2 = GetComplement(bytTemp1)              ' Get complement of previous byte (Ex: 170 = 0xAA = 10101010)
                                astrData(0) = String$(lngRecordSize, bytTemp1)  ' Create string using random byte
                                astrData(1) = String$(lngRecordSize, bytTemp2)  ' Create string using complement byte
                                
                                ' Fill with random generated data
                                lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                                ReDim abytTemp(lngRecordSize)                   ' Size temp array
                                CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                                ReDim Preserve abytTemp(lngRecordSize - 1)      ' Re-Size temp array
                                astrData(2) = ByteArrayToString(abytTemp())     ' Convert temp array to string data
                                                                 
                                ' Optional last write using zeroes
                                If mblnZeroLastWrite Then
                                    astrData(3) = String$(lngRecordSize, 0)     ' Stream of binary 0's (ASCII decimal 0)
                                End If
                                                                
                                ' An error occurred or user opted to STOP processing
                                DoEvents
                                If gblnStopProcessing Then
                                    Exit Do      ' exit this DO loop
                                End If
                              
                                ' overwrite this position multiple times
                                For lngIdx = 0 To lngNoOfWrites
    
                                    Erase abytTemp()
                                    abytTemp() = StringToByteArray(astrData(lngIdx))
                                                            
                                    ' Excessive number of DoEvents is to allow the read and
                                    ' write process to stay in sync with the application speed.
                                    '
                                    ' Write data to file
                                    DoEvents
                                    If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                                        gblnStopProcessing = True
                                        Exit For    ' exit For..Next loop
                                    End If
                   
                                    ' Verification has been requested then read the record
                                    ' into a buffer and verify it's contents against what was
                                    ' supposed to be written.
                                    If blnLastPass Then
                                    
                                        ' Type of data used for error reporting
                                        Select Case lngIdx
                                               Case 0:    strPattern = "0x" & Right$("00" & Hex$(bytTemp1), 2)
                                               Case 1:    strPattern = "0x" & Right$("00" & Hex$(bytTemp2), 2)
                                               Case 2:    strPattern = "Random"
                                               Case Else: strPattern = "0x00"
                                        End Select
                                         
                                        ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                                        
                                        ' Capture the data that was just written to the file
                                        DoEvents
                                        If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                            gblnStopProcessing = True
                                            Exit For    ' exit For..Next loop
                                        End If
                                    
                                        If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                               strFileName, curFilePos, strPattern) Then
                                                               
                                            gblnStopProcessing = True
                                            Exit For    ' exit For..Next loop
                                        End If
                                    
                                    End If
    
                                    ' An error occurred or user opted to STOP processing
                                    DoEvents
                                    If gblnStopProcessing Then
                                        Exit For    ' exit For..Next loop
                                    End If
    
                                Next lngIdx
                               
                           Case eWipe_US_DoD_Long   ' US DoD Long (7 overwrites)
                                lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                                bytTemp1 = abytRnd(lngPointer)                  ' Capture one byte of random data (Ex:  85 = 0x55 = 01010101)
                                bytTemp2 = GetComplement(bytTemp1)              ' Get complement of previous byte (Ex: 170 = 0xAA = 10101010)
                                astrData(0) = String$(lngRecordSize, bytTemp1)  ' Create string using random byte
                                astrData(1) = String$(lngRecordSize, bytTemp2)  ' Create string using complement byte
                                
                                ' Fill with random generated data
                                lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                                ReDim abytTemp(lngRecordSize)                   ' Size temp array
                                CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                                ReDim Preserve abytTemp(lngRecordSize - 1)      ' Re-Size temp array
                                astrData(2) = ByteArrayToString(abytTemp())     ' Convert temp array to string data
                                
                                lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                                bytTemp3 = abytRnd(lngPointer)                  ' Capture one byte of random data (Ex:  15 = 0x0F = 00001111)
                                astrData(3) = String$(lngRecordSize, bytTemp3)  ' Create string using random byte
                                
                                ' get one char from random data byte array
                                lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                                bytTemp4 = abytRnd(lngPointer)                  ' Capture one byte of random data (Ex:  85 = 0x55 = 01010101)
                                bytTemp5 = GetComplement(bytTemp4)              ' Get complement of previous byte (Ex: 170 = 0xAA = 10101010)
                                astrData(4) = String$(lngRecordSize, bytTemp4)  ' Create string using random byte
                                astrData(5) = String$(lngRecordSize, bytTemp5)  ' Create string using complement byte
                                
                                ' Fill with random generated data
                                lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                                ReDim abytTemp(lngRecordSize)                   ' Size temp array
                                CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                                ReDim Preserve abytTemp(lngRecordSize - 1)      ' Re-Size temp array
                                astrData(6) = ByteArrayToString(abytTemp())     ' Convert temp array to string data
                                
                                ' Optional last write using zeroes
                                If mblnZeroLastWrite Then
                                    astrData(7) = String$(lngRecordSize, 0)     ' Stream of binary 0's (ASCII decimal 0)
                                End If
                                
                                ' An error occurred or user opted to STOP processing
                                DoEvents
                                If gblnStopProcessing Then
                                    Exit Do      ' exit this DO loop
                                End If
                              
                                ' overwrite this position multiple times
                                For lngIdx = 0 To lngNoOfWrites
    
                                    Erase abytTemp()                                  ' Verify array is empty
                                    abytTemp() = StringToByteArray(astrData(lngIdx))  ' Convert string dta to byte array
                                                            
                                    ' Excessive number of DoEvents is to allow the read and
                                    ' write process to stay in sync with the application speed.
                                    '
                                    ' Write data to file
                                    DoEvents
                                    If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                                        gblnStopProcessing = True
                                    End If
                   
                                    ' An error occurred or user opted to STOP processing
                                    DoEvents
                                    If gblnStopProcessing Then
                                        Exit For    ' exit For..Next loop
                                    End If
    
                                    ' Verification has been requested then read the record
                                    ' into a buffer and verify it's contents against what was
                                    ' supposed to be written.
                                    If blnLastPass Then
                                    
                                        ' Type of data used for error reporting
                                        Select Case lngIdx
                                               Case 0:    strPattern = "0x" & Right$("00" & Hex$(bytTemp1), 2)
                                               Case 1:    strPattern = "0x" & Right$("00" & Hex$(bytTemp2), 2)
                                               Case 2:    strPattern = "Random"
                                               Case 3:    strPattern = "0x" & Right$("00" & Hex$(bytTemp3), 2)
                                               Case 4:    strPattern = "0x" & Right$("00" & Hex$(bytTemp4), 2)
                                               Case 5:    strPattern = "0x" & Right$("00" & Hex$(bytTemp5), 2)
                                               Case 6:    strPattern = "Random"
                                               Case Else: strPattern = "0x00"
                                        End Select
                                         
                                        ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                                        
                                        ' Capture the data that was just written to the file
                                        DoEvents
                                        If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                            gblnStopProcessing = True
                                            Exit For    ' exit For..Next loop
                                        End If
                                    
                                        If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                               strFileName, curFilePos, strPattern) Then
                                                               
                                            gblnStopProcessing = True
                                            Exit For    ' exit For..Next loop
                                        End If
                                    
                                    End If
    
                                    ' An error occurred or user opted to STOP processing
                                    DoEvents
                                    If gblnStopProcessing Then
                                        Exit For    ' exit For..Next loop
                                    End If
    
                                Next lngIdx
                    End Select
                                
                End If
    
                Erase abytTemp()   ' Empty temp array
                
                 ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do      ' exit this DO loop
                End If
                              
                DoEvents
                curFilePos = curFilePos + CCur(lngRecordSize)          ' Adjust pointers accordingly
                lngWriteCnt = lngWriteCnt + 1                          ' update overwrite counter
                curAmtLeft = curAmtLeft - lngRecordSize                ' calc how much is left
                mcurCurrPerPass = mcurCurrPerPass + lngRecordSize      ' Current amt written this pass
                mcurOA_CurrentAmt = mcurOA_CurrentAmt + lngRecordSize  ' calc overall progress for pbar
                lngRecordSize = GetBlockSize(curAmtLeft)               ' determine record size of output record
    
                DoEvents
                RaiseEvent CurrentProgress(curFilePos, curFileSize)    ' individual ongoing totals
                CalculateProgress curFilePos, curFileSize, False       ' individual progress
    
                RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxPerFile)  ' maintain overall totals
                CalculateProgress mcurOA_CurrentAmt, mcurOA_MaxAmt, True        ' overall progress
                DoEvents
    
                If curAmtLeft < 1 Then
                    Exit Do   ' exit Do..Loop
                End If
                    
            Loop
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            ' Verify file handle is released
            mobjBigFiles.API_CloseFile hFile
    
            ' The pause is here to allow
            ' the write process to sync up
            ' with the speed of the CPU
            DoEvents
            lngPause = GetTickCount + 200
            Do While lngPause > GetTickCount
                DoEvents
            Loop
                    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
            CalculateProgress 100, 100, False                      ' individual progress

            Erase abytTemp()   ' Empty temp byte array
            Erase astrData()   ' Empty temp string array
                                
            If lngIndex < mlngPasses Then
                mcurCurrPerPass = mcurCurrPerPass - curFileSize
            End If
    
        Next lngIndex
    End With
    
    DoDPatterns = True
    DoEvents
    
    RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
    CalculateProgress 100, 100, False                      ' individual progress

DoDPatterns_CleanUp:
    ' Verify file handle is released
    mobjBigFiles.API_CloseFile hFile
    
    ' Update log file, if requested
    If gblnStopProcessing Then
        If mblnLogData Then
            strRecord = "ABENDED:" & vbTab & "DoD file wipe on " & strFileName & vbNewLine & _
                        Space$(32) & ABEND_MSG
            RaiseEvent UpdateLogData(strRecord)
        End If
    End If

    Erase abytRnd()    ' Always empty arrays when not needed
    Erase abytTemp()
    Erase abytVerify()
    Erase astrData()
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

DoDPatterns_Error:
    ' update log file, if requested
    If mblnLogData Then
        strRecord = vbTab & vbTab & "File:  " & strFileName & vbNewLine & _
                    vbTab & vbTab & vbTab & vbTab & _
                    CStr(Err.Number) & "  " & Err.Description
        RaiseEvent UpdateLogData(strRecord)
    End If

    strRecord = "ERROR File:  " & strFileName & vbNewLine & _
                CStr(Err.Number) & "  " & Err.Description
    ErrorMsg MODULE_NAME, ROUTINE_NAME, strRecord

    DoDPatterns = False
    Resume DoDPatterns_CleanUp

End Function

' **************************************************************************
' Routine:       EuroPatterns
'
' Description:   Loops thru an open file and overwrites each sector with a
'                pre-defined pattern.
'
'                North Atlantic Treaty Organization (NATO) will overwrite the
'                target data area 7 times.  For the first 6 passes, each wipe
'                reverses the bit pattern of previous wipe. The 7th overwrite
'                is with a random data stream.
'
'                NATO  Write 1 - One random character
'                      Write 2 - Complement of previous character
'
'                      Repeat steps 1-2 three times
'                      Write 7 - Random data stream
'                      (7 overwrites equal one pass)
'
'                Germany BSI Verschlusssachen-IT-Richtlinien (VSITR) Standard
'                The German Federal Office for IT Security released the VSITR
'                standard, which wipes the drive with seven passes. For the
'                first 6 passes, each wipe reverses the bit pattern of the
'                previous wipe.
'
'                Flipping the bits in this way is designed to destabilise the
'                remnants of data that may exist on the edges of the track of
'                the disk to which the data is written. The final pass
'                amplifies this effect by overwriting with 0x55.
'
'                German VSITR  Write 1 - One random character
'                              Write 2 - Complement of previous character
'
'                              Repeat steps 1-2 three times
'                              Write 7 - Data stream of 0x55 (ASCII 85)
'                              (7 overwrites equal one pass)
'
'                Verification will be checked to give you a true security
'                overwrite.  The final pass will be compared with the data
'                that was supposed to be written.
'
' Important:     This is good security for up thru "SECRET" documents only.
'
' Disposal:      If the hard disk has had data classified greater than
'                "CONFIDENTIAL", then the disk should be replaced with a new
'                one.  Since the cost of a hard disk has dropped so
'                dramatically, this should not be a factor.  You should be
'                considering the question, "What is my information worth to
'                someone else?".
'
'                Steps to follow to dispose of the old hard drive:
'                1. Overwrite multiple times with random data (Min 5 times).
'                   I recommend the Dban web site and creating a bootable CD
'                   or USB device that will wipe every sector on a disk. This
'                   is freeware and several governments approve its use.
'                   http://www.dban.org/
'                2. Remove disk from the old desktop or laptop and record the
'                   manufacturer, model, serial number, date of destruction and
'                   name of individual performing this process.
'                3. Plate area should be drilled in several places using a
'                   1/2 inch drill bit.
'                4. Disintegrate, incinerate, pulverize, shred, or melt the
'                   hard drive.
'
'                All of the above should be witnessed by at least two
'                additional persons and documented.
'
'                Ref:  Guidelines for Media Sanitization (SP800-88 dtd Sep-2006)
'                      http://csrc.nist.gov/publications/nistpubs/800-88/NISTSP800-88_rev1.pdf
'
' Parameters:    strFilename - Name of the file to be processed
'
' Returns:       TRUE - Successful completion
'                FALSE - something went wrong
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Jan-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 16-Jan-2008  Kenneth Ives  kenaso@tx.rr.com
'              The write process has been speeded up by 50% or more by
'              adjusting the record length based on amount of data left to
'              write.  This process calls the routine GetBlockSize().
' 24-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added option to write zeroes as the last write
' 14-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Simplified logic for when to create random data
' ***************************************************************************
Private Function EuroPatterns(ByVal strFileName As String) As Boolean

    Dim hFile         As Long      ' file handle
    Dim lngIdx        As Long      ' loop counter
    Dim lngPause      As Long
    Dim lngIndex      As Long      ' loop counter
    Dim lngWriteCnt   As Long      ' Number of overwrites to a file
    Dim lngMaxDataAmt As Long      ' Max amount of random data to create
    Dim lngNoOfWrites As Long      ' number of writes to a file to equal one overwrite
    Dim lngPointer    As Long      ' Rnd data string pointer
    Dim lngFilePos    As Long      ' Starting position in the file
    Dim lngRecordSize As Long      ' length of data record
    Dim blnLastPass   As Boolean   ' Flag designating last pass being performed
    Dim curFileSize   As Currency  ' original file size
    Dim curAmtLeft    As Currency  ' used to determine how much is left to write
    Dim curFilePos    As Currency  ' Starting position in the file
    Dim strRecord     As String    ' log file message
    Dim strPattern    As String    ' log info designating the pattern in use
    Dim astrData(8)   As String    ' data array to hold overwrite data
    Dim bytTemp1      As Byte
    Dim bytTemp2      As Byte
    Dim bytTemp3      As Byte
    Dim bytTemp4      As Byte
    Dim bytTemp5      As Byte
    Dim bytTemp6      As Byte
    Dim abytRnd()     As Byte      ' random data
    Dim abytTemp()    As Byte      ' Temp data
    Dim abytVerify()  As Byte      ' used to verify written data
    
    Const ROUTINE_NAME As String = "EuroPatterns"

    On Error GoTo EuroPatterns_Error

    Erase abytRnd()  ' make sure we start with empty arrays
    Erase abytTemp()
    Erase abytVerify()
    Erase astrData()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Determine if the last write is to be all zeroes
    lngNoOfWrites = IIf(mblnZeroLastWrite, 7, 6)
    
    lngWriteCnt = MAX_INT   ' Init file overwrite counter
    blnLastPass = False     ' Init last pass flag
    
    ' Reset file attributes
    SetFileAttributes strFileName, FILE_ATTRIBUTE_NORMAL
    
    mobjBigFiles.CalcFileSize strFileName, curFileSize   ' Get file size

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo EuroPatterns_CleanUp
    End If

    With mobjRandom
        
        Rnd -1               ' Reset VB random number generator
        Randomize .RndSeed   ' Reseed VB random number generator
        
        For lngIndex = 1 To mlngPasses
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            DoEvents
            RaiseEvent CountPasses(lngIndex, mlngPasses)
            RaiseEvent CurrentProgress(0, curFileSize)    ' maintain the ongoing totals
            CalculateProgress 0, curFileSize, False       ' individual progress
            DoEvents
    
            If lngIndex = mlngPasses Then
                blnLastPass = True
            End If
            
            lngFilePos = 1                             ' Set pointer to beginning of file
            curFilePos = 0                             ' Set pointer to beginning of file
            curAmtLeft = curFileSize                   ' Work with copy of file size
            lngRecordSize = GetBlockSize(curAmtLeft)   ' determine record size of output record
            lngMaxDataAmt = KB_4 + lngRecordSize       ' Max amount of random data to create
                    
            ' Open target file
            If Not mobjBigFiles.OpenReadWrite(strFileName, hFile) Then
                gblnStopProcessing = True
            End If
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            Do
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
        
                Erase abytTemp()   ' Always start with an empty array
                Erase astrData()
                    
                If lngRecordSize > 1 Then
    
                    ' After 50 overwrites, generate new data
                    If lngWriteCnt >= 50 Then
                        lngWriteCnt = 0                                     ' reset overwrite counter
                        Erase abytRnd()                                     ' Always start with an empty array
                        abytRnd() = .BuildRndData(lngMaxDataAmt, , False)   ' Fill a byte array with random data
                    End If
                    
                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        Exit Do   ' exit Do..Loop
                    End If

                    lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                    bytTemp1 = abytRnd(lngPointer)                  ' Capture one byte of random data (Ex:  85 = 0x55 = 01010101)
                    bytTemp2 = GetComplement(bytTemp1)              ' Get complement of previous byte (Ex: 170 = 0xAA = 10101010)
                    astrData(0) = String$(lngRecordSize, bytTemp1)  ' Create string using random byte
                    astrData(1) = String$(lngRecordSize, bytTemp2)  ' Create string using complement byte
                    
                    lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                    bytTemp3 = abytRnd(lngPointer)                  ' Capture one byte of random data (Ex:  85 = 0x55 = 01010101)
                    bytTemp4 = GetComplement(bytTemp3)              ' Get complement of previous byte (Ex: 170 = 0xAA = 10101010)
                    astrData(2) = String$(lngRecordSize, bytTemp3)  ' Create string using random byte
                    astrData(3) = String$(lngRecordSize, bytTemp4)  ' Create string using complement byte
                    
                    lngPointer = Int(Rnd() * KB_4)                  ' Create a starting point between 0-4095
                    bytTemp5 = abytRnd(lngPointer)                  ' Capture one byte of random data (Ex:  85 = 0x55 = 01010101)
                    bytTemp6 = GetComplement(bytTemp5)              ' Get complement of previous byte (Ex: 170 = 0xAA = 10101010)
                    astrData(4) = String$(lngRecordSize, bytTemp5)  ' Create string using random byte
                    astrData(5) = String$(lngRecordSize, bytTemp6)  ' Create string using complement byte
                                
                    Select Case mlngWipeMethod
                           
                           Case eWipe_NATO    ' NATO  (7 overwrites)
                                ' Fill with random generated data
                                lngPointer = Int(Rnd() * KB_4)                ' Create a starting point between 0-4095
                                ReDim abytTemp(lngRecordSize)                 ' Size temp array
                                CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                                ReDim Preserve abytTemp(lngRecordSize - 1)    ' Re-Size temp array
                                astrData(6) = ByteArrayToString(abytTemp())   ' Convert temp array to string data
                    
                           Case eWipe_GermanVISTR    ' German VSITR  (7 overwrites)
                                astrData(6) = String$(lngRecordSize, 85)      ' Stream of 0x55 values (Binary 01010101)
                                                                 
                    End Select
                    
                    ' Optional last write using zeroes
                    If mblnZeroLastWrite Then
                        astrData(7) = String$(lngRecordSize, 0)   ' Stream of 0x00 values (Binary 00000000)
                    End If
                    
                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        Exit Do      ' exit this DO loop
                    End If
                              
                    ' overwrite this position multiple times
                    For lngIdx = 0 To lngNoOfWrites
                    
                        Erase abytTemp()                                  ' Verify array is empty
                        abytTemp() = StringToByteArray(astrData(lngIdx))  ' Convert string dta to byte array
                                                
                        ' Excessive number of DoEvents is to allow the read and
                        ' write process to stay in sync with the application speed.
                        '
                        ' Write data to file
                        DoEvents
                        If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                            gblnStopProcessing = True
                        End If
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
                    
                        ' Verification has been requested then read the record
                        ' into a buffer and verify it's contents against what was
                        ' supposed to be written.
                        If blnLastPass Then
                        
                            ' Type of data used for error reporting
                            Select Case lngIdx
                                   Case 0: strPattern = "0x" & Right$("0" & Hex$(bytTemp1), 2)
                                   Case 1: strPattern = "0x" & Right$("0" & Hex$(bytTemp2), 2)
                                   Case 2: strPattern = "0x" & Right$("0" & Hex$(bytTemp3), 2)
                                   Case 3: strPattern = "0x" & Right$("0" & Hex$(bytTemp4), 2)
                                   Case 4: strPattern = "0x" & Right$("0" & Hex$(bytTemp5), 2)
                                   Case 5: strPattern = "0x" & Right$("0" & Hex$(bytTemp6), 2)
                                   Case 6: strPattern = IIf(mlngWipeMethod = eWipe_NATO, "Random", "0x55")
                                   Case Else: strPattern = "0x00"
                            End Select
                             
                            ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                            
                            ' Capture the data that was just written to the file
                            DoEvents
                            If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                gblnStopProcessing = True
                                Exit For    ' exit For..Next loop
                            End If
                        
                            If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                   strFileName, curFilePos, strPattern) Then
                                                   
                                gblnStopProcessing = True
                                Exit For    ' exit For..Next loop
                            End If
                                    
                        End If
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
                    
                    Next lngIdx
                    
                End If
    
                Erase abytTemp()   ' Empty temp array
                
                 ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do      ' exit this DO loop
                End If
                
                DoEvents
                curFilePos = curFilePos + CCur(lngRecordSize)          ' Adjust pointers accordingly
                lngWriteCnt = lngWriteCnt + 1                          ' update overwrite counter
                curAmtLeft = curAmtLeft - lngRecordSize                ' calc how much is left
                mcurCurrPerPass = mcurCurrPerPass + lngRecordSize      ' Current amt written this pass
                mcurOA_CurrentAmt = mcurOA_CurrentAmt + lngRecordSize  ' calc overall progress for pbar
                lngRecordSize = GetBlockSize(curAmtLeft)               ' determine record size of output record
    
                DoEvents
                RaiseEvent CurrentProgress(curFilePos, curFileSize)    ' individual ongoing totals
                CalculateProgress curFilePos, curFileSize, False       ' individual progress
    
                RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxPerFile)  ' maintain overall totals
                CalculateProgress mcurOA_CurrentAmt, mcurOA_MaxAmt, True        ' overall progress
                DoEvents
    
                If curAmtLeft < 1 Then
                    Exit Do   ' exit Do..Loop
                End If
                        
            Loop
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            ' Verify file handle is released
            mobjBigFiles.API_CloseFile hFile
    
            ' The pause is here to allow
            ' the write process to sync up
            ' with the speed of the CPU
            DoEvents
            lngPause = GetTickCount + 200
            Do While lngPause > GetTickCount
                DoEvents
            Loop
                    
            RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
            CalculateProgress 100, 100, False                      ' individual progress
            
            Erase abytTemp()   ' Empty temp byte array
            Erase astrData()   ' Empty temp string array
                                
            If lngIndex < mlngPasses Then
                mcurCurrPerPass = mcurCurrPerPass - curFileSize
            End If
    
        Next lngIndex
    End With
    
    EuroPatterns = True
    DoEvents
    
    RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
    CalculateProgress 100, 100, False                      ' individual progress

EuroPatterns_CleanUp:
    ' Verify file handle is released
    mobjBigFiles.API_CloseFile hFile

    ' Update log file, if requested
    If gblnStopProcessing Then
        If mblnLogData Then
            strRecord = "ABENDED:" & vbTab & "Euro file wipe on " & strFileName & vbNewLine & _
                        Space$(32) & ABEND_MSG
            RaiseEvent UpdateLogData(strRecord)
        End If
    End If

    Erase abytRnd()    ' Always empty arrays when not needed
    Erase abytTemp()
    Erase abytVerify()
    Erase astrData()

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

EuroPatterns_Error:
    ' Update log file, if requested
    If mblnLogData Then
        strRecord = vbTab & vbTab & "File:  " & strFileName & vbNewLine & _
                    vbTab & vbTab & vbTab & vbTab & _
                    CStr(Err.Number) & "  " & Err.Description
        RaiseEvent UpdateLogData(strRecord)
    End If

    strRecord = "ERROR File:  " & strFileName & vbNewLine & _
                CStr(Err.Number) & "  " & Err.Description
    ErrorMsg MODULE_NAME, ROUTINE_NAME, strRecord

    EuroPatterns = False
    Resume EuroPatterns_CleanUp

End Function

' **************************************************************************
' Routine:       GutmannPattern
'
' Description:   Loops thru an open file and overwrites each byte of data
'                with a pre-defined pattern.  This process performs 35
'                writes as decribed below.
'
'                Based on Peter Gutmann`s paper "Secure Deletion of Data from
'                Magnetic and Solid-State Memory", Peter Gutmann 1996,
'                http://www.cs.auckland.ac.nz/~pgut001/secure_del.html
'
'                Recommended wiping is done by using the following sequence of
'                35 consecutive writes to erase data: (overwriting values are
'                given in hex except for random data)
'
'                   1.  Random             13.  0x33       25.  0xFF
'                   2.  Random             14.  0x44       26.  0x92 0x49 0x24
'                   3.  Random             15.  0x55       27.  0x49 0x24 0x92
'                   4.  Random             16.  0x66       28.  0x24 0x92 0x49
'                   5.  0x55               17.  0x77       29.  0x6D 0xB6 0xDB
'                   6.  0xAA               18.  0x88       30.  0xB6 0xDB 0x6D
'                   7.  0x92 0x49 0x24     19.  0x99       31.  0xDB 0x6D 0xB6
'                   8.  0x49 0x24 0x92     20.  0xAA       32.  Random
'                   9.  0x24 0x92 0x49     21.  0xBB       33.  Random
'                  10.  0x00               22.  0xCC       34.  Random
'                  11.  0x11               23.  0xDD       35.  Random
'                  12.  0x22               24.  0xEE
'
'                  (35 overwrites equal one pass)
'
'                ===============================================================
'
'                Most of the overwrites in the Gutmann wipe are designed to flip
'                bits in MFM/RLL encoded disks, which is an encoding that modern
'                hard disks do not use.
'
'                Gutmann himself has responded to some of these criticisms and
'                also criticized how his algorithm has been abused in an epilogue
'                to his original paper, in which he states:
'
'                "In the time since this paper was published, some people have
'                treated the 35-pass overwrite technique described in it more
'                as a kind of voodoo incantation to banish evil spirits than
'                the result of a technical analysis of drive encoding
'                techniques. As a result, they advocate applying the voodoo to
'                PRML and EPRML drives even though it will have no more effect
'                than a simple scrubbing with random data. In fact performing
'                the full 35-pass overwrite is pointless for any drive since it
'                targets a blend of scenarios involving all types of
'                (normally-used) encoding technology, which covers everything
'                back to 30+-year-old MFM methods (if you don't understand that
'                statement, re-read the paper). If you're using a drive which
'                uses encoding technology X, you only need to perform the
'                passes specific to X, and you never need to perform all 35
'                passes. For any modern PRML/EPRML drive, a few passes of
'                random scrubbing is the best you can do. As the paper says,
'                'A good scrubbing with random data will do about as well as
'                can be expected'. This was true in 1996, and is still true now."
'
' References:    Peter Gutmann home page
'                http://www.cs.auckland.ac.nz/~pgut001/
'
'                Wikipedia, the free encyclopedia
'                http://wapedia.mobi/en/Gutmann_method
'                http://en.wikipedia.org/wiki/Data_remanence
'
'                Secure Deletion of Data from Magnetic and Solid-State Memory
'                http://www.cs.auckland.ac.nz/~pgut001/pubs/secure_del.html
'
' Note:          If verification is requested then the final pass will
'                be compared with the data that was supposed to be written.
'
' Important:     This is good security for up thru "SECRET" documents only.
'
' Disposal:      If the hard disk has had data classified greater than
'                "CONFIDENTIAL", then the disk should be replaced with a new
'                one.  Since the cost of a hard disk has dropped so
'                dramatically, this should not be a factor.  You should be
'                considering the question, "What is my information worth to
'                someone else?".
'
'                Steps to follow to dispose of the old hard drive:
'                1. Overwrite multiple times with random data (Min 5 times).
'                   I recommend the Dban web site and creating a bootable CD
'                   or USB device that will wipe every sector on a disk. This
'                   is freeware and several governments approve its use.
'                   http://www.dban.org/
'                2. Remove disk from the old desktop or laptop and record the
'                   manufacturer, model, serial number, date of destruction and
'                   name of individual performing this process.
'                3. Plate area should be drilled in several places using a
'                   1/2 inch drill bit.
'                4. Disintegrate, incinerate, pulverize, shred, or melt the
'                   hard drive.
'
'                All of the above should be witnessed by at least two
'                additional persons and documented.
'
'                Ref:  Guidelines for Media Sanitization (SP800-88 dtd Sep-2006)
'                      http://csrc.nist.gov/publications/nistpubs/800-88/NISTSP800-88_rev1.pdf
'
' Parameters:    strFilename - Name of the file to be processed
'
' Returns:       TRUE - Successful completion
'                FALSE - something went wrong
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 25-FEB-2005  Kenneth Ives  kenaso@tx.rr.com
'              Updated how random data was to be determined for the output
'              record.  This speeded up the process considerably.
' 02-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in data array that referenced multiple values
' 14-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              The write process has been speeded up by 50% or more by
'              adjusting the record length based on amount of data left to
'              write.  This process calls the routine GetBlockSize().
' 24-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added option to write zeroes as the last write
' 14-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Simplified logic for when to create random data
' 14-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in creating a record length and loading
'              data.
' 15-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Verify temp arrays are empty when not needed.  Save memory
'              resources.
' ***************************************************************************
Private Function GutmannPattern(ByVal strFileName As String) As Boolean

    Dim hFile         As Long      ' file handle
    Dim lngIndex      As Long      ' loop counter
    Dim lngIdx        As Long      ' loop counter
    Dim lngPause      As Long
    Dim lngPointer    As Long      ' Random data string pointer
    Dim lngFilePos    As Long      ' Starting position in the file
    Dim lngWriteCnt   As Long      ' Number of overwrites to a file
    Dim lngNoOfWrites As Long      ' number of writes to a file to equal one overwrite
    Dim lngRecordSize As Long      ' length of data record
    Dim blnLastPass   As Boolean   ' Flag designating last pass being performed
    Dim curFileSize   As Currency  ' original file size
    Dim curAmtLeft    As Currency  ' used to determine how much is left to write
    Dim curFilePos    As Currency  ' Starting position in the file
    Dim strRecord     As String    ' log file message
    Dim strPattern    As String    ' log info designating the pattern in use
    Dim astrThree()   As String    ' Three character patterns
    Dim abytOne()     As Byte      ' Single character patterns
    Dim abytRnd()     As Byte      ' random data
    Dim abytTemp()    As Byte      ' Temp hold array
    Dim abytVerify()  As Byte      ' used to verify written data

    Const ROUTINE_NAME As String = "GutmannPattern"

    On Error GoTo GutmannPattern_Error

    Erase abytOne()     ' Always start with empty arrays
    Erase abytRnd()
    Erase abytTemp()
    Erase abytVerify()
    Erase astrThree()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Load specific elements only
    ReDim abytOne(25)      ' Hex    Binary
                           ' --------------
    abytOne(4) = 85        '  55   01010101
    abytOne(5) = 170       '  AA   10101010
    abytOne(9) = 0         '  00   00000000
    abytOne(10) = 17       '  11   00010001
    abytOne(11) = 34       '  22   00100010
    abytOne(12) = 51       '  33   00110011
    abytOne(13) = 68       '  44   01000100
    abytOne(14) = 85       '  55   01010101
    abytOne(15) = 102      '  66   01100110
    abytOne(16) = 119      '  77   01110111
    abytOne(17) = 136      '  88   10001000
    abytOne(18) = 153      '  99   10011001
    abytOne(19) = 170      '  AA   10101010
    abytOne(20) = 187      '  BB   10111011
    abytOne(21) = 204      '  CC   11001100
    abytOne(22) = 221      '  DD   11011101
    abytOne(23) = 238      '  EE   11101110
    abytOne(24) = 255      '  FF   11111111
    
    ' Load 3 character string patterns
    ReDim astrThree(6)  ' ASCII TEXT                         HEX                   BINARY
    astrThree(0) = Chr$(146) & Chr$(73) & Chr$(36)    ' 0x92 0x49 0x24 = 10010010 01001001 00100100
    astrThree(1) = Chr$(73) & Chr$(36) & Chr$(146)    ' 0x49 0x24 0x92 = 01001001 00100100 10010010
    astrThree(2) = Chr$(36) & Chr$(146) & Chr$(73)    ' 0x24 0x92 0x49 = 00100100 10010010 01001001
    astrThree(3) = Chr$(109) & Chr$(182) & Chr$(219)  ' 0x6D 0xB6 0xDB = 01101101 10110110 11011011
    astrThree(4) = Chr$(182) & Chr$(219) & Chr$(109)  ' 0xB6 0xDB 0x6D = 10110110 11011011 01101101
    astrThree(5) = Chr$(219) & Chr$(109) & Chr$(182)  ' 0xDB 0x6D 0xB6 = 11011011 01101101 10110110
    
    lngNoOfWrites = IIf(mblnZeroLastWrite, 35, 34)    ' Determine if the last write is to be all zeroes
    
    lngWriteCnt = MAX_INT   ' Init number of overwrites
    blnLastPass = False     ' Init last pass flag
    
    ' Reset file attributes
    SetFileAttributes strFileName, FILE_ATTRIBUTE_NORMAL
    
    mobjBigFiles.CalcFileSize strFileName, curFileSize   ' Get file size

    With mobjRandom
        
        Rnd -1               ' Reset VB random number generator
        Randomize .RndSeed   ' Reseed VB random number generator
        
        For lngIndex = 1 To mlngPasses
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            DoEvents
            RaiseEvent CountPasses(lngIndex, mlngPasses)
            RaiseEvent CurrentProgress(0, curFileSize)    ' maintain the ongoing totals
            CalculateProgress 0, curFileSize, False       ' individual progress
            DoEvents
    
            If lngIndex = mlngPasses Then
                blnLastPass = True
            End If
            
            lngFilePos = 1            ' Set pointer to beginning of file
            curFilePos = 0            ' Set pointer to beginning of file
            curAmtLeft = curFileSize  ' Work with copy of file size
                    
            ' Due to the amount of memory resources being
            ' used by this pattern mixture, I have elected
            ' to limit the record size to 4096 bytes.
            Select Case curAmtLeft
                   Case Is >= KB_4: lngRecordSize = KB_4
                   Case Else:       lngRecordSize = CLng(curAmtLeft)
            End Select
                                
            ' Open target file
            If Not mobjBigFiles.OpenReadWrite(strFileName, hFile) Then
                gblnStopProcessing = True
            End If
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            Do
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
        
                Erase abytTemp()    ' Always start with an empty array
                Erase abytVerify()
                                
                If lngRecordSize > 1 Then
    
                    ' After 50 overwrites, generate new data
                    If lngWriteCnt >= 50 Then
                        lngWriteCnt = 0                                   ' reset overwrite counter
                        Erase abytRnd()                                   ' Always start with an empty array
                        abytRnd() = .BuildRndData(KB_4 + KB_4, , False)   ' Fill a byte array with random data
                    End If
                    
                    ' An error occurred or user opted to STOP processing
                    DoEvents
                    If gblnStopProcessing Then
                        Exit Do   ' exit Do..Loop
                    End If
    
                    ' Write data records on top of each other
                    For lngIdx = 0 To lngNoOfWrites
                        
                        Select Case lngIdx
                                
                               Case 0 To 3, 31 To 34  ' Random data stream
                                    lngPointer = Int(Rnd() * KB_4)               ' Create a starting point between 0-4095
                                    ReDim abytTemp(lngRecordSize)                ' Size temp array
                                    CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                                    ReDim Preserve abytTemp(lngRecordSize - 1)   ' Re-Size temp array
                                    strRecord = ByteArrayToString(abytTemp())    ' Convert byte array to string data
                                    Erase abytTemp()                             ' Verify temp array is empty
                                    
                               Case 6, 25  ' Three bytes (0x92 0x49 0x24)
                                    strRecord = Replicate(Int(lngRecordSize / 2.99), astrThree(0))
                                    strRecord = Left$(strRecord, lngRecordSize)
                               
                               Case 7, 26  ' Three bytes (0x49 0x24 0x92)
                                    strRecord = Replicate(Int(lngRecordSize / 2.99), astrThree(1))
                                    strRecord = Left$(strRecord, lngRecordSize)
                               
                               Case 8, 27  ' Three bytes (0x24 0x92 0x49)
                                    strRecord = Replicate(Int(lngRecordSize / 2.99), astrThree(2))
                                    strRecord = Left$(strRecord, lngRecordSize)
                               
                               Case 28  ' Three bytes (0x6D 0xB6 0xDB)
                                    strRecord = Replicate(Int(lngRecordSize / 2.99), astrThree(3))
                                    strRecord = Left$(strRecord, lngRecordSize)
                               
                               Case 29  ' Three bytes (0xB6 0xDB 0x6D)
                                    strRecord = Replicate(Int(lngRecordSize / 2.99), astrThree(4))
                                    strRecord = Left$(strRecord, lngRecordSize)
                               
                               Case 30  ' Three bytes (0xDB 0x6D 0xB6)
                                    strRecord = Replicate(Int(lngRecordSize / 2.99), astrThree(5))
                                    strRecord = Left$(strRecord, lngRecordSize)
                               
                               Case 35  ' Optional last write of binary zeroes
                                    strRecord = String$(lngRecordSize, 0)
                                    
                               Case Else   ' Single byte replicated
                                    strRecord = String$(lngRecordSize, abytOne(lngIdx))
                        End Select
                                
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
                        
                        abytTemp() = StringToByteArray(strRecord)   ' Convert data to byte array
                        strRecord = vbNullString                              ' Empty string variable
                                                                 
                        ' Excessive number of DoEvents is to allow the read and
                        ' write process to stay in sync with the application speed.
                        '
                        ' Write data to file
                        DoEvents
                        If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                            gblnStopProcessing = True
                        End If
                        
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
                        
                        ' If verification has been requested then read the record
                        ' into a buffer and verify it's contents against what was
                        ' supposed to be written.
                        If mblnVerifyData Then
                            If blnLastPass Then
                    
                                ' Type of data used for error reporting
                                Select Case lngIdx
                                       Case 0 To 3, 31 To 34: strPattern = "Random"
                                       Case 4, 14: strPattern = "0x55"
                                       Case 5, 19: strPattern = "0xAA"
                                       Case 6, 25: strPattern = "0x92 0x49 0x24"
                                       Case 7, 26: strPattern = "0x49 0x24 0x92"
                                       Case 8, 27: strPattern = "0x24 0x92 0x49"
                                       Case 9, 35: strPattern = "0x00"
                                       Case 10: strPattern = "0x11"
                                       Case 11: strPattern = "0x22"
                                       Case 12: strPattern = "0x33"
                                       Case 13: strPattern = "0x44"
                                       Case 14: strPattern = "0x55"
                                       Case 15: strPattern = "0x66"
                                       Case 16: strPattern = "0x77"
                                       Case 17: strPattern = "0x88"
                                       Case 18: strPattern = "0x99"
                                       Case 19: strPattern = "0xAA"
                                       Case 20: strPattern = "0xBB"
                                       Case 21: strPattern = "0xCC"
                                       Case 22: strPattern = "0xDD"
                                       Case 23: strPattern = "0xEE"
                                       Case 24: strPattern = "0xFF"
                                       Case 28: strPattern = "0x6D 0xB6 0xDB"
                                       Case 29: strPattern = "0xB6 0xDB 0x6D"
                                       Case 30: strPattern = "0xDB 0x6D 0xB6"
                                End Select
                                
                                ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                    
                                ' Capture the data that was just written to the file
                                DoEvents
                                If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                    gblnStopProcessing = True
                                    Exit For    ' exit For..Next loop
                                End If
                            
                                If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                       strFileName, curFilePos, strPattern) Then
                                                       
                                    gblnStopProcessing = True
                                    Exit For    ' exit For..Next loop
                                End If
                                    
                            End If
                        End If
                    
                        Erase abytTemp()   ' Verify arrays are empty
                        Erase abytVerify()
                        
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
                    
                    Next lngIdx
                
                End If
    
                Erase abytTemp()   ' Empty temp array
                
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
                    
                DoEvents
                curFilePos = curFilePos + CCur(lngRecordSize)          ' Adjust pointers accordingly
                lngWriteCnt = lngWriteCnt + 1                          ' update overwrite counter
                curAmtLeft = curAmtLeft - lngRecordSize                ' calc how much is left
                mcurCurrPerPass = mcurCurrPerPass + lngRecordSize      ' Current amt written this pass
                mcurOA_CurrentAmt = mcurOA_CurrentAmt + lngRecordSize  ' calc overall progress for pbar
    
                ' Due to the amount of memory resources being
                ' used by this pattern mixture, I have elected
                ' to limit the record size to 4096 bytes.
                Select Case curAmtLeft
                       Case Is >= KB_4: lngRecordSize = KB_4
                       Case Else:       lngRecordSize = CLng(curAmtLeft)
                End Select
            
                DoEvents
                RaiseEvent CurrentProgress(curFilePos, curFileSize)    ' individual ongoing totals
                CalculateProgress curFilePos, curFileSize, False       ' individual progress
    
                RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxPerFile)  ' maintain overall totals
                CalculateProgress mcurOA_CurrentAmt, mcurOA_MaxAmt, True        ' overall progress
                DoEvents
    
                If curAmtLeft < 1 Then
                    Exit Do   ' exit Do..Loop
                End If
                    
            Loop
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            ' Verify file handle is released
            mobjBigFiles.API_CloseFile hFile
    
            ' The pause is here to allow
            ' the write process to sync up
            ' with the speed of the CPU
            DoEvents
            lngPause = GetTickCount + 200
            Do While lngPause > GetTickCount
                DoEvents
            Loop
                    
            RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
            CalculateProgress 100, 100, False                      ' individual progress

            Erase abytTemp()     ' Empty temp byte array
            Erase abytVerify()
                                
            If lngIndex < mlngPasses Then
                mcurCurrPerPass = mcurCurrPerPass - curFileSize
            End If
    
        Next lngIndex
    End With
    
    GutmannPattern = True
    DoEvents
    
    RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
    CalculateProgress 100, 100, False                      ' individual progress

GutmannPattern_CleanUp:
    ' Verify file handle is released
    mobjBigFiles.API_CloseFile hFile
    
    Erase abytRnd()    ' Always empty arrays when not needed
    Erase abytTemp()
    Erase abytVerify()
    Erase abytOne()
    Erase astrThree()
    
    ' Update log file, if requested
    If gblnStopProcessing Then
        If mblnLogData Then
            strRecord = "ABENDED:" & vbTab & "Gutmann file wipe on " & strFileName & vbNewLine & _
                        Space$(32) & ABEND_MSG
            RaiseEvent UpdateLogData(strRecord)
        End If
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

GutmannPattern_Error:
    ' Update log file, if requested
    If mblnLogData Then
        strRecord = vbTab & vbTab & "File:  " & strFileName & vbNewLine & _
                    vbTab & vbTab & vbTab & vbTab & _
                    CStr(Err.Number) & "  " & Err.Description
        RaiseEvent UpdateLogData(strRecord)
    End If

    strRecord = "ERROR File:  " & strFileName & vbNewLine & _
                CStr(Err.Number) & "  " & Err.Description
    ErrorMsg MODULE_NAME, ROUTINE_NAME, strRecord

    GutmannPattern = False
    Resume GutmannPattern_CleanUp

End Function

' ***************************************************************************
' Routine:       Replicate
'
' Description:   Replicates a string of data.
'
' Parameters:    lngRetLength - Length of the return string
'                strPattern - Data to be replicated
'
' Returns:       Data string with replicated pattern
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Dec-2006  Peter Weighill, pweighill@btinternet.com
'              http://www.xbeat.net/vbspeed/c_Replicate.htm
' 02-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              Modifed and documented
' ***************************************************************************
Private Function Replicate(ByVal lngRetLength As Long, _
                           ByVal strPattern As String) As String
                          
    ' Called by GutmannPattern()
    
    Dim lngPatternLen As Long
    
    Replicate = vbNullString   ' Verify empty return string
    
    ' Verify length desired
    ' is greater than zero
    If lngRetLength < 1 Then
        Exit Function
    End If
        
    lngPatternLen = Len(strPattern)   ' Capture pattern length
    
    ' Check pattern length
    If lngPatternLen < 1 Then
        GoTo Replicate_CleanUp
    End If
            
    ' Preload return string with first character in pattern
    Replicate = String$(lngRetLength * lngPatternLen, strPattern)
    
    ' See if length of pattern
    ' is more than one character
    If lngPatternLen > 1 Then
        
        ' Insert a copy of pattern starting in first position
        Mid$(Replicate, 1, lngPatternLen) = strPattern
        
        ' See if return length is greater than pattern length
        If lngRetLength > lngPatternLen Then
            
            ' Replicate pattern to end of allocated string
            Mid$(Replicate, lngPatternLen + 1) = Replicate
        End If
    
    End If

Replicate_CleanUp:
    ' If there is a problem then
    ' return a string of null values
    If Len(Trim$(Replicate)) = 0 Then
        Replicate = String$(lngRetLength, 0)
    End If

End Function

' **************************************************************************
' Routine:       EncryptPattern
'
' Description:   Loops thru an open file and performs encryption using
'                one of four strong encryption algorithms.  they are:
'
'                     - Rijndael [US Advanced Encryption Standard (AES)]
'                     - Blowfish [Strong encryption algorithm]
'                     - Twofish  [Finalist in NIST encryption contest]
'                     - ArcFour  [Strong encryption algorithm]
'
'                Uses a random generated password 15-50 bytes long
'                consisting of ASCII decimal values 0-255. Key length
'                and block size are also randomly selected. This selection
'                process takes place prior to each pass.
'
' Important:     This is good security for up thru "SECRET" documents only.
'
' Disposal:      If the hard disk has had data classified greater than
'                "CONFIDENTIAL", then the disk should be replaced with a new
'                one.  Since the cost of a hard disk has dropped so
'                dramatically, this should not be a factor.  You should be
'                considering the question, "What is my information worth to
'                someone else?".
'
'                Steps to follow to dispose of the old hard drive:
'                1. Overwrite multiple times with random data (Min 5 times).
'                   I recommend the Dban web site and creating a bootable CD
'                   or USB device that will wipe every sector on a disk. This
'                   is freeware and several governments approve its use.
'                   http://www.dban.org/
'                2. Remove disk from the old desktop or laptop and record the
'                   manufacturer, model, serial number, date of destruction and
'                   name of individual performing this process.
'                3. Plate area should be drilled in several places using a
'                   1/2 inch drill bit.
'                4. Disintegrate, incinerate, pulverize, shred, or melt the
'                   hard drive.
'
'                All of the above should be witnessed by at least two
'                additional persons and documented.
'
'                Ref:  Guidelines for Media Sanitization (SP800-88 dtd Sep-2006)
'                      http://csrc.nist.gov/publications/nistpubs/800-88/NISTSP800-88_rev1.pdf
'
' Parameters:    strFilename - Name of the file to be processed
'
' Returns:       TRUE  - Successful completion
'                FALSE - Something went wrong
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine to use Rijndael (AES) cipher class
' 02-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added three more encryption options
' ***************************************************************************
Private Function EncryptPattern(ByVal strFileName As String) As Boolean

    Dim strRecord    As String    ' log file message
    Dim lngKeyLength As Long      ' Random generated key length
    Dim lngBlockSize As Long      ' Random generated AES block size
    Dim abytPwd()    As Byte      ' Random generated password
    Dim abytFile()   As Byte      ' Holds path\filename
    
    Const ROUTINE_NAME As String = "EncryptPattern"

    On Error GoTo EncryptPattern_Error

    Erase abytPwd()   ' Always start with empty arrays
    Erase abytFile()
    
    ' Verify encryption class objects are not active
    Set mobjRijndael = Nothing
    Set mobjBlowfish = Nothing
    Set mobjTwofish = Nothing
    Set mobjArcFour = Nothing
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Used in object progress (see bottom of this module)
    mcurFilesize = 0@
    mcurByteCount = 0@
    
    ' If class object is not
    ' active then activate it
    If mobjBigFiles Is Nothing Then
        Set mobjBigFiles = New cBigFiles
    End If
    
    SetFileAttributes strFileName, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes
    mobjBigFiles.CalcFileSize strFileName, mcurFilesize   ' Get file size
    abytFile() = StringToByteArray(strFileName)           ' Convert path\filename to byte array
    
    ' Free class object from
    ' memory (Save resources).
    ' Will restart later.
    Set mobjBigFiles = Nothing
        
    ' Determine which class
    ' needs to be instantiated
    Select Case mlngEncryptAlgo
           Case eAlgo_Rijndael: Set mobjRijndael = New cRijndael
           Case eAlgo_Blowfish: Set mobjBlowfish = New cBlowFish
           Case eAlgo_Twofish:  Set mobjTwofish = New cTwofish
           Case eAlgo_ArcFour:  Set mobjArcFour = New cArcFour
    End Select
    
    DoEvents
    RaiseEvent CountPasses(1, 1)                  ' Display current pass
    RaiseEvent CurrentProgress(0, mcurFilesize)   ' maintain ongoing totals
    CalculateProgress 0, mcurFilesize, False      ' individual progress
    DoEvents
    
    ' Get new password, key length, block size
    ChangeAlgorithms abytPwd(), lngKeyLength, lngBlockSize
            
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo EncryptPattern_CleanUp
    End If

    If mblnLogEncryptParms Then
        strRecord = Space$(15) & "Parameters:  Pwd len-" & Format$(UBound(abytPwd), "!@@@@@") & _
                                 "Key len-" & Format$(lngKeyLength, "!@@@@@")
    End If
    
    ' Update property values and perform encryption
    Select Case mlngEncryptAlgo
           
           Case eAlgo_Rijndael
                ' Update log file with parameters
                DoEvents
                If mblnLogEncryptParms Then
                    RaiseEvent UpdateLogData(strRecord & "Block size-" & CStr(lngBlockSize))
                End If
               
                ' Pass parameters and perform encryption
                With mobjRijndael
                    .CreateNewFile = False      ' Overwrite current file
                    .Password = abytPwd()       ' password
                    .BlockSize = lngBlockSize   ' block size
                    .KeyLength = lngKeyLength   ' key length
                    .Rijndael_File abytFile()   ' Path\filename to be encrypted
                End With
    
           Case eAlgo_Blowfish
                ' Update log file with parameters
                DoEvents
                If mblnLogEncryptParms Then
                    RaiseEvent UpdateLogData(strRecord)
                End If
               
                ' Pass parameters and perform encryption
                With mobjBlowfish
                    .CreateNewFile = False      ' Overwrite current file
                    .Password = abytPwd()       ' password
                    .KeyLength = lngKeyLength   ' key length
                    .Blowfish_File abytFile()   ' Path\filename to be encrypted
                End With
           
           Case eAlgo_Twofish
                ' Update log file with parameters
                DoEvents
                If mblnLogEncryptParms Then
                    RaiseEvent UpdateLogData(strRecord)
                End If
               
                ' Pass parameters and perform encryption
                With mobjTwofish
                    .CreateNewFile = False      ' Overwrite current file
                    .Password = abytPwd()       ' password
                    .KeyLength = lngKeyLength   ' key length
                    .Twofish_File abytFile()    ' Path\filename to be encrypted
                End With
    
           Case eAlgo_ArcFour
                ' Update log file with parameters
                DoEvents
                If mblnLogEncryptParms Then
                    RaiseEvent UpdateLogData(strRecord)
                End If
               
                ' Pass parameters and perform encryption
                With mobjArcFour
                    .CreateNewFile = False      ' Overwrite current file
                    .Password = abytPwd()       ' password
                    .KeyLength = lngKeyLength   ' key length
                    .ArcFour_File abytFile()    ' Path\filename to be encrypted
                End With
    End Select
        
EncryptPattern_CleanUp:
    Erase abytPwd()    ' Always empty arrays when not needed
    Erase abytFile()
    
    ' Free class objects from memory
    Set mobjRijndael = Nothing
    Set mobjBlowfish = Nothing
    Set mobjTwofish = Nothing
    Set mobjArcFour = Nothing
    
    Set mobjBigFiles = New cBigFiles  ' Instantiate class object (Needed elsewhere)
    
    DoEvents
    If gblnStopProcessing Then
        
        ' Update log file, if requested
        If mblnLogData Then
            strRecord = "ABENDED:" & vbTab & "Encryption file wipe on " & strFileName & vbNewLine & _
                        Space$(32) & ABEND_MSG
            RaiseEvent UpdateLogData(strRecord)
        End If
    
        EncryptPattern = False   ' An error occurred or user opted to STOP processing
    Else
        EncryptPattern = True    ' Good finish
    End If
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

EncryptPattern_Error:
    Select Case Err.Number
           Case 13   ' Fall thru (mismatch)
           Case Else ' Valid error
                ' update log file, if requested
                If mblnLogData Then
                    strRecord = vbTab & vbTab & "File:  " & strFileName & vbNewLine & _
                                vbTab & vbTab & vbTab & vbTab & _
                                CStr(Err.Number) & "  " & Err.Description
                    RaiseEvent UpdateLogData(strRecord)
                End If
            
                strRecord = "ERROR File:  " & strFileName & vbNewLine & _
                            CStr(Err.Number) & "  " & Err.Description
                ErrorMsg MODULE_NAME, ROUTINE_NAME, strRecord
    End Select
    
    gblnStopProcessing = True
    Resume EncryptPattern_CleanUp

End Function

' ***************************************************************************
' Routine:       ChangeAlgorithms
'
' Description:   Create random generated password array using ASCII decimal
'                values 32-255.  Random selection of password key length and
'                block size.
'
' Parameters:    abytPwd()    - byte array to hold new password
'                lngKeyLength - key length selection
'                lngBlockSize - block size selection (Rijndael only)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-May-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 11-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Consolidated cipher and hash selection to a single routine
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated algorithm selection process
' 02-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added three more encryption options
' ***************************************************************************
Private Sub ChangeAlgorithms(ByRef abytPwd() As Byte, _
                             ByRef lngKeyLength As Long, _
                             ByRef lngBlockSize As Long)
                             
    ' Called by EncryptPattern()
    
    Const ROUTINE_NAME As String = "ChangeAlgorithms"

    On Error GoTo ChangeAlgorithms_Error

    Dim lngIndex  As Long
    Dim lngPwdLen As Long
    Dim avntData  As Variant
    
    Const MIN As Long = 15   ' Min password length
    Const MAX As Long = 50   ' Max password length
                             '     Ex:  50 = Len("Create random generated password array using ASCII")
                             
    Erase abytPwd()   ' Always start with empty arrays
    avntData = Empty  ' Always start with empty variants
    lngKeyLength = 0
    lngBlockSize = 0
    
    ' Create a random generated password
    DoEvents
    With mobjRandom
        ' Rijndael notes:  Maximum password length 32 bytes (32 bytes = 256 bits \ 8 bits).
        '                  Password length will be adjusted internally based on key length.
        lngPwdLen = Val(Int(Rnd() * (MAX - MIN + 1)) + MIN)   ' Random generated password length
        abytPwd() = .BuildWithinRange(lngPwdLen, 32, 255)     ' Create random generated password
    End With
    
    Select Case mlngEncryptAlgo
    
           Case eAlgo_Rijndael
                ' Load array with possible password
                ' key lengths and block sizes
                avntData = Array(128, 160, 192, 224, 256)
                
                lngIndex = Val(Int(Rnd() * 5))       ' Create index pointer (0-4)
                lngKeyLength = avntData(lngIndex)    ' Capture pwd key length from array
                
                lngIndex = Val(Int(Rnd() * 5))       ' Create index pointer (0-4)
                lngBlockSize = avntData(lngIndex)    ' Capture block size from array
    
           Case eAlgo_Blowfish, eAlgo_Twofish
                ' Load array with possible password key lengths
                avntData = Array(32, 64, 96, 128, 160, 192, 224, _
                                 256, 288, 320, 352, 384, 416, 448)
                
                lngIndex = Val(Int(Rnd() * 14))      ' Create index pointer (0-13)
                lngKeyLength = avntData(lngIndex)    ' Capture pwd key length from array
           
           Case eAlgo_ArcFour
                ' Load array with possible password key lengths
                avntData = Array(128, 160, 192, 224, 256, _
                                 288, 320, 352, 384, 416, _
                                 448, 512, 576, 640, 704, _
                                 768, 832, 896, 960, 1024)
                
                lngIndex = Val(Int(Rnd() * 20))      ' Create index pointer (0-19)
                lngKeyLength = avntData(lngIndex)    ' Capture pwd key length from array
    End Select
    
ChangeAlgorithms_CleanUp:
    avntData = Empty  ' Always empty variants when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub
    
ChangeAlgorithms_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ChangeAlgorithms_CleanUp

End Sub

' **************************************************************************
' Routine:       CustomPattern
'
' Description:   Loops thru an open file and overwrites each byte of data
'                with a user defined pattern.  This process is performed
'                from 1 to 99 times as defined by the user.
'
'                If verification is requested then the final pass will
'                be compared with the data that was supposed to be written.
'
' Important:     This is good security for up thru "SECRET" documents only if
'                some thought has gone into the pattern creation.
'
' Disposal:      If the hard disk has had data classified greater than
'                "CONFIDENTIAL", then the disk should be replaced with a new
'                one.  Since the cost of a hard disk has dropped so
'                dramatically, this should not be a factor.  You should be
'                considering the question, "What is my information worth to
'                someone else?".
'
'                Steps to follow to dispose of the old hard drive:
'                1. Overwrite multiple times with random data (Min 5 times).
'                   I recommend the Dban web site and creating a bootable CD
'                   or USB device that will wipe every sector on a disk. This
'                   is freeware and several governments approve its use.
'                   http://www.dban.org/
'                2. Remove disk from the old desktop or laptop and record the
'                   manufacturer, model, serial number, date of destruction and
'                   name of individual performing this process.
'                3. Plate area should be drilled in several places using a
'                   1/2 inch drill bit.
'                4. Disintegrate, incinerate, pulverize, shred, or melt the
'                   hard drive.
'
'                All of the above should be witnessed by at least two
'                additional persons and documented.
'
'                Ref:  Guidelines for Media Sanitization (SP800-88 dtd Sep-2006)
'                      http://csrc.nist.gov/publications/nistpubs/800-88/NISTSP800-88_rev1.pdf
'
' Parameters:    strFilename - Name of the file to be processed
'
' Returns:       TRUE - Successful completion
'                FALSE - something went wrong
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 14-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              The write process has been speeded up by 50% or more by
'              adjusting the record length based on amount of data left to
'              write.  This process calls the routine GetBlockSize().
' 14-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in designating data patterns.  Simplified logic
'              for when to create random data.
' ***************************************************************************
Private Function CustomPattern(ByVal strFileName As String) As Boolean

    Dim hFile         As Long       ' file handle
    Dim lngIdx        As Long       ' loop counter
    Dim lngIndex      As Long       ' loop counter
    Dim lngPause      As Long
    Dim lngPointer    As Long
    Dim lngIdxCnt     As Long       ' Element counter
    Dim lngFilePos    As Long       ' Starting position in the file
    Dim lngWriteCnt   As Long       ' Number of overwrites to file
    Dim lngRecordSize As Long       ' size of output record
    Dim lngElements   As Long
    Dim lngMaxDataAmt As Long       ' Max amount of random data to create
    Dim lngNoOfWrites As Long
    Dim blnRandom     As Boolean
    Dim blnLastPass   As Boolean    ' Flag designating last pass being performed
    Dim curFileSize   As Currency   ' original file size
    Dim curAmtLeft    As Currency   ' used to determine how much is left to write
    Dim curFilePos    As Currency   ' Starting position in the file
    Dim strRecord     As String     ' log file message
    Dim astrPattern() As String     ' Used for reporting purposes
    Dim astrData()    As String     ' Wipe pattern(s)
    Dim abytRnd()     As Byte       ' random data
    Dim abytTemp()    As Byte
    Dim abytData()    As Byte       ' Temp hold of random data
    Dim abytVerify()  As Byte       ' used to verify written data
    Dim abytRndFlag() As Byte       ' designates if random data for this element
    
    Const ROUTINE_NAME As String = "CustomPattern"

    On Error GoTo CustomPattern_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    Erase astrData()     ' Always start with empty arrays
    Erase astrPattern()
    Erase abytData()
    Erase abytRnd()
    Erase abytTemp()
    Erase abytVerify()
    Erase abytRndFlag()
    CustomPattern = False
    
    If Not CBool(IsArrayInitialized(mastrCustom)) Then
        InfoMsg "Cannot identify any custom patterns to wipe data." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If
    
    lngElements = UBound(mastrCustom)   ' capture number of elements in the array
    ReDim astrData(lngElements)         ' Size the custom pattern array
    ReDim astrPattern(lngElements)      ' for reporting purposes only
    ReDim abytRndFlag(lngElements)      ' Designates if random data for this element
    lngIdxCnt = 0

    lngWriteCnt = MAX_INT   ' Init file overwrite counter
    blnLastPass = False     ' Init last pass flag
    
    ' Reset file attributes
    SetFileAttributes strFileName, FILE_ATTRIBUTE_NORMAL
    
    mobjBigFiles.CalcFileSize strFileName, curFileSize   ' Get file size
    lngRecordSize = GetBlockSize(curFileSize)            ' Get maximum record size

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo CustomPattern_CleanUp
    End If

    If mblnMultiple Then
    
        ' multiple writes of the same value
        If Val(mastrCustom(0)) > 255 Then
            astrData(0) = " "
            astrPattern(0) = "Random"
            abytRndFlag(0) = 1      ' flag to create random data
            blnRandom = True
        Else
            astrData(0) = String$(lngRecordSize, Chr$(CInt(mastrCustom(lngIndex))))
            astrPattern(0) = "0x" & Right$("00" & Hex$(mastrCustom(lngIndex)), 2)
            abytRndFlag(lngIdxCnt) = 0      ' flag not to create random data
        End If

        lngNoOfWrites = Val(mastrCustom(1))
            
    Else
        ' load the custom patterns
        For lngIndex = 0 To lngElements
            
            If Len(Trim$(mastrCustom(lngIndex))) > 0 Then
                
                If Val(mastrCustom(lngIndex)) > 255 Then
                    astrData(lngIdxCnt) = " "
                    astrPattern(lngIdxCnt) = "Random"
                    abytRndFlag(lngIdxCnt) = 1      ' flag to create random data
                    blnRandom = True
                Else
                    astrData(lngIdxCnt) = String$(lngRecordSize, Chr$(CInt(mastrCustom(lngIndex))))
                    astrPattern(lngIdxCnt) = "0x" & Right$("0" & Hex$(mastrCustom(lngIndex)), 2)
                    abytRndFlag(lngIdxCnt) = 0      ' flag not to create random data
                End If
                
                lngIdxCnt = lngIdxCnt + 1
            
            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
        
        Next lngIndex
    
    End If

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo CustomPattern_CleanUp
    End If

    lngElements = lngIdxCnt
    ReDim Preserve astrData(lngElements)      ' Size the custom pattern array
    ReDim Preserve astrPattern(lngElements)   ' for reporting purposes only
    ReDim Preserve abytRndFlag(lngElements)

    With mobjRandom
        
        Rnd -1               ' Reset VB random number generator
        Randomize .RndSeed   ' Reseed VB random number generator
        
        For lngIndex = 1 To mlngPasses
    
            DoEvents
            RaiseEvent CountPasses(lngIndex, mlngPasses)
            RaiseEvent CurrentProgress(0, curFileSize)     ' maintain the ongoing totals
            CalculateProgress 0, curFileSize, False        ' individual progress
            DoEvents
    
            If lngIndex = mlngPasses Then
                blnLastPass = True
            End If
            
            lngFilePos = 1                             ' Set pointer to beginning of file
            curFilePos = 0                             ' Set pointer to beginning of file
            curAmtLeft = curFileSize                   ' Work with copy of file size
            lngRecordSize = GetBlockSize(curAmtLeft)   ' determine record size of output record
            lngMaxDataAmt = KB_4 + lngRecordSize       ' Max amount of random data to create
                    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            ' Open target file
            If Not mobjBigFiles.OpenReadWrite(strFileName, hFile) Then
                gblnStopProcessing = True
            End If
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            Do
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
                
                Erase abytTemp()   ' Verify temp array is empty
                
                If lngRecordSize > 1 Then
                          
                    ' See if random data is to be
                    ' used as one of the patterns
                    If blnRandom Then
                        
                        ' After 50 overwrites, generate new data
                        If lngWriteCnt >= 50 Then
                            lngWriteCnt = 0                                     ' reset overwrite counter
                            Erase abytRnd()                                     ' Always start with an empty array
                            abytRnd() = .BuildRndData(lngMaxDataAmt, , False)   ' Fill a byte array with random data
                        End If
                    End If
    
                    ' overwrite the same position x times
                    For lngIdx = 0 To lngElements - 1
    
                        If mblnMultiple Then
                            
                            For lngIdxCnt = 1 To lngNoOfWrites
    
                                ReDim abytTemp(lngRecordSize - 1)  ' Size temp array
                            
                                ' if this element is filled with random data,
                                ' then it is time to refresh the data
                                If CBool(abytRndFlag(lngIdx)) Then
                                    lngPointer = Int(Rnd() * KB_4)                              ' Create a starting point between 0-4095
                                    CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                                End If
    
                                ' Excessive number of DoEvents is to allow the read and
                                ' write process to stay in sync with the application speed.
                                DoEvents
                                If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                                    gblnStopProcessing = True
                                End If
               
                                ' An error occurred or user opted to STOP processing
                                DoEvents
                                If gblnStopProcessing Then
                                    Exit For    ' exit For..Next loop
                                End If
    
                                ' If verification has been requested then read the record
                                ' into a buffer and verify it's contents against what was
                                ' supposed to be written.
                                If mblnVerifyData Then
                                    If blnLastPass Then
                                        
                                        ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                            
                                        ' Capture the data that was just written to the file
                                        DoEvents
                                        If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                            gblnStopProcessing = True
                                            Exit For    ' exit For..Next loop
                                        End If
                                    
                                        If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                               strFileName, curFilePos, "Custom") Then
                                                               
                                            gblnStopProcessing = True
                                            Exit For    ' exit For..Next loop
                                        End If
                                    
                                    End If
                                End If
    
                                Erase abytTemp()   ' Empty arrays are empty
                                Erase abytVerify()
                
                                ' An error occurred or user opted to STOP processing
                                DoEvents
                                If gblnStopProcessing Then
                                    Exit For    ' exit For..Next loop
                                End If
    
                            Next lngIdxCnt
    
                        Else
                            
                            ReDim abytTemp(lngRecordSize - 1)  ' Size temp array
                    
                            ' if this element is filled with random data,
                            ' then it is time to refresh the data
                            If CBool(abytRndFlag(lngIdx)) Then
                                lngPointer = Int(Rnd() * KB_4) ' Create a starting point between 0-4095
                                CopyMemory abytTemp(0), abytRnd(lngPointer), lngRecordSize  ' Fill temp array
                            End If
    
                            ' Excessive number of DoEvents is to allow the read and
                            ' write process to stay in sync with the application speed.
                            DoEvents
                            If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytTemp()) Then
                                gblnStopProcessing = True
                            End If
           
                            ' An error occurred or user opted to STOP processing
                            DoEvents
                            If gblnStopProcessing Then
                                Exit For    ' exit For..Next loop
                            End If

                            ' If verification has been requested then read the record
                            ' into a buffer and verify it's contents against what was
                            ' supposed to be written.
                            If mblnVerifyData Then
                                If blnLastPass Then
                                        
                                    ReDim abytVerify(lngRecordSize - 1)  ' Size receiving buffer
                        
                                    ' Capture the data that was just written to the file
                                    DoEvents
                                    If Not mobjBigFiles.API_ReadFile(hFile, curFilePos, abytVerify()) Then
                                        gblnStopProcessing = True
                                        Exit For    ' exit For..Next loop
                                    End If
                                
                                    If Not VerifyLastWrite(abytTemp(), abytVerify(), _
                                                           strFileName, curFilePos, "Custom") Then
                                                           
                                        gblnStopProcessing = True
                                        Exit For    ' exit For..Next loop
                                    End If
                                    
                                End If
                            End If
                        End If
    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
    
                        Erase abytTemp()   ' Empty arrays are empty
                        Erase abytVerify()
                
                    Next lngIdx
    
                End If
    
                Erase abytTemp()   ' Empty arrays are empty
                Erase abytVerify()
                
                ' An error occurred or user opted to STOP processing or
                ' the last record was written
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
    
                DoEvents
                curFilePos = curFilePos + CCur(lngRecordSize)            ' Adjust pointers accordingly
                lngWriteCnt = lngWriteCnt + 1                            ' Update file overwrite counter
                curAmtLeft = curAmtLeft - lngRecordSize                  ' calc how much is left
                mcurCurrPerPass = mcurCurrPerPass + lngRecordSize        ' Current amt written this pass
                mcurOA_CurrentAmt = mcurOA_CurrentAmt + lngRecordSize    ' calc overall progress for pbar
                lngRecordSize = GetBlockSize(curAmtLeft)                 ' determine record size of output record
    
                DoEvents
                RaiseEvent CurrentProgress(curFilePos, curFileSize)      ' individual ongoing totals
                CalculateProgress curFilePos, curFileSize, False         ' individual progress
    
                RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxPerFile)  ' maintain overall totals
                CalculateProgress mcurOA_CurrentAmt, mcurOA_MaxAmt, True        ' overall progress
                DoEvents
    
                ' An error occurred or user opted to STOP processing or
                ' the last record was written
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If
    
                If curAmtLeft < 1 Then
                    Exit Do   ' exit Do..Loop
                End If
                    
            Loop
    
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
            ' Verify file handle is released
            mobjBigFiles.API_CloseFile hFile
    
            ' The pause is here to allow
            ' the write process to sync up
            ' with the speed of the CPU
            DoEvents
            lngPause = GetTickCount + 200
            Do While lngPause > GetTickCount
                DoEvents
            Loop
                    
            RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
            CalculateProgress 100, 100, False                      ' individual progress

            Erase abytTemp()     ' Empty temp byte array
            Erase abytVerify()
                                
            If lngIndex < mlngPasses Then
                mcurCurrPerPass = mcurCurrPerPass - curFileSize
            End If
    
        Next lngIndex
    End With
    
    CustomPattern = True
    DoEvents
    
    RaiseEvent CurrentProgress(curFileSize, curFileSize)   ' Show individual total
    CalculateProgress 100, 100, False                      ' individual progress

CustomPattern_CleanUp:
    ' Verify file handle is released
    mobjBigFiles.API_CloseFile hFile
    
    Erase astrData()     ' Always empty arrays when not needed
    Erase astrPattern()
    Erase abytData()
    Erase abytRnd()
    Erase abytTemp()
    Erase abytVerify()
    Erase abytRndFlag()
    
    ' Update log file, if requested
    If gblnStopProcessing Then
        If mblnLogData Then
            strRecord = "ABENDED:" & vbTab & "Custom file wipe on " & strFileName & vbNewLine & _
                        Space$(32) & ABEND_MSG
            RaiseEvent UpdateLogData(strRecord)
        End If
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CustomPattern_Error:
    ' Update log file, if requested
    If mblnLogData Then
        strRecord = vbTab & vbTab & "File:  " & strFileName & vbNewLine & _
                    vbTab & vbTab & vbTab & vbTab & _
                    CStr(Err.Number) & "  " & Err.Description
        RaiseEvent UpdateLogData(strRecord)
    End If

    strRecord = "ERROR File:  " & strFileName & vbNewLine & _
                CStr(Err.Number) & "  " & Err.Description
    ErrorMsg MODULE_NAME, ROUTINE_NAME, strRecord

    CustomPattern = False
    Resume CustomPattern_CleanUp

End Function

' **************************************************************************
' Routine:       SetPathFileDate
'
' Description:   Creates a unique date-time stamp to be used to reset the
'                folder or file that is being deleted.  The three properties
'                are:       Date created
'                           Date last accessed
'                           Date last modified
'
' Parameters:    strPathFileName - Name of folder or file to be updated
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Designate all three date properties to be updated.
'              Updated selection process for date and time.
' 28-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated reference to new routine name - .SetDateProperty()
' 26-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Joe for noticing the erroneous way of performing
'              the selection of date and time.  I was getting one less than
'              expected.  The necessary "Int(" part of the RND function
'              already drops all by one so I had to set the value to one
'              more than I wanted.
' 16-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated year selection process
' ***************************************************************************
Private Sub SetPathFileDate(ByVal strPathFileName As String)

    ' Called by WipeFile()
    '           WipeFolder()
    
    Dim datDate   As Date
    Dim intDay    As Integer
    Dim intYear   As Integer
    Dim intHour   As Integer
    Dim intMonth  As Integer
    Dim intMinute As Integer
    Dim intSecond As Integer
    
    strPathFileName = UnQualifyPath(strPathFileName)  ' Remove trailing backslash
    
    ' Random date creation
    intYear = CInt(Int(Rnd() * ((Year(Now) - 1) - 1990 + 1)) + 1990)   ' 1990 to (current year - 1)
    intMonth = CInt(Int(Rnd() * 12) + 1)   ' Months 1-12
    intDay = CInt(Int(Rnd() * 31) + 1)     ' Days 1-31  If number of days selected exceeds the number
                                           ' of days allowed for the selected month then DateSerial()
                                           ' will automatically adjust the date correctly for the
                                           ' following month.
                                           ' ex:  2/31/2012 will become 3/2/2012
    ' Random time creation
    intHour = CInt(Int(Rnd() * 24))        ' Hours 0-23  Twenty-four hour clock
    intMinute = CInt(Int(Rnd() * 60))      ' Minutes 0-59
    intSecond = CInt(Int(Rnd() * 60))      ' Seconds 0-59
    
    ' Concatenate to create new timestamp
    ' Ex:  5/9/2002 9:15:51 AM
    datDate = CDate(DateSerial(intYear, intMonth, intDay) & " " & _
                    TimeSerial(intHour, intMinute, intSecond))
    
    ' Update folder or file with new timestamp
    With mobjFileDate
        .PathFileName = strPathFileName    ' Name of folder or file to update
        .SelectedDateField = eThreeDates   ' Change all three date properties
        .TimeStamp = datDate               ' New timestamp to be used
        .SetDateProperty                   ' Perform update on folder or file
    End With
    DoEvents
    
End Sub

' **************************************************************************
' Routine:       CalculateProgress
'
' Description:   Calculates the current amount of completion
'
' Parameters:    dblCurrAmt - current value
'                dblMaxAmount - maximum value to be MAX_PERCENT%
'                blnOA_Progress [Optional]
'                    (Default) TRUE=Calculate the overall progress
'                    FALSE=Calculate the progress of the current item
'
' Returns:       Raises events denoting the percentage of progression,
'                Elapsed time, rate of data being written.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 15-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated time remaining calculations
' ***************************************************************************
Private Sub CalculateProgress(ByVal dblCurrAmt As Double, _
                              ByVal dblMaxAmount As Double, _
                     Optional ByVal blnOA_Progress As Boolean = True)

    Dim lngElapsed       As Long      ' total elapsed number of seconds
    Dim lngPercent       As Long      ' percentage to be calcuated
    Dim lngSecondsLeft   As Long      ' calculated seconds remaining
    Dim dblXferRate      As Double    ' calculated transfer rate
    Dim strElapsedTime   As String    ' formatted display of elasped time
    Dim strTransferRate  As String    ' formatted display of transfer rate per second
    Dim strTimeRemaining As String    ' formatted display of time remaining

    DoEvents
    If dblCurrAmt = 0@ And dblMaxAmount = 0@ Then
        
        ' Update progress bars
        If blnOA_Progress Then
            RaiseEvent TotalProgress(0)
        Else
            RaiseEvent FileProgress(0)
        End If
        
        strTimeRemaining = vbNullString
        strTransferRate = "0.0 Bytes"
        
        ' Update display status
        RaiseEvent ElapsedTime(vbNullString)
        RaiseEvent TimeRemaining(strTimeRemaining, strTransferRate)
        
        DoEvents
        Exit Sub
        
    End If

    ' Calculate percentage of progress
    lngPercent = CalcProgress(dblCurrAmt, dblMaxAmount)
    
    If blnOA_Progress Then   ' TRUE - Overall progress

        RaiseEvent TotalProgress(lngPercent)
        
        ' No more information needed if wiping
        ' free space with binary zeroes only
        If mblnWipeFreeSpace Then
           If Not mblnAlternateMethod Then
               Exit Sub
           End If
        End If
        
        ' update status display every 1/2 second
        If GetTickCount >= mlngPause Then

            mlngPause = GetTickCount + 500             ' set a 1/2 second delay
            lngElapsed = GetTickCount - mlngStartTime  ' Calculate seconds elapsed
        
            If lngElapsed > 0 Then
        
                strElapsedTime = FormatTimeDisplay(lngElapsed)  ' Calculate elapsed time
                RaiseEvent ElapsedTime(strElapsedTime)          ' Display elapsed time
            
                ' Calculate transfer rate
                dblXferRate = (dblCurrAmt / (lngElapsed / 1000))
                strTransferRate = DisplayNumber(dblXferRate) & " per sec"
                    
                ' Calculate time remaining
                If lngPercent <= 0 Then
                    lngSecondsLeft = (((100 / 1) * 2) * lngElapsed) - (((100 / 100) * 2) * lngElapsed)
                Else
                    lngSecondsLeft = (((100 / lngPercent) * 2) * lngElapsed) - (((100 / 100) * 2) * lngElapsed)
                End If
            
                strTimeRemaining = FormatTimeDisplay(lngSecondsLeft)         ' Format time remaining
                RaiseEvent TimeRemaining(strTimeRemaining, strTransferRate)  ' Update status
                
            End If
        End If
    Else
        RaiseEvent FileProgress(lngPercent)  ' FALSE - Individual progress
    End If

End Sub

' **************************************************************************
' Routine:       VerifyLastWrite
'
' Description:   Verify data that was written to a file versus the data that
'                was supposed to be written.
'
' Parameters:    abytSentToFile - Data that was sent to the file
'                abytReadFromFile - Data that was written to the file
'                strFilename - name of the file
'                curFilePos - pointer to starting position within the file
'                strPattern - Type of pattern that was tested
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-DEC-2001  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 10-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated to use byte arrays for data comparison
' ***************************************************************************
Private Function VerifyLastWrite(ByRef abytSentToFile() As Byte, _
                                 ByRef abytReadFromFile() As Byte, _
                                 ByVal strFileName As String, _
                                 ByVal curFilePos As Currency, _
                                 ByVal strPattern As String) As Boolean
    
    ' Called by MiscPatterns()
    '           DoDPatterns()
    '           EuroPatterns()
    '           GutmannPattern()
    '           CustomPattern()

    Dim lngSentLength As Long
    Dim lngReadLength As Long
    Dim lngPosition   As Long
    Dim strRecord     As String
    Dim blnBadData    As Boolean
    
    Const ROUTINE_NAME As String = "VerifyLastWrite"

    blnBadData = False    ' Assume good data
    
    ' Capture array lengths
    lngSentLength = UBound(abytSentToFile)
    lngReadLength = UBound(abytReadFromFile)
    
    ' See if data arrays are same length
    If lngSentLength = lngReadLength Then
           
        ' Now compare the data
        For lngPosition = 0 To lngSentLength
            
            ' Compare each byte.  They should be the same.
            If abytSentToFile(lngPosition) <> abytReadFromFile(lngPosition) Then
                blnBadData = True
                Exit For    ' exit For..Next loop
            End If
            
        Next lngPosition
           
    Else
        ' Data array lengths are different
        blnBadData = True
    End If

    ' If there is a bad data comparison
    If blnBadData Then
                
        ' Update exact position in file
        curFilePos = curFilePos + lngPosition + 1
        
        ' Write this message to a log file
        If mblnLogData Then
            
            ' Format log file message
            strRecord = "VERIFY FAILED:  " & strFileName & vbNewLine & _
                        Space$(35) & "Position:  " & Format$(curFilePos, "#,##0") & _
                        Space$(3) & "Pattern:  " & strPattern & vbNewLine & vbNewLine & _
                        "Sent to file:  " & mobjRandom.ByteArrayToString(abytSentToFile) & _
                        vbNewLine & vbNewLine & _
                        "Read from file:  " & mobjRandom.ByteArrayToString(abytReadFromFile)
            RaiseEvent UpdateLogData(strRecord)
            strRecord = vbNullString
            
        End If
         
        ' See if user wants to see any error messages
        If mblnDisplayMsgs Then
            
            ' If there is not a match then display a message
            strRecord = "VERIFY FAILED:  " & strFileName & vbNewLine & vbNewLine & _
                        "Position:  " & Format$(curFilePos, "#,##0") & vbNewLine & _
                        "Pattern:  " & strPattern
            
            InfoMsg strRecord & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        End If
    End If
    
    VerifyLastWrite = Not blnBadData  ' Opposite of current status
    
End Function

' **************************************************************************
' Routine:       SortFolderList
'
' Description:   Sort the collection of folders in ascending order.
'                    1.  transfer data from collection to an array
'                    2.  sort array
'                    3.  empty collection
'                    4.  copy sorted data back to collection
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-DEC-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Sub SortFolderList()

    ' Called by ProcessFolders()
    
    Dim lngIndex   As Long
    Dim lngCount   As Long
    Dim astrData() As String

    Erase astrData()  ' Always start with an empty array

    lngCount = mcolFolders.Count  ' # of items in collection
    ReDim astrData(lngCount - 1)  ' resize temp array

    ' transfer data from collection to temp array
    For lngIndex = 0 To lngCount - 1
        astrData(lngIndex) = mcolFolders.Item(lngIndex + 1)
    Next lngIndex

    ' sort temp array in ascending order
    If mobjRandom.CombSort(astrData()) Then

        ' Successful sort
        EmptyCollection mcolFolders        ' empty folder collection
        Set mcolFolders = New Collection   ' Instantiate new collection

        ' load collection with sorted array data
        For lngIndex = 0 To lngCount - 1
            mcolFolders.Add astrData(lngIndex)
        Next lngIndex
    
    End If
    
    Erase astrData()  ' Always empty arrays when not needed

End Sub

Private Sub ResetCounters()

    mcurFileCount = 0@
    mcurByteCount = 0@
    mcurFolderCount = 0@
    mcurCurrPerPass = 0@
    mcurOA_CurrentAmt = 0@
    mlngStartTime = GetTickCount()
    mlngPause = GetTickCount()

End Sub

Private Sub ResetVariables()
    
    mcurFilesize = 0@
    mcurBlockSize = 0@
    mcurFileCount = 0@
    mcurByteCount = 0@
    mcurFolderCount = 0@
    
    mcurOA_MaxAmt = 0@
    mcurOA_CurrentAmt = 0@
    mcurOA_MaxPerFile = 0@
    
    If mobjRandom Is Nothing Then
        Exit Sub
    End If
    
    mobjRandom.EmptyCollection mcolFolders  ' Verify collection object is empty
    
End Sub

Private Sub ResetFlags()
   
    mblnLogData = False           ' flag denotes whether to update a log file
    mblnDoFolders = False         ' flag denotes process a folder
    mblnUSB_Drive = False         ' flag denotes a USB device (Greater than 3mb)
    mblnFlashDrive = False        ' flag denotes a Flash drive (32gb or less)
    mblnVerifyData = False        ' flag denotes to verify the last write to file
    mblnProcessUSB = False        ' Used for update a log file with the USB process
    mblnDisplayMsgs = False       ' flag denotes to display verify messages
    mblnDoSubFolders = False      ' flag denotes process subfolders
    mblnZeroLastWrite = False     ' flag denotes using binary zeroes on the last pass
    mblnWipeFreeSpace = False     ' flag denotes wipe free space only
    mblnRemoveTopFolder = False   ' flag denotes wipe top level folder also
    mblnAlternateMethod = False   ' flag denotes using DoD Short to wipe free space
    mblnLogEncryptParms = False   ' Flag denotes logging AES encrypt parameters
   
End Sub

Private Sub FreeObjects()

    ' Free objects from memory
    Set mobjRandom = Nothing
    Set mobjBigFiles = Nothing
    Set mobjFileDate = Nothing
    Set mobjRijndael = Nothing
    Set mobjBlowfish = Nothing
    Set mobjArcFour = Nothing
    Set mobjTwofish = Nothing
    
End Sub

Private Sub Class_Initialize()
        
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objWipe = new cWipe
            
    FreeObjects   ' Verify class objects are empty
    
    ' Instantiate class objects
    Set mobjRandom = New cPrng
    Set mobjBigFiles = New cBigFiles
    Set mobjFileDate = New cFileDate

    ResetVariables   ' Initialize variables
    ResetFlags       ' Set module level flags to FALSE
    GetMemorySize    ' Get size of physical memory for this PC
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objWipe = Nothing
            
    ResetVariables   ' Empty variables
    ResetFlags       ' Set module level flags to FALSE
    FreeObjects      ' Verify class objects are empty
    
End Sub


' ***************************************************************************
' ****                     WithEvents                                    ****
' ***************************************************************************

' These routines were created thru the use of 'WithEvents'
' when declaring an object that has events within.  Parameter
' "lngProgress" is ignored.
Private Sub mobjArcFour_CipherProgress(ByVal lngProgress As Long)
    
    mcurByteCount = mcurByteCount + mcurBlockSize               ' Current amount written this pass
    mcurCurrPerPass = mcurCurrPerPass + mcurBlockSize           ' Accured amount written
    
    RaiseEvent CurrentProgress(mcurByteCount, mcurFilesize)     ' individual ongoing totals
    RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxAmt)  ' maintain overall totals
    
    CalculateProgress mcurByteCount, mcurFilesize, False        ' Individual progress
    CalculateProgress mcurCurrPerPass, mcurOA_MaxAmt, True      ' overall progress
    
End Sub

Private Sub mobjArcFour_CurrentBlockSize(ByVal lngBlockSize As Long)
    mcurBlockSize = CCur(lngBlockSize)   ' Data retrieved prior to cipher progress
End Sub

Private Sub mobjRijndael_CipherProgress(ByVal lngProgress As Long)
    
    mcurByteCount = mcurByteCount + mcurBlockSize               ' Current amount written this pass
    mcurCurrPerPass = mcurCurrPerPass + mcurBlockSize           ' Accured amount written
    
    RaiseEvent CurrentProgress(mcurByteCount, mcurFilesize)     ' individual ongoing totals
    RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxAmt)  ' maintain overall totals
    
    CalculateProgress mcurByteCount, mcurFilesize, False        ' Individual progress
    CalculateProgress mcurCurrPerPass, mcurOA_MaxAmt, True      ' overall progress
    
End Sub

Private Sub mobjRijndael_CurrentBlockSize(ByVal lngBlockSize As Long)
    mcurBlockSize = CCur(lngBlockSize)   ' Data retrieved prior to cipher progress
End Sub

Private Sub mobjBlowfish_CipherProgress(ByVal lngProgress As Long)
    
    mcurByteCount = mcurByteCount + mcurBlockSize               ' Current amount written this pass
    mcurCurrPerPass = mcurCurrPerPass + mcurBlockSize           ' Accured amount written
    
    RaiseEvent CurrentProgress(mcurByteCount, mcurFilesize)     ' individual ongoing totals
    RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxAmt)  ' maintain overall totals
    
    CalculateProgress mcurByteCount, mcurFilesize, False        ' Individual progress
    CalculateProgress mcurCurrPerPass, mcurOA_MaxAmt, True      ' overall progress
    
End Sub

Private Sub mobjBlowfish_CurrentBlockSize(ByVal lngBlockSize As Long)
    mcurBlockSize = CCur(lngBlockSize)   ' Data retrieved prior to cipher progress
End Sub

Private Sub mobjTwofish_CipherProgress(ByVal lngProgress As Long)
    
    mcurByteCount = mcurByteCount + mcurBlockSize               ' Current amount written this pass
    mcurCurrPerPass = mcurCurrPerPass + mcurBlockSize           ' Accured amount written
    
    RaiseEvent CurrentProgress(mcurByteCount, mcurFilesize)     ' individual ongoing totals
    RaiseEvent OverallProgress(mcurCurrPerPass, mcurOA_MaxAmt)  ' maintain overall totals
    
    CalculateProgress mcurByteCount, mcurFilesize, False        ' Individual progress
    CalculateProgress mcurCurrPerPass, mcurOA_MaxAmt, True      ' overall progress
    
End Sub

Private Sub mobjTwofish_CurrentBlockSize(ByVal lngBlockSize As Long)
    mcurBlockSize = CCur(lngBlockSize)   ' Data retrieved prior to cipher progress
End Sub
