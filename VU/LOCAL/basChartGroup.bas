Attribute VB_Name = "basChartGroup"
Option Explicit
'Version 1.0c
'Date: 09/06/01

'Ver    Date        By      Bug #       Notes
'----   --------    ---     --------    ---------------------------------------------------
'3.5    02/07/03    JPB                 Using Tracker database "iFIX 3.5"
'       02/07/03    JPB     T463        In CGW_ApplyFileToChart, don't append path to error string
'                                       if FilePath already has a path
'4.0    05/11/05    MDK     T1436       integrated Siebel 1-32281101
'    >> 11/12/04    hj      C1-32281101 Modified CGW_ApplyChartGroupSettings to set a pen's TimeZoneBiasRelative and
'                                       TimeZoneBiasExplicit property with Long type of data instead of Integer type.
'       09/09/2005  PBH                 PCM Changes:
'                                       Added Hook, UnHook, and WindowProc for if we decided we need to send a message to the
'                                       picture asking it for a file name and then let the picture send a different message back to us
'                                       telling us the file name.  This code is commented out, but left in-line.  Currently we send a message with a reference to a string and the picture
'                                       fills in the string.
'                                       The WM_COPYDATA solution was abandoned because the MFC documentation indicates that the data
'                                       passed in should be considered read only.
'       10/12/2005  kei     T2517       Changed CGW_SaveChartGroupFile to return Boolean,
'                                       and not to call frmChartGroupFileManagementForm.Show within this function
'       10/14/2005  ab      T1327       Ported SIM 272466
'       10/14/2005  ab      T1538       Ported:C288261
'>>     05/27/04    hj      C288261     Modified CGW_ApplyChartGroupSettings to set FetchPenLimits before setting
'                                       source to prevent EGU limits from getting into one shot bin and causing
'                                       race condition between the data system and CGW when FetchPenLimits is false.
'       10/14/2005  rp      T2031       Added 2 new methods CGW_HasEscKeySent, CGW_EscKeyHasBeenSent and a flag bHasSentEsc
'4.5    03/17/2007  ab      T4057       Windows Vista does not support SendKeys
'5.0    05/01/2008  mr                  Ported from 4.0 thc1-229269113 In CGW_ApplyChartGroupSettings, modify the way StartTime and TimeBeforeNow are set
'                                       1-262472724
'5.0    05/01/2008  mr                  Ported from 4.0 thc1-208005699 In CGW_ApplyChartGroupSetting, check FetchPenLimits before using limits from the .CSV file
'5.0    05/01/2008  mr                  Ported from 4.0 hj1-414967375 Ported SIM 284209 from 3.5 with some additional changes
'5.0    05/01/2008  mr                  Ported from 4.0 hj C284209 Modified CGW_ApplyChartGroupSettings and CGW_OpenChartGroupFileSettings
'5.0    05/01/2008  mr                  Ported from 4.0 hj1-439272898 Modified CGW_ApplyChartGroupSettings to not format the fixed date and time when
'                                       applying the settings in order to avoid seconds being stripped from the start time.
'5.0    12/16/2011  siva   1-1313672732 Trend Chart - display pen # when applying pens via .CSV file
'5.9    03/22/2017  raj    DE24889      Ported from 5.1:
'                                       MTK  1-3487849151 Trend doesn't load pens
Private Const KEYEVENTF_KEYUP = &H2
Private Const INPUT_KEYBOARD = 1

Private Type KEYBDINPUT
wVk As Integer
wScan As Integer
dwFlags As Long
time As Long
dwExtraInfo As Long
End Type

Private Type GENERALINPUT
dwType As Long
xi(0 To 23) As Byte
End Type

Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'htc_path
Global Const ghtc_path = 10

'Historical Sample Types
Global Const gintSample = 0
Global Const gintHigh = 1
Global Const gintLow = 2
Global Const gintAvg = 3

'Marker Character
Global Const gintNoMarker = 0
Global Const gintRectangleMarker = 1
Global Const gintOvalMarker = 2
Global Const gintDiamondMarker = 3
Global Const gintCharacterMarker = 4

'Pen Line Style
Global Const gintEdgeStyleSolid = 0
Global Const gintEdgeStyleDash = 1
Global Const gintEdgeStyleDot = 2
Global Const gintEdgeStyleDashDot = 3
Global Const gintEdgeStyleDashDotDot = 4
Global Const gintEdgeStyleNone = 5
Global Const gintEdgeStyleInsideFrame = 6

'StartDateType and StartTimeType
Global Const gintRelative = 0
Global Const gintFixed = 1

'Daylight Savings Time etc
Global Const gintInterval = 10000
Global Const gintDisplayMS = 0
Global Const gintDaylightSavingsTime = 1
Global Const gintTimeZone = 0

Global Const g_NT_SECURITY_ERROR = 75

' Structure to hold the pen properties
Public Type gPenProperties
  ApplyToAll As Boolean
  DaysBeforeNow As Integer
  Duration As Long
  FetchPenLimits As Boolean
  FixedDate As Date
  FixedTime As Date
  HiLimit As Double
  HistoricalSampleType As Integer
  LoLimit As Double
  MarkerChar As String
  MarkerStyle As Integer
  PenLineColor As Long
  PenLineWidth As Long
  PenLineStyle As Integer
  Source As String
  TimeBeforeNow As Long
  StartDateType As Integer
  StartTimeType As Integer
  Interval As Long
  DisplayMS As Boolean
  DaylightSavingsTime As Boolean
  TimeZone As Integer
End Type

'Array containing all the pen properites of the chart or the file
Public guPenPropertiesArray() As gPenProperties
'Global Error for bad file name
Public gblnBadFileName As Boolean
'Set to true if the group was saved
Global gblnSaved As Boolean
'Flag which determines if any changes were made to the file
Public gblnFormChangedFlag As Boolean
'If from Save Button(3), Save As button(2), or Select Button(1), initialized to (0)
Public gintSaveOkSaveAndApply As Integer
'The file name
Public gFileName As String
Public gFullPath As String
Public gblnNoMarkerChar As Boolean

'Collection of Colors
Public gcolColorCollection As New Collection

'Some Global Booleans needed for within the Project
Public gblnSaveComingFromCGF As Boolean
Public gblnUserClickedCancelInFileMgmtForm As Boolean
Public gblnCameFromApplyButton As Boolean
Public gblnANewFile As Boolean
Public gobjSelectedChart As Object
Public gblnApplyFileToChart As Boolean
Public gblnScriptAuthor As Boolean
Public gblnEditPens As Boolean

'System default date
Public gdtSystemDefaultDate As Date 'mr050108 Ported from 4.0 - hj022304 The system default date is 30 December 1899, midnight

'CSV file version
Public giVer As Integer

'NLS Object
Dim objStrMgr As Object

'AppObj Object
'jes012804 Case 283023 must be global
'Dim AppObj As Object 'hj072903
Global AppObj As Object

'NLS Strings
Global Const NLS_FONTNAME = 1000 'kei080102 Tracker #4057
Global Const NLS_FONTSIZE = 1001 'kei080102 Tracker #4057

Const NLS_TITLE = 5500
Const NLS_DURATIONERROR = 5510
Const NLS_OVERWRITE = 5520
Const NLS_EXC = 5530
Const NLS_APPLYALL = 5532
Const NLS_DAYSBEFORENOW = 5534
Const NLS_DURATION = 5536
Const NLS_FETCHPENLIMITS = 5538
Const NLS_FIXEDDATE = 5540
Const NLS_FIXEDTIME = 5542
Const NLS_HILIMIT = 5544
Const NLS_HISTORICALSAMPLETYPE = 5546
Const NLS_LOLIMIT = 5548
Const NLS_MARKERCHAR = 5550
Const NLS_MARKERSTYLE = 5552
Const NLS_PENLINECOLOR = 5554
Const NLS_PENLINESTYLE = 5556
Const NLS_PENLINEWIDTH = 5558
Const NLS_SOURCE = 5560
Const NLS_STARTDATETYPE = 5562
Const NLS_STARTTIMETYPE = 5564
Const NLS_TIMEBEFORENOW = 5566
Const NLS_WRONGFORMAT1 = 5570
Const NLS_WRONGFORMAT2 = 5572
Const NLS_WRONGFORMAT3 = 5574
Const NLS_WRONGFORMAT4 = 5576
Const NLS_WRONGFORMAT5 = 5578
Const NLS_WRONGFORMAT6 = 5580
Const NLS_CHARTDNE = 5582
Const NLS_BADFORMAT1 = 5584
Const NLS_BADFORMAT2 = 5586
Const NLS_NOCHART = 5588
Const NLS_INVALIDPEN1 = 5590
Const NLS_INVALIDPEN2 = 5592
Const NLS_INVALIDPEN3 = 5594
Const NLS_SAVEERROR1 = 5595
Const NLS_SAVEERROR2 = 5596
Const NLS_SAVEERROR3 = 5597
Const NLS_CGFILEDNE1 = 6840
Const NLS_CGFILEDNE2 = 6850
Const NLS_CGFILEDNE3 = 6852
Const NLS_CGFILEDNE4 = 6854
Const NLS_CGFILEDNE5 = 6856
Const NLS_CGFILEDNE6 = 6858
Const NLS_NOCHARTSINPIC = 6860
Const NLS_TBS = 6980
Const NLS_CHARTDNE1 = 7050
Const NLS_CHARTDNE2 = 7060
Const NLS_DST_HEAD = 7092
Const NLS_TIMEZONE_HEAD = 7168
Const NLS_INTERVAL = 7172
Const NLS_INTERVALERROR = 7174
Const NLS_INTERVALERROR2 = 7180
Const NLS_DISPLAYMS = 7176

' this is the section and entry name in the INI
Private Const HISTORIAN_INI_SECTION = "Historian"
Private Const HISTORIAN_INI_ENTRY = "CurrentHistorian"
' these are the possible values for current historian
Private Const IHISTORIAN_INI_VAL = "iHistorian"
Private Const CLASSIC_INI_VAL = "Classic"

'Private Const GWL_WNDPROC = (-4)
'Private lpPrevWndProc As Long
'Private gHW As Long

'rp101405 T2031
Private bHasSentEsc As Boolean

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub CGW_GenerateColors(NumColors)
'Generates a default gcolColorCollection
    Dim intForLoop As Integer
    Dim intNextColor As Integer
    Dim intRed As Integer
    Dim intGreen As Integer
    Dim intBlue As Integer
    Dim intRedIn As Integer
    Dim intGreenIn As Integer
    Dim intBlueIn As Integer
    Dim intColorIncrement As Integer
    
    On Error GoTo ErrorHandler
    
    'Remove items from the colorcollection
    For intForLoop = 1 To gcolColorCollection.Count
        gcolColorCollection.Remove 1
    Next
    intColorIncrement = 1
    If (NumColors > 0 And (NumColors < 255 Or NumColors = 255)) Then
        intColorIncrement = 255 / NumColors
    End If
    intRed = 0
    intGreen = 127
    intBlue = 0
    'intNextColor is the counter for the following color algorithim.  When intNextColor is
    'incremented by 1, the base color changes.  For example, when intNextColor is 1, it is a shade
    'of red.  When intNextColor is 2, it is a shade of green.
    intNextColor = 0
    For intForLoop = 0 To NumColors
        intNextColor = intNextColor + 1
        'Choose a blue shade
        If intNextColor = 1 Then
            intBlueIn = 255 - intBlue
            gcolColorCollection.Add RGB(0, 0, intBlueIn)
        'Choose a green shade
        ElseIf intNextColor = 2 Then
        'intGreen is the system default of green
            intGreenIn = 255 - intGreen
            gcolColorCollection.Add RGB(0, intGreenIn, 0)
        'Choose a red shade
        ElseIf intNextColor = 3 Then
            intRedIn = 255 - intRed
            gcolColorCollection.Add RGB(intRedIn, 0, 0) 'QBColor(12)
        'Choose a Cyan shade
        ElseIf intNextColor = 4 Then
            intBlueIn = 255 - intGreen
            intGreenIn = 255 - intGreen
            gcolColorCollection.Add RGB(0, intGreenIn, intBlueIn) 'QBColor(3)
            intGreen = 0
        'Choose a Magenta shade
        ElseIf intNextColor = 5 Then
            intBlueIn = 255 - intBlue
            intRedIn = 255 - intRed
            gcolColorCollection.Add RGB(intRedIn, 0, intBlueIn) 'QBColor(13)
        'Choose a Yellow shade
        ElseIf intNextColor = 6 Then
            intRedIn = 255 - intRed
            intGreenIn = 255 - intGreen
            gcolColorCollection.Add RGB(intRedIn, intGreenIn, 0) 'QBColor(14)
        'Choose a gray shade
        ElseIf intNextColor = 7 Then
            intRedIn = 255 - intRed
            intBlueIn = 255 - intBlue
            intGreenIn = 255 - intGreen
            gcolColorCollection.Add RGB(intRedIn, intGreenIn, intBlueIn)
            intGreen = 127
        'Choose any color
        Else
            intBlueIn = 255 - intBlue
            intGreenIn = 255 - intGreen
            intRedIn = 255 - intRed
            gcolColorCollection.Add RGB(intRedIn, intGreenIn, intBlueIn)
            intBlue = intBlue + intColorIncrement
            intGreen = intGreen + intColorIncrement
            intRed = intRed + intColorIncrement
            intNextColor = 0
        End If
     Next
     Exit Sub
    
ErrorHandler:
    HandleError
End Sub
Public Function GetChartGroupFileName() As String
    Dim intLenOfHTRDir As Integer
    Dim strHTRPath As String
    Dim strFileNameWOHTRPath As String
    Dim intLenOfgFullPath As Integer
    
    On Error GoTo ErrorHandler
    gintSaveOkSaveAndApply = 8
    frmChartGroupFileManagementForm.Show
    
    strHTRPath = System.FixPath(ghtc_path)
    intLenOfHTRDir = Len(strHTRPath)
    intLenOfgFullPath = Len(gFullPath)
    'just return the filename, and any other folder it involves other than
    'the default HTR path
    If UCase(Left(gFullPath, intLenOfHTRDir)) = UCase(strHTRPath) Then
        'add one for the slash before the filename
        intLenOfHTRDir = intLenOfHTRDir + 1
        strFileNameWOHTRPath = Right(gFullPath, intLenOfgFullPath - intLenOfHTRDir)
    Else
        strFileNameWOHTRPath = gFileName
    End If
    
    GetChartGroupFileName = strFileNameWOHTRPath
    ' Unhook - PBH - See Unhook() function for comments
    Unload frmChartGroupFileManagementForm
    Exit Function
    
ErrorHandler:
    HandleError
End Function
Public Sub CGW_ApplyFileToChart(Optional strFileName As String, Optional strChartName As String)

    Dim objChart As Object
    Dim objVar As Object
    Dim blnVarValueSet As Boolean
    Dim objCheckForVar As Object
    Dim strTitle As String
    Dim strError As String
    Dim MyFile As String
    Dim mstrRootDir As String
    
    On Error GoTo ErrorHandler
        ' hj072903
    ' Is this script running in the workspace or background?
    If TypeName(Application) = "CFixApp" Then
        ' running in the workspace
        Set AppObj = Application
    Else
        ' running in the background
            
        ' see if we can get the workspace object
        Set AppObj = GetObject(, "Workspace.Application")
        
        If AppObj Is Nothing Then
            Exit Sub
        End If
    End If
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    gblnApplyFileToChart = True
    blnVarValueSet = False
    mstrRootDir = System.FixPath(ghtc_path)
    
    'if user did not enter a Chart Name by either putting a blank string or leaving the space empty,
    'show the Chart List form to allow them to choose a chart
    'Check to see what Chart Name was entered
    If strChartName = "" Or IsMissing(strChartName) Then
        If strFileName = "" Then
            frmChartList.txtFileName.Text = objStrMgr.GetNLSStr(CLng(NLS_TBS))
        Else
            frmChartList.txtFileName.Text = strFileName
        End If
        frmChartList.Show
        'if there are no charts in the picture
        If frmChartList.lstChartList.ListCount = 0 Then
           'Error message already was shown
            GoTo ExitThisSub
        End If
        strChartName = frmChartList.lstChartList.Text
        If strChartName <> "" Then
            Unload frmChartList
            Set objChart = AppObj.ActiveDocument.Page.FindObject(strChartName) 'hj072903
            Set gobjSelectedChart = objChart
        Else
            'User clicked Cancel
            GoTo ExitThisSub
        End If
    Else
        Set objChart = AppObj.ActiveDocument.Page.FindObject(strChartName) 'hj072903
        Set gobjSelectedChart = objChart
    End If
    
    'Check to see what File Name was entered
    If strFileName = "" Then
        frmChartGroupFileManagementForm.Show
        'loop thru the contained objects for a var, if one DNE, create it.
        strFileName = gFullPath
        'assignVar and exit this sub
        'See if user clicked cancel in FileManagementform, if yes, exitthissub
        If gblnUserClickedCancelInFileMgmtForm Then
            GoTo ExitThisSub
        Else
        'Hit OK, need to set the InitValue of the Var, and Build one if necc.
            GoTo SetVar
        End If
    ElseIf IsMissing(strFileName) = True Then
    'no file name was given, set it equal to gFullPath
        strFileName = gFullPath
    ElseIf IsMissing(strFileName) = False Then
        'Check to see if file's extension is supplied.
        If InStr(1, strFileName, ".", vbTextCompare) = 0 Then
            'add on ".csv" to the file name
            strFileName = strFileName & ".csv"
        'If the file extension is supplied, make sure it is ".csv".  If it is not, goto NoSuchPicture
        ElseIf InStr(1, strFileName, "csv", vbTextCompare) = 0 Then
          GoTo NoSuchPicture
        End If
        'need to find the FQN of the CGFile
        'We give it the DDS and HTR path if it is not at the beginning
        'hj112001 Comparison here should be case insensitive
        'If InStr(1, strFileName, mstrRootDir) = 0 Then
        If InStr(1, strFileName, mstrRootDir, vbTextCompare) = 0 Then
            strFileName = mstrRootDir & "\" & strFileName
        End If
        MyFile = Dir(strFileName)
        If MyFile = "" Then
            GoTo NoSuchPicture
        End If
    End If
    'Will go here if the strFilename was given in the call
    'Changed to Instr comparison in case strChartName is the fully qualified name for the chart (SED 6/23/99)
    If Not (InStr(1, strChartName, gobjSelectedChart.Name, vbTextCompare) = 0) Then
        CGW_OpenChartGroupFile strFileName
        CGW_ApplyChartGroupSettings gobjSelectedChart
    End If
    GoTo SetVar
    
NoSuchPicture:
    'JPB020703 Tracker #463 if there is not a path in the file name then pre-pend one
    If InStr(1, strFileName, "\", vbTextCompare) = 0 Then
        strFileName = mstrRootDir & "\" & strFileName
    End If
    'strError = objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE1)) & System.FixPath(ghtc_path) & "\" & strFileName & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE2)) & vbCrLf & vbCrLf & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE3)) & vbCrLf & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE4)) & System.FixPath(ghtc_path) & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE5)) & vbCrLf & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE6))
    strError = objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE1)) & strFileName & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE2)) & vbCrLf & vbCrLf & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE3)) & vbCrLf & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE4)) & System.FixPath(ghtc_path) & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE5)) & vbCrLf & objStrMgr.GetNLSStr(CLng(NLS_CGFILEDNE6))
    'End JPB020703 Tracker #463
    strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
    MsgBox strError, vbOKOnly, strTitle
    GoTo ExitThisSub
    
SetVar:
    'assign the init value of the var to the strFileName
    'if a var DNE on the Chart, create One.
    For Each objCheckForVar In gobjSelectedChart.ContainedObjects
        If objCheckForVar.ClassName = "Variable" And Left(objCheckForVar.Name, 8) = "FileName" Then
            objCheckForVar.InitialValue = strFileName
            blnVarValueSet = True
        End If
    Next objCheckForVar
            
    If Not blnVarValueSet Then
        Set objVar = objChart.BuildObject("Variable")
        objVar.Name = "FileName"
        objVar.VariableType = vbString
        objVar.InitialValue = strFileName
        objVar.EnableAsVBAControl = False
    End If
    GoTo ExitThisSub
    
ErrorHandler:
    If Err.Number = -2147197772 Then
    'the chart does not exist in the currently active picture
        strError = objStrMgr.GetNLSStr(CLng(NLS_CHARTDNE1)) & strChartName & objStrMgr.GetNLSStr(CLng(NLS_CHARTDNE2))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
        MsgBox strError, vbOKOnly, strTitle
    Else
        HandleError
    End If
    
ExitThisSub:
    Set gobjSelectedChart = Nothing
    gblnApplyFileToChart = False
    gFileName = ""
    gFullPath = ""
    
End Sub

Public Sub CGW_ApplyFileNameToUserGlobal(strFileName As String)
   On Error GoTo ErrorHandler
   gFileName = strFileName
   Exit Sub
    
ErrorHandler:
    HandleError
End Sub

Public Function CGW_SetDuration(lngDuration As Long) As String
    Dim strDays As String
    Dim strHours As String
    Dim strMinutes As String
    Dim strSeconds As String
    Dim strTempMinutes As String
    Dim lngx As Long
    Dim lngy As Long
    
    On Error GoTo ErrorHandler
    lngy = 60
    lngx = 24

    strDays = (lngDuration - lngDuration Mod (lngx * lngy * lngy)) / (lngx * lngy * lngy)
    If Len(strDays) = 1 Then
      strDays = "0" + strDays
    End If
    strHours = ((lngDuration Mod (lngx * lngy * lngy)) - ((lngDuration Mod (lngx * lngy * lngy)) Mod (lngy * lngy))) / (lngy * lngy)
    If Len(strHours) = 1 Then
      strHours = "0" + strHours
    End If
    strTempMinutes = ((lngDuration Mod (lngx * lngy * lngy)) Mod (lngy * lngy))
    strMinutes = (strTempMinutes - (strTempMinutes Mod lngy)) / lngy
    If Len(strMinutes) = 1 Then
      strMinutes = "0" + strMinutes
    End If
    strSeconds = strTempMinutes Mod lngy
    If Len(strSeconds) = 1 Then
      strSeconds = "0" + strSeconds
    End If
    CGW_SetDuration = Trim(strDays) + ":" + Trim(strHours) + ":" + Trim(strMinutes) + ":" + Trim(strSeconds)
    Exit Function
    
ErrorHandler:
    HandleError
End Function

Public Function CGW_GetDuration(strDuration As String) As Long
    Dim strDays As String
    Dim strHours As String
    Dim strMinutes As String
    Dim strSeconds As String
    Dim strChecktxtDuration As String
    Dim lngHoursInaDay As Long
    Dim lngMinutesSecondsInaday As Long
    Dim strError As String
    Dim strTitle As String
    
    On Error GoTo ErrorHandler
    lngHoursInaDay = 24
    lngMinutesSecondsInaday = 60

    strDuration = Trim(strDuration)
    If Len(strDuration) <> 11 Then
      GoTo ErrorHandler
    End If

    strDays = Left(strDuration, 2)
    If (Not IsNumeric(strDays)) Or (CLng(strDays) < 0) Then
        GoTo ErrorHandler
    End If

    strChecktxtDuration = Left(Right(strDuration, Len(strDuration) - 2), 1)
    If strChecktxtDuration <> ":" Then
        GoTo ErrorHandler
    End If

    strHours = Left(Right(strDuration, Len(strDuration) - 3), 2)
    If (Not IsNumeric(strHours)) Or (CLng(strHours) >= 24) Or (CLng(strHours) < 0) Then
        GoTo ErrorHandler
    End If

    strChecktxtDuration = Left(Right(strDuration, Len(strDuration) - 5), 1)
    If strChecktxtDuration <> ":" Then
        GoTo ErrorHandler
    End If

    strMinutes = Left(Right(strDuration, Len(strDuration) - 6), 2)
    If (Not IsNumeric(strMinutes)) Or (CLng(strMinutes) >= 60) Or (CLng(strMinutes) < 0) Then
        GoTo ErrorHandler
    End If

    strChecktxtDuration = Left(Right(strDuration, Len(strDuration) - 8), 1)
    If strChecktxtDuration <> ":" Then
        GoTo ErrorHandler
    End If

    strSeconds = Left(Right(strDuration, Len(strDuration) - 9), 2)
    If (Not IsNumeric(strSeconds)) Or (CLng(strSeconds) >= 60) Or (CLng(strSeconds) < 0) Then
        GoTo ErrorHandler
    End If
    
    CGW_GetDuration = CLng(strDays) * lngHoursInaDay * lngMinutesSecondsInaday * lngMinutesSecondsInaday + CLng(strHours) * lngMinutesSecondsInaday * lngMinutesSecondsInaday + CLng(strMinutes) * lngMinutesSecondsInaday + CLng(strSeconds)
    
    Exit Function

ErrorHandler:
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    strError = objStrMgr.GetNLSStr(CLng(NLS_DURATIONERROR))
    strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
    MsgBox strError, , strTitle
    frmChartGroupForm.txtDuration = "00:00:01:00"
End Function
Public Function CGW_SetTimeBeforeNow(lngTimeBeforeNow As Long) As String
    Dim strDays As String
    Dim strHours As String
    Dim strMinutes As String
    Dim strSeconds As String
    Dim strTempMinutes As String
    Dim lngx As Long
    Dim lngy As Long
    
    On Error GoTo ErrorHandler
    lngy = 60
    lngx = 24
    
    strDays = (lngTimeBeforeNow - lngTimeBeforeNow Mod (lngx * lngy * lngy)) / (lngx * lngy * lngy)
    If Len(strDays) = 1 Then
      strDays = "0" + strDays
    End If
    strHours = ((lngTimeBeforeNow Mod (lngx * lngy * lngy)) - ((lngTimeBeforeNow Mod (lngx * lngy * lngy)) Mod (lngy * lngy))) / (lngy * lngy)
    If Len(strHours) = 1 Then
      strHours = "0" + strHours
    End If
    strTempMinutes = ((lngTimeBeforeNow Mod (lngx * lngy * lngy)) Mod (lngy * lngy))
    strMinutes = (strTempMinutes - (strTempMinutes Mod lngy)) / lngy
    If Len(strMinutes) = 1 Then
      strMinutes = "0" + strMinutes
    End If
    strSeconds = strTempMinutes Mod lngy
    If Len(strSeconds) = 1 Then
      strSeconds = "0" + strSeconds
    End If
    CGW_SetTimeBeforeNow = Trim(strHours) + ":" + Trim(strMinutes) + ":" + Trim(strSeconds)
    Exit Function
    
ErrorHandler:
    HandleError
    
End Function

Public Function CGW_GetTimeBeforeNow(strTimeBeforeNow As String) As Long
    Dim strDays As String
    Dim strHours As String
    Dim strMinutes As String
    Dim strSeconds As String
    Dim strChecktxtDuration As String
    Dim lngHoursInaDay As Long
    Dim lngMinutesSecondsInaday As Long
    Dim strError As String
    Dim strTitle As String
    
    On Error GoTo ErrorHandler
    lngHoursInaDay = 24
    lngMinutesSecondsInaday = 60
    
    strTimeBeforeNow = Trim(strTimeBeforeNow)
    If Len(strTimeBeforeNow) <> 8 Then
      GoTo ErrorHandler
    End If
    
    strHours = Left(strTimeBeforeNow, 2)
    If (Not IsNumeric(strHours)) Or (CLng(strHours) >= 24) Or (CLng(strHours) < 0) Then
        GoTo ErrorHandler
    End If
    
    strChecktxtDuration = Left(Right(strTimeBeforeNow, Len(strTimeBeforeNow) - 2), 1)
    If strChecktxtDuration <> ":" Then
        GoTo ErrorHandler
    End If
    
    strMinutes = Left(Right(strTimeBeforeNow, Len(strTimeBeforeNow) - 3), 2)
    If (Not IsNumeric(strMinutes)) Or (CLng(strMinutes) >= 60) Or (CLng(strMinutes) < 0) Then
        GoTo ErrorHandler
    End If
    
    strChecktxtDuration = Left(Right(strTimeBeforeNow, Len(strTimeBeforeNow) - 5), 1)
    If strChecktxtDuration <> ":" Then
        GoTo ErrorHandler
    End If
    
    strSeconds = Left(Right(strTimeBeforeNow, Len(strTimeBeforeNow) - 6), 2)
    If (Not IsNumeric(strSeconds)) Or (CLng(strSeconds) >= 60) Or (CLng(strSeconds) < 0) Then
        GoTo ErrorHandler
    End If
    
    CGW_GetTimeBeforeNow = CLng(strHours) * lngMinutesSecondsInaday * lngMinutesSecondsInaday + CLng(strMinutes) * lngMinutesSecondsInaday + CLng(strSeconds)
    
    Exit Function

ErrorHandler:
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    strError = objStrMgr.GetNLSStr(CLng(NLS_DURATIONERROR))
    strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
    MsgBox strError, , strTitle
    frmChartGroupForm.cboTimeBeforeNow = "000:02:00:00"
End Function
Sub CGW_InitializePenPropertiesArray(ArrayIndex As Integer, strSource As String)

  Dim dtDateHolder As Date
  Dim varResults As Variant
  Dim varAttributeNames As Variant
  Dim lngStatus As Long
  Dim strFullyQualifiedName As String
  
  On Error GoTo ErrorHandler
  strFullyQualifiedName = strSource
  guPenPropertiesArray(ArrayIndex).DaysBeforeNow = 0
  guPenPropertiesArray(ArrayIndex).Duration = 300
  guPenPropertiesArray(ArrayIndex).FetchPenLimits = True
  dtDateHolder = DateAdd("h", 0, Now)
  'jes clarify #235605
  'guPenPropertiesArray(ArrayIndex).FixedDate = Format(dtDateHolder, "mm/dd/yy")
  guPenPropertiesArray(ArrayIndex).FixedDate = Format(dtDateHolder, "Short Date")
  'guPenPropertiesArray(ArrayIndex).FixedTime = Format(dtDateHolder, "hh:mm:ss AM/PM")
  guPenPropertiesArray(ArrayIndex).FixedTime = Format(dtDateHolder, "Short Time")
  
  guPenPropertiesArray(ArrayIndex).HistoricalSampleType = gintSample
  
  guPenPropertiesArray(ArrayIndex).MarkerChar = ""
  guPenPropertiesArray(ArrayIndex).MarkerStyle = gintNoMarker
  CGW_GenerateColors (ArrayIndex)
  guPenPropertiesArray(ArrayIndex).PenLineColor = CLng(gcolColorCollection(ArrayIndex + 1))
  frmChartGroupForm.ctrlColorButton.Color = CLng(gcolColorCollection(ArrayIndex + 1))
  guPenPropertiesArray(ArrayIndex).PenLineWidth = 1
  guPenPropertiesArray(ArrayIndex).PenLineStyle = gintEdgeStyleSolid
  guPenPropertiesArray(ArrayIndex).TimeBeforeNow = 300
  guPenPropertiesArray(ArrayIndex).StartDateType = gintRelative
  guPenPropertiesArray(ArrayIndex).StartTimeType = gintRelative
  guPenPropertiesArray(ArrayIndex).Interval = gintInterval
  guPenPropertiesArray(ArrayIndex).DisplayMS = gintDisplayMS
  guPenPropertiesArray(ArrayIndex).DaylightSavingsTime = gintDaylightSavingsTime
  guPenPropertiesArray(ArrayIndex).TimeZone = gintTimeZone
  System.GetPropertyAttributes strFullyQualifiedName, 2, varResults, varAttributeNames, lngStatus
  
  If lngStatus = 0 Then
    guPenPropertiesArray(ArrayIndex).HiLimit = varResults(1)
    guPenPropertiesArray(ArrayIndex).LoLimit = CDbl(varResults(0))
    frmChartGroupForm.txtHiLimit = guPenPropertiesArray(ArrayIndex).HiLimit
    frmChartGroupForm.txtLowLimit = guPenPropertiesArray(ArrayIndex).LoLimit
  Else
    guPenPropertiesArray(ArrayIndex).HiLimit = 100
    guPenPropertiesArray(ArrayIndex).LoLimit = CDbl(0)
    frmChartGroupForm.txtHiLimit = guPenPropertiesArray(ArrayIndex).HiLimit
    frmChartGroupForm.txtLowLimit.Text = CStr(guPenPropertiesArray(ArrayIndex).LoLimit)
  End If
  Exit Sub
    
ErrorHandler:
    HandleError
  
End Sub


' ================================== PBH ================================
' PBH - Code saved in case we need to send a message to the picture and then let the picture send us back another message that tells
' us the filename.  Currently, we pass a reference to a string and the picture fills in the string.  The rest of the useful code is
' located in basChartGroup.bas and is Hook(), UnHook() and WindowProc()
'Public Sub Hook()
'lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
'Debug.Print lpPrevWndProc
'End Sub
'
'Public Sub Unhook()
'Dim temp As Long
'temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
'End Sub
'
'Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
'ByVal wParam As Long, ByVal lParam As Long) As Long
'If uMsg = 2839 Then
'Stop
'Call CGW_Do_Open_File(lParam)
'End If
'WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
'End Function
'==========================================================================




'**************************************************************************************
'
'   Subroutine:   OpenChartGroupForm()
'
'   Procedure:    This Procedure opens the chart group form in both configure mode
'                 run mode
'
'**************************************************************************************
Sub CGW_OpenChartGroupForm(Optional strChartName As String)
    
    On Error GoTo ErrorHandler
    ' hj072903
    ' Is this script running in the workspace or background?
    If TypeName(Application) = "CFixApp" Then
        ' running in the workspace
        Set AppObj = Application
    Else
        ' running in the background
            
        ' see if we can get the workspace object
        Set AppObj = GetObject(, "Workspace.Application")
        
        If AppObj Is Nothing Then
            Exit Sub
        End If
    End If
  
    gFileName = ""
    gFullPath = ""
    
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    If strChartName <> "" Then
    'do a find object on the chart name and if it exists, assign it to  gobjSelectedChart
        Set gobjSelectedChart = AppObj.ActiveDocument.Page.FindObject(strChartName) ' hj072903
    End If
    If AppObj.Mode = 4 Then ' hj072903
        frmChartGroupFileManagementForm.Show
    Else
        'Hook - PBH - See Hook() function for comments
        frmChartGroupForm.Show
    End If
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub

'*************************************************************************************
'
'   Subroutine:  SaveChartGroupFile(FileName as String)
'
'   Purpose:  Save a chart group file in folder under HTR path specified in the
'             SCU directory.
'
'   Inputs:  strFullPath:  Contains the full name of the chart group file
'
'*************************************************************************************
'Public Sub CGW_SaveChartGroupFile(strfullpath As String, Optional blnReadOnly As Boolean)
Public Function CGW_SaveChartGroupFile(strfullpath As String, Optional blnReadOnly As Boolean) As Boolean
    Dim strFileToSave As String     'the name of the file to save
    Dim intUserResponse As Integer  'the users response as to whether to overwrite the existing file
    Dim intFileHandle As String     'Used to create the new file
    Dim intForLoop As Integer       'Counter for the for Loop
    Dim intAttr As Integer          'Gets the return value when calling GetAttr to find out if a file is read-only.
    Dim blnNewFile As Boolean       'Set to true if the file is new.  False if the file already exists.
    Dim strError As String
    Dim strTitle As String
    Dim lngAttr As Long
    
    On Error GoTo ErrorHandler
    
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    'If save was clicked, and this is coming from the Save button of the CGC and not from the File form
    If gintSaveOkSaveAndApply = 3 And gblnSaveComingFromCGF Then
        'We have the file name, no need to search
        strFileToSave = strfullpath
    Else
        'Coming from the File Management dialog and needs
        'to determine if the file exists or not
        strFileToSave = Dir(strfullpath)
    End If
    
    'If a file name was found, prompt them to overwrite the file
    If strFileToSave <> "" Then
        ' PCM save code
        
        '#define PCMPATH_CHART_GROUP 4
        'CHECKTYPE_FILESAVE = 2
        
        Dim bDoSave As Boolean
        bDoSave = True
        'kei10122005 iFIX4.0 Trk #2517 we should always check PCM file status
        ' this code will get called from the save-as dialog too so if it was PCM that caused the save-as, then skip this code.
        'If (False = gblnSaveAsFromPCM) Then
        Dim bIsConnected As Boolean
        bIsConnected = ChangeManagerFunctions.IsConnected()
            
        If (True = bIsConnected) Then
            
            Const PCMPATH_CHART_GROUP As Long = 4
            Const CHECKTYPE_FILESAVE As Long = 2
                
            Dim ServerPath As String
            Dim bSuccess As Boolean
            Dim CGWPath As String
            CGWPath = System.FixPath(ghtc_path)
            bSuccess = System.PCM_GetServerPathByIndex(PCMPATH_CHART_GROUP, ServerPath)
            If (True = bSuccess) Then
                Dim SetWDRetVal As Long
                SetWDRetVal = ChangeManagerFunctions.SetWorkingDirectory(CGWPath, ServerPath)
                    
                If (0 = SetWDRetVal) Then
                    Dim EmptyStringArray() As String
                    Dim CurrentUser As String
                    Dim bPCMAccess As Boolean
                    
                    Dim CheckOutRetVal As Long
                    Dim sNodeName As String
                    Dim bKeepLocalCopy As Boolean
                    Dim bProceed As Boolean
                    Dim FileNameNoPath As String
                    'kei10122005 iFIX4.0 Trk #2517 get file name from full path
                    FileNameNoPath = Dir(strfullpath)
                    'kei09222005 iFIX4.0 Trk #2388 - set true to Keep local copy
                    bKeepLocalCopy = True
                    CheckOutRetVal = ChangeManagerFunctions.IsFileCheckedOut(FileNameNoPath, CHECKTYPE_FILESAVE, FileNameNoPath, EmptyStringArray, False, bKeepLocalCopy, bProceed, 0)
                    If (False = bProceed) Then
                        bDoSave = False
                    End If
                Else
                    bDoSave = False
                End If
            Else
                bDoSave = False
            End If
        End If
                
        If (True = bDoSave) Then
            Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
            strError = objStrMgr.GetNLSStr(CLng(NLS_OVERWRITE))
            strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
            intUserResponse = MsgBox(strError, vbYesNo, strTitle)
            If intUserResponse = vbNo Then
                gblnBadFileName = True
                CGW_SaveChartGroupFile = True 'kei10122005 iFIX4.0 Trk #2517
                Exit Function
            Else
                strFileToSave = strfullpath
            End If
            lngAttr = GetAttr(strFileToSave)
            If (lngAttr Mod 2) > 0 Then
                'read only error
                strError = strFileToSave & vbCrLf & objStrMgr.GetNLSStr(CLng(NLS_SAVEERROR2)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_SAVEERROR3))
                strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
                MsgBox strError, vbOKOnly, strTitle
                CGW_SaveChartGroupFile = True 'kei10122005 iFIX4.0 Trk #2517
                Exit Function
            End If
        Else
            'kei10122005 iFIX4.0 Trk #2517
            ' return false here instead of calling frmChartGroupFileManagementForm
            ' caller should check return value
            CGW_SaveChartGroupFile = False
            Exit Function
        End If
            
    ElseIf strFileToSave = "" Then
        'this is a new file
        strFileToSave = strfullpath
        blnNewFile = True
    End If
      
    intFileHandle = FreeFile
    
    'If the file does not exist and user want it read only.
    If blnReadOnly And Not (blnNewFile) Then
        Open strFileToSave For Output Lock Read As #intFileHandle
    Else
        Open strFileToSave For Output As #intFileHandle
    End If
        
        Write #intFileHandle, objStrMgr.GetNLSStr(CLng(NLS_EXC)), objStrMgr.GetNLSStr(CLng(NLS_APPLYALL)), objStrMgr.GetNLSStr(CLng(NLS_DAYSBEFORENOW)), objStrMgr.GetNLSStr(CLng(NLS_DURATION)), objStrMgr.GetNLSStr(CLng(NLS_FETCHPENLIMITS)), objStrMgr.GetNLSStr(CLng(NLS_FIXEDDATE)), objStrMgr.GetNLSStr(CLng(NLS_FIXEDTIME)), objStrMgr.GetNLSStr(CLng(NLS_HILIMIT)), objStrMgr.GetNLSStr(CLng(NLS_HISTORICALSAMPLETYPE)), objStrMgr.GetNLSStr(CLng(NLS_LOLIMIT)), objStrMgr.GetNLSStr(CLng(NLS_MARKERCHAR)), objStrMgr.GetNLSStr(CLng(NLS_MARKERSTYLE)), objStrMgr.GetNLSStr(CLng(NLS_PENLINECOLOR)), objStrMgr.GetNLSStr(CLng(NLS_PENLINESTYLE)), objStrMgr.GetNLSStr(CLng(NLS_PENLINEWIDTH)), objStrMgr.GetNLSStr(CLng(NLS_SOURCE)), objStrMgr.GetNLSStr(CLng(NLS_STARTDATETYPE)), objStrMgr.GetNLSStr(CLng(NLS_STARTTIMETYPE)), objStrMgr.GetNLSStr(CLng(NLS_TIMEBEFORENOW)), _
                objStrMgr.GetNLSStr(CLng(NLS_INTERVAL)), objStrMgr.GetNLSStr(CLng(NLS_DISPLAYMS)), objStrMgr.GetNLSStr(CLng(NLS_DST_HEAD)), objStrMgr.GetNLSStr(CLng(NLS_TIMEZONE_HEAD))
        For intForLoop = LBound(guPenPropertiesArray) To UBound(guPenPropertiesArray)
            'jes clarify #235605
            Write #intFileHandle, "&&", guPenPropertiesArray(intForLoop).ApplyToAll, guPenPropertiesArray(intForLoop).DaysBeforeNow, guPenPropertiesArray(intForLoop).Duration, guPenPropertiesArray(intForLoop).FetchPenLimits, guPenPropertiesArray(intForLoop).FixedDate, Format(guPenPropertiesArray(intForLoop).FixedTime, "Short Time"), guPenPropertiesArray(intForLoop).HiLimit, guPenPropertiesArray(intForLoop).HistoricalSampleType, guPenPropertiesArray(intForLoop).LoLimit, guPenPropertiesArray(intForLoop).MarkerChar, guPenPropertiesArray(intForLoop).MarkerStyle, guPenPropertiesArray(intForLoop).PenLineColor, guPenPropertiesArray(intForLoop).PenLineStyle, guPenPropertiesArray(intForLoop).PenLineWidth, guPenPropertiesArray(intForLoop).Source, guPenPropertiesArray(intForLoop).StartDateType, guPenPropertiesArray(intForLoop).StartTimeType, guPenPropertiesArray(intForLoop).TimeBeforeNow, _
                    guPenPropertiesArray(intForLoop).Interval, guPenPropertiesArray(intForLoop).DisplayMS, guPenPropertiesArray(intForLoop).DaylightSavingsTime, guPenPropertiesArray(intForLoop).TimeZone
        Next intForLoop
    Close #intFileHandle
    DoEvents

    CGW_SaveChartGroupFile = True 'kei10122005 iFIX4.0 Trk #2517
    Exit Function

ErrorHandler:
    gblnBadFileName = True
    If Err.Number = g_NT_SECURITY_ERROR Then
        strError = strFileToSave & vbCrLf & objStrMgr.GetNLSStr(CLng(NLS_SAVEERROR2)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_SAVEERROR3))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
        MsgBox strError, vbOKOnly, strTitle
    Else
        HandleError
    End If
    CGW_SaveChartGroupFile = True 'kei10122005 iFIX4.0 Trk #2517
    Exit Function
End Function

'*************************************************************************************
'
'Sub Routine CGW_OpenChartGroupFile (strFileName As String)
'
'Purpose to read chart group information from the file
'
'Inputs:   strFileName: contains full path of the file to be read from the disk
'
'*************************************************************************************

Public Sub CGW_OpenChartGroupFile(strFileName As String)

    Dim uTempPenPropertiesArray() As gPenProperties
    Dim intFileHandle As String
    Dim intForLoop As Integer
    Dim intCount As Integer
    Dim strTempStorage As String
    Dim strPropertiesNames() As Variant
    Dim strPenPropertyName() As String
    Dim strTempValueStorage As Variant
    Dim strCasePropertiesNames As String
    Dim strError As String
    Dim strTitle As String
    
    On Error GoTo ErrorHandler
    CGW_ChartGroupFileVersion (strFileName)
    
    Select Case giVer
    Case 0
        ReDim strPropertiesNames(18)
        ReDim strPenPropertyName(17)
    Case Else
        ReDim strPropertiesNames(22)
        ReDim strPenPropertyName(21)
    End Select
    
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    intForLoop = 0
    intCount = 0
    ReDim Preserve uTempPenPropertiesArray(0)
    intFileHandle = FreeFile
    Open strFileName For Input As #intFileHandle
    Do Until EOF(intFileHandle)
    If giVer = 0 Then
        Input #1, strPropertiesNames(0), strPropertiesNames(1), strPropertiesNames(2), strPropertiesNames(3), strPropertiesNames(4), strPropertiesNames(5), strPropertiesNames(6), strPropertiesNames(7), strPropertiesNames(8), strPropertiesNames(9), strPropertiesNames(10), strPropertiesNames(11), strPropertiesNames(12), strPropertiesNames(13), strPropertiesNames(14), strPropertiesNames(15), strPropertiesNames(16), strPropertiesNames(17), strPropertiesNames(18)
    Else
        Input #1, strPropertiesNames(0), strPropertiesNames(1), strPropertiesNames(2), strPropertiesNames(3), strPropertiesNames(4), strPropertiesNames(5), strPropertiesNames(6), strPropertiesNames(7), strPropertiesNames(8), strPropertiesNames(9), strPropertiesNames(10), strPropertiesNames(11), strPropertiesNames(12), strPropertiesNames(13), strPropertiesNames(14), strPropertiesNames(15), strPropertiesNames(16), strPropertiesNames(17), strPropertiesNames(18), strPropertiesNames(19), strPropertiesNames(20), strPropertiesNames(21), strPropertiesNames(22)
    End If
        If Left(strPropertiesNames(0), 1) <> "'" Then
            If Len(strPropertiesNames(0)) <> 0 Then
                If Left(strPropertiesNames(0), 1) = "!" Then
                    For intForLoop = 1 To UBound(strPropertiesNames)
                        strPenPropertyName(intForLoop - 1) = CStr(strPropertiesNames(intForLoop))
                    Next intForLoop
                ElseIf strPropertiesNames(0) = "&&" Then
                     
                    For intForLoop = 0 To UBound(strPenPropertyName)
                    strCasePropertiesNames = UCase(strPenPropertyName(intForLoop))
                
                        Select Case strCasePropertiesNames
                        Case objStrMgr.GetNLSStr(CLng(NLS_APPLYALL))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).ApplyToAll = CBool(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_DAYSBEFORENOW))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).DaysBeforeNow = CInt(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_DURATION))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).Duration = CLng(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_FETCHPENLIMITS))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).FetchPenLimits = CBool(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_FIXEDDATE))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).FixedDate = CDate(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_FIXEDTIME))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).FixedTime = CDate(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_HILIMIT))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).HiLimit = CDbl(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_HISTORICALSAMPLETYPE))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).HistoricalSampleType = CInt(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_LOLIMIT))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).LoLimit = CDbl(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_MARKERCHAR))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).MarkerChar = CStr(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_MARKERSTYLE))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).MarkerStyle = CInt(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_PENLINECOLOR))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).PenLineColor = CLng(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_PENLINESTYLE))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).PenLineStyle = CInt(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_PENLINEWIDTH))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).PenLineWidth = CLng(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_SOURCE))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).Source = CStr(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_STARTTIMETYPE))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).StartTimeType = CInt(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_TIMEBEFORENOW))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).TimeBeforeNow = CLng(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_STARTDATETYPE))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).StartDateType = CInt(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_INTERVAL))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).Interval = CLng(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_DISPLAYMS))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).DisplayMS = CBool(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_DST_HEAD))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).DaylightSavingsTime = CBool(strPropertiesNames(intForLoop + 1))
                        Case objStrMgr.GetNLSStr(CLng(NLS_TIMEZONE_HEAD))
                            On Error GoTo WrongFormat
                            uTempPenPropertiesArray(intCount).TimeZone = CInt(strPropertiesNames(intForLoop + 1))
                        Case Else
                            GoTo WrongFormat
                        End Select
                    Next intForLoop
            
                    If intCount + 1 > UBound(uTempPenPropertiesArray) Then
                        ReDim Preserve uTempPenPropertiesArray(intCount + 1)
                    End If
                
                    intCount = intCount + 1
            
                ElseIf (Left(strPropertiesNames(0), 1) <> "!") And (strPropertiesNames(0) <> "&&") Then
                    'this is a corrupted file or not a file created for the CGW
                    gblnBadFileName = True
                    strError = objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT1)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT2)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT3)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT4)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT5)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT6))
                    strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
                    MsgBox strError, vbExclamation, strTitle
                    Close #intFileHandle
                    Exit Sub
                End If
            End If
        End If
         
    Loop


    Close #intFileHandle
    ReDim Preserve uTempPenPropertiesArray(intCount - 1)
    ReDim guPenPropertiesArray(0 To UBound(uTempPenPropertiesArray))

    For intForLoop = 0 To UBound(guPenPropertiesArray)
        
        guPenPropertiesArray(intForLoop).ApplyToAll = uTempPenPropertiesArray(intForLoop).ApplyToAll
        guPenPropertiesArray(intForLoop).DaysBeforeNow = uTempPenPropertiesArray(intForLoop).DaysBeforeNow
        guPenPropertiesArray(intForLoop).Duration = uTempPenPropertiesArray(intForLoop).Duration
        guPenPropertiesArray(intForLoop).FetchPenLimits = uTempPenPropertiesArray(intForLoop).FetchPenLimits
        guPenPropertiesArray(intForLoop).FixedDate = uTempPenPropertiesArray(intForLoop).FixedDate
        guPenPropertiesArray(intForLoop).FixedTime = uTempPenPropertiesArray(intForLoop).FixedTime
        guPenPropertiesArray(intForLoop).HiLimit = uTempPenPropertiesArray(intForLoop).HiLimit
        guPenPropertiesArray(intForLoop).HistoricalSampleType = uTempPenPropertiesArray(intForLoop).HistoricalSampleType
        guPenPropertiesArray(intForLoop).LoLimit = uTempPenPropertiesArray(intForLoop).LoLimit
        guPenPropertiesArray(intForLoop).MarkerChar = uTempPenPropertiesArray(intForLoop).MarkerChar
        guPenPropertiesArray(intForLoop).MarkerStyle = uTempPenPropertiesArray(intForLoop).MarkerStyle
        guPenPropertiesArray(intForLoop).PenLineColor = uTempPenPropertiesArray(intForLoop).PenLineColor
        guPenPropertiesArray(intForLoop).PenLineWidth = uTempPenPropertiesArray(intForLoop).PenLineWidth
        guPenPropertiesArray(intForLoop).PenLineStyle = uTempPenPropertiesArray(intForLoop).PenLineStyle
        guPenPropertiesArray(intForLoop).TimeBeforeNow = uTempPenPropertiesArray(intForLoop).TimeBeforeNow
        guPenPropertiesArray(intForLoop).StartDateType = uTempPenPropertiesArray(intForLoop).StartDateType
        guPenPropertiesArray(intForLoop).StartTimeType = uTempPenPropertiesArray(intForLoop).StartTimeType
        guPenPropertiesArray(intForLoop).Source = uTempPenPropertiesArray(intForLoop).Source
        If giVer = 0 Then 'older files default
            guPenPropertiesArray(intForLoop).Interval = gintInterval
            guPenPropertiesArray(intForLoop).DisplayMS = gintDisplayMS
            guPenPropertiesArray(intForLoop).DaylightSavingsTime = gintDaylightSavingsTime
            guPenPropertiesArray(intForLoop).TimeZone = gintTimeZone
        Else
            guPenPropertiesArray(intForLoop).Interval = uTempPenPropertiesArray(intForLoop).Interval
            guPenPropertiesArray(intForLoop).DisplayMS = uTempPenPropertiesArray(intForLoop).DisplayMS
            guPenPropertiesArray(intForLoop).DaylightSavingsTime = uTempPenPropertiesArray(intForLoop).DaylightSavingsTime
            guPenPropertiesArray(intForLoop).TimeZone = uTempPenPropertiesArray(intForLoop).TimeZone
        End If
    Next intForLoop

    Exit Sub
    
WrongFormat:
    'this is a corrupted file or not a file created for the CGW
        gblnBadFileName = True
        strError = objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT1)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT2)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT3)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT4)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT5)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_WRONGFORMAT6))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
        MsgBox strError, vbExclamation, strTitle
        Close #intFileHandle
        Exit Sub
        
ErrorHandler:
    gblnBadFileName = True
    If Err.Number = 76 Then
        strError = objStrMgr.GetNLSStr(CLng(NLS_CHARTDNE))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
        MsgBox strError, , strTitle
    ElseIf Err.Number = 65 Then
        strError = objStrMgr.GetNLSStr(CLng(NLS_BADFORMAT1)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_BADFORMAT2))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
        MsgBox strError, vbExclamation, strTitle
    Else
        HandleError
    End If
    Close #intFileHandle

End Sub

'*************************************************************************************
'
'Sub Routine CGW_SetChartGroupFileName(objObjectSelected As Object, strFileName As String)
'
'Purpose:  To apply Filename to InitialValue property of the string variable
'
'Inputs:   strFileName: contains the complete filename
'
'*************************************************************************************

Public Sub CGW_SetChartGroupFileName(objObjectSelected As Object, strFileName As String)

    Dim subobjObjectSelected As Object
    
    On Error GoTo ErrorHandler
    If objObjectSelected.ClassName = "Variable" And Left(objObjectSelected.Name, 8) = "FileName" Then
      objObjectSelected.InitialValue = strFileName
    End If
    
    For Each subobjObjectSelected In objObjectSelected.ContainedObjects
        CGW_SetChartGroupFileName subobjObjectSelected, strFileName
    Next subobjObjectSelected
    Exit Sub
    
ErrorHandler:
    HandleError
    
End Sub

'*************************************************************************************
'
'Sub Routine CGW_GetChartGroupFileName(objObjectSelected As Object, strFileName As String)
'
'Purpose:  To apply Filename to InitialValue property of the string varaiable
'
'Inputs:   strFileName: contains the complete filename
'
'*************************************************************************************

Public Sub CGW_GetChartGroupFileName(objObjectSelected As Object)

    Dim subobjObjectSelected As Object
    Dim strError As String
    Dim strTitle As String
    
    On Error GoTo ErrorHandler
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    If (TypeName(objObjectSelected) = "Nothing") Then
        strError = objStrMgr.GetNLSStr(CLng(NLS_NOCHART))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
        MsgBox strError, , strTitle
        End
    End If
    If objObjectSelected.ClassName = "Variable" And Left(objObjectSelected.Name, 8) = "FileName" Then
        gFullPath = objObjectSelected.InitialValue
    End If
    
    For Each subobjObjectSelected In objObjectSelected.ContainedObjects
        If subobjObjectSelected.ClassName = "Variable" Then
            CGW_GetChartGroupFileName subobjObjectSelected
        End If
    Next subobjObjectSelected
    Exit Sub

ErrorHandler:
    HandleError
    
End Sub

Public Sub CGW_ApplyChartGroupSettings(objChart As Object)

    Dim objSubChart As Object
    Dim intForLoop As Integer
    Dim strError As String
    Dim strTitle As String
    Dim blnErrorDisplayed As Boolean
    Dim lngStatus As Long
    Dim objRetObject As Object
    Dim strPropName As String
    Dim bAllowValueAxisReset As Boolean 'JPB010203 port hj110802
    Dim dtToday As Date   'mr050108 Ported from 4.0
                
    On Error GoTo ErrorHandler
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    'mr050108 Ported from 4.0 Get today's date
    dtToday = Format(Now, "Short Date")
    If objChart.ClassName = "Chart" Then
        'If there are the same number of pens being applied than there already displayed in the
        'chart, just loop through them and re-assign the new values
        If UBound(guPenPropertiesArray) = objChart.pens.Count - 1 Then
            For intForLoop = 1 To objChart.pens.Count
                'hj052704
                objChart.pens.Item(intForLoop).FetchPenLimits = guPenPropertiesArray(intForLoop - 1).FetchPenLimits
                'If source exists, use SetSource, if it DNE, use Source
                If guPenPropertiesArray(intForLoop - 1).Source <> "" Then
                    objChart.pens.Item(intForLoop).SetSource guPenPropertiesArray(intForLoop - 1).Source, True
                End If
                'find out if the source is a use anyway or if it is real, if useanyway
                'then set its CurrentValue to 0
                System.ValidateSource guPenPropertiesArray(intForLoop - 1).Source, lngStatus, objRetObject, strPropName
                If lngStatus <> 0 Then
                    objChart.pens.Item(intForLoop).legend.legendvalue = "****"
                    objChart.pens.Item(intForLoop).legend.legenddesc = ""
                    If objChart.pens.Item(intForLoop).CurrentValue <> 0 And objChart.pens.Item(intForLoop).legend.legendvalue <> "****" Then
                        objChart.pens.Item(intForLoop).CurrentValue = 0
                    End If
                End If
                objChart.pens.Item(intForLoop).DaysBeforeNow = guPenPropertiesArray(intForLoop - 1).DaysBeforeNow
                objChart.pens.Item(intForLoop).Duration = guPenPropertiesArray(intForLoop - 1).Duration
                'mr050108 Port from 4.0 thc070907  1-208005699 Overwrite hi & lo limit with values from file only if FetchPenLimits is False
                'If FetchPenLimits is True, the limits were fetched earlier when SetSource was called.
                If guPenPropertiesArray(intForLoop - 1).FetchPenLimits = False Then
                    'JPB010203 port hj110802 Use AllowValueAxisReset to make sure we can set the high and low limits into the pen
                    bAllowValueAxisReset = objChart.pens.Item(intForLoop).AllowValueAxisReset
                    objChart.pens.Item(intForLoop).AllowValueAxisReset = True
                    objChart.pens.Item(intForLoop).HiLimit = guPenPropertiesArray(intForLoop - 1).HiLimit
                    objChart.pens.Item(intForLoop).LoLimit = guPenPropertiesArray(intForLoop - 1).LoLimit
                    'JPB010203 port hj110802
                    objChart.pens.Item(intForLoop).AllowValueAxisReset = bAllowValueAxisReset
                End If
                
                objChart.pens.Item(intForLoop).FetchPenLimits = guPenPropertiesArray(intForLoop - 1).FetchPenLimits
                'mr050108 Port from 4.0 thc050707 Set StartTime later
               'objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                objChart.pens.Item(intForLoop).HistoricalSampleType = guPenPropertiesArray(intForLoop - 1).HistoricalSampleType
                objChart.pens.Item(intForLoop).MarkerChar = guPenPropertiesArray(intForLoop - 1).MarkerChar
                objChart.pens.Item(intForLoop).MarkerStyle = guPenPropertiesArray(intForLoop - 1).MarkerStyle
                objChart.pens.Item(intForLoop).PenLineColor = guPenPropertiesArray(intForLoop - 1).PenLineColor
                objChart.pens.Item(intForLoop).PenLineStyle = CLng(guPenPropertiesArray(intForLoop - 1).PenLineStyle)
                objChart.pens.Item(intForLoop).PenLineWidth = guPenPropertiesArray(intForLoop - 1).PenLineWidth
                objChart.pens.Item(intForLoop).StartDateType = guPenPropertiesArray(intForLoop - 1).StartDateType
                objChart.pens.Item(intForLoop).StartTimeType = guPenPropertiesArray(intForLoop - 1).StartTimeType
                
                'mr050108 Port from 4.0 thc050707 TimeBeforeNow on chart object needs to be 24 hours or less.  So take the seconds
                ' associated with the DaysBeforeNow out.
                ' objChart.pens.Item(intForLoop).TimeBeforeNow = guPenPropertiesArray(intForLoop - 1).TimeBeforeNow
                objChart.pens.Item(intForLoop).TimeBeforeNow = CGW_CalcRelativeTime(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)

                'mr050108 Port from 4.0 thc050707 Added 2 cases to this If to handle DateFixed/TimeRelative and DateRelative/TimeFixed
                If guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed And guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    'hj121807 Ported from 3.5 - hj022304 If the FixedDate is the system default date (12/30/1899), don't use it. Instead, use today's date.
                    'objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    ''hj121401 Need set these two properties if using FixedStartTime
                    'objChart.pens.Item(intForLoop).FixedDate = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    'objChart.pens.Item(intForLoop).FixedTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    If guPenPropertiesArray(intForLoop - 1).FixedDate = gdtSystemDefaultDate Then
                        'hj012308 Changed to not format the date and time
                        'objChart.pens.Item(intForLoop).StartTime = CStr(Format(dtToday, "Short Date") + " " + Format(guPenPropertiesArray(intForLoop - 1).FixedTime, "Short Time"))
                        objChart.pens.Item(intForLoop).StartTime = CStr(dtToday) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    Else
                        'hj012308 Changed to not format the date and time
                        'objChart.pens.Item(intForLoop).StartTime = CStr(Format(guPenPropertiesArray(intForLoop - 1).FixedDate, "Short Date") + " " + Format(guPenPropertiesArray(intForLoop - 1).FixedTime, "Short Time"))
                        objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    End If
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintRelative And guPenPropertiesArray(intForLoop - 1).StartDateType = gintRelative Then
                    objChart.pens.Item(intForLoop).StartTime = CGW_SetDate(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintRelative And guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    'hj121807 If the FixedDate is the system default date (12/30/1899), don't use it. Instead, use today's date.
                    If guPenPropertiesArray(intForLoop - 1).FixedDate = gdtSystemDefaultDate Then
                        objChart.pens.Item(intForLoop).StartTime = CGW_SetRelativeTime(dtToday, guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                    Else
                        objChart.pens.Item(intForLoop).StartTime = CGW_SetRelativeTime(guPenPropertiesArray(intForLoop - 1).FixedDate, guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                    End If
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed And guPenPropertiesArray(intForLoop - 1).StartDateType = gintRelative Then
                    objChart.pens.Item(intForLoop).StartTime = CStr(CGW_SetRelativeDate(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow)) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                End If
                
                'mr050108 Port from 4.0 hj121807 Set these two properties if necessary
                If guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    objChart.pens.Item(intForLoop).FixedDate = objChart.pens.Item(intForLoop).StartTime
                End If
                If guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed Then
                    objChart.pens.Item(intForLoop).FixedTime = objChart.pens.Item(intForLoop).StartTime
                End If
                
                objChart.pens.Item(intForLoop).IntervalMilliseconds = (guPenPropertiesArray(intForLoop - 1).Interval) Mod 1000
                objChart.pens.Item(intForLoop).Interval = (guPenPropertiesArray(intForLoop - 1).Interval - objChart.pens.Item(intForLoop).IntervalMilliseconds) / 1000
                objChart.pens.Item(intForLoop).DisplayMilliseconds = guPenPropertiesArray(intForLoop - 1).DisplayMS
                objChart.pens.Item(intForLoop).DaylightSavingsTime = guPenPropertiesArray(intForLoop - 1).DaylightSavingsTime
                If guPenPropertiesArray(intForLoop - 1).TimeZone > 2 Then
                    'hj111204 Should set the property with Long type of data instead of Integer type
                    'objChart.pens.Item(intForLoop).TimeZoneBiasRelative = 3
                    'objChart.pens.Item(intForLoop).TimeZoneBiasExplicit = guPenPropertiesArray(intForLoop - 1).TimeZone - 3
                    objChart.pens.Item(intForLoop).TimeZoneBiasRelative = CLng(3)
                    objChart.pens.Item(intForLoop).TimeZoneBiasExplicit = CLng(guPenPropertiesArray(intForLoop - 1).TimeZone - 3)
                Else
                    'hj111204 Should set the property with Long type of data instead of Integer type
                    'objChart.pens.Item(intForLoop).TimeZoneBiasRelative = guPenPropertiesArray(intForLoop - 1).TimeZone
                    objChart.pens.Item(intForLoop).TimeZoneBiasRelative = CLng(guPenPropertiesArray(intForLoop - 1).TimeZone)
                End If
            Next intForLoop
        'Applying more pens to the chart than what is displayed there now.
        'For as many pens displayed previously, re-assign the new values
        'For the additional pens, we must AddPen then SetProperty
        ElseIf UBound(guPenPropertiesArray) > objChart.pens.Count - 1 Then
            For intForLoop = 1 To objChart.pens.Count
                'hj052704
                objChart.pens.Item(intForLoop).FetchPenLimits = guPenPropertiesArray(intForLoop - 1).FetchPenLimits
                'If source exists, use SetSource, if it DNE, use Source
                If guPenPropertiesArray(intForLoop - 1).Source <> "" Then
                    objChart.pens.Item(intForLoop).SetSource guPenPropertiesArray(intForLoop - 1).Source, True
                End If
                'find out if the source is a use anyway or if it is real, if useanyway
                'then set its CurrentValue to 0
                System.ValidateSource guPenPropertiesArray(intForLoop - 1).Source, lngStatus, objRetObject, strPropName
                If lngStatus <> 0 Then
                    objChart.pens.Item(intForLoop).legend.legendvalue = "****"
                    objChart.pens.Item(intForLoop).legend.legenddesc = ""
                    If objChart.pens.Item(intForLoop).CurrentValue <> 0 And objChart.pens.Item(intForLoop).legend.legendvalue <> "****" Then
                        objChart.pens.Item(intForLoop).CurrentValue = 0
                    End If
                End If
                objChart.pens.Item(intForLoop).DaysBeforeNow = guPenPropertiesArray(intForLoop - 1).DaysBeforeNow
                objChart.pens.Item(intForLoop).Duration = guPenPropertiesArray(intForLoop - 1).Duration
                
                'mr050108 Port from 4.0thc070907  1-208005699 Overwrite hi & lo limit with values from file only if FetchPenLimits is False
                'If FetchPenLimits is True, the limits were fetched earlier when SetSource was called.
                If guPenPropertiesArray(intForLoop - 1).FetchPenLimits = False Then
                    'JPB010203 port hj110802 Use AllowValueAxisReset to make sure we can set the high and low limits into the pen
                    bAllowValueAxisReset = objChart.pens.Item(intForLoop).AllowValueAxisReset
                    objChart.pens.Item(intForLoop).AllowValueAxisReset = True
                    objChart.pens.Item(intForLoop).HiLimit = guPenPropertiesArray(intForLoop - 1).HiLimit
                    objChart.pens.Item(intForLoop).LoLimit = guPenPropertiesArray(intForLoop - 1).LoLimit
                    'JPB010203 port hj110802
                    objChart.pens.Item(intForLoop).AllowValueAxisReset = bAllowValueAxisReset
                End If
                
                objChart.pens.Item(intForLoop).FetchPenLimits = guPenPropertiesArray(intForLoop - 1).FetchPenLimits
                'mr050108 Port from 4.0 thc050707 Set StartTime later
                'objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                objChart.pens.Item(intForLoop).HistoricalSampleType = guPenPropertiesArray(intForLoop - 1).HistoricalSampleType
                objChart.pens.Item(intForLoop).MarkerChar = guPenPropertiesArray(intForLoop - 1).MarkerChar
                objChart.pens.Item(intForLoop).MarkerStyle = guPenPropertiesArray(intForLoop - 1).MarkerStyle
                objChart.pens.Item(intForLoop).PenLineColor = guPenPropertiesArray(intForLoop - 1).PenLineColor
                objChart.pens.Item(intForLoop).PenLineStyle = CLng(guPenPropertiesArray(intForLoop - 1).PenLineStyle)
                objChart.pens.Item(intForLoop).PenLineWidth = guPenPropertiesArray(intForLoop - 1).PenLineWidth
                objChart.pens.Item(intForLoop).StartDateType = guPenPropertiesArray(intForLoop - 1).StartDateType
                objChart.pens.Item(intForLoop).StartTimeType = guPenPropertiesArray(intForLoop - 1).StartTimeType
                
                'mr050108 Port from 4.0 thc050707 TimeBeforeNow on chart object needs to be 24 hours or less.  So take the seconds
                ' associated with the DaysBeforeNow out.
                'objChart.pens.Item(intForLoop).TimeBeforeNow = guPenPropertiesArray(intForLoop - 1).TimeBeforeNow
                objChart.pens.Item(intForLoop).TimeBeforeNow = CGW_CalcRelativeTime(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                
                'mr050108 Port from 4.0 thc050707 Added 2 cases to this If to handle DateFixed/TimeRelative and DateRelative/TimeFixed
                If guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed And guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    'hj121807 Ported from 3.5 - hj022304 If the FixedDate is the system default date (12/30/1899), don't use it. Instead, use today's date.
                    'objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    ''hj121401 Need set these two properties if using FixedStartTime
                    'objChart.pens.Item(intForLoop).FixedDate = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    'objChart.pens.Item(intForLoop).FixedTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    If guPenPropertiesArray(intForLoop - 1).FixedDate = gdtSystemDefaultDate Then
                        'hj012308 Changed to not format the date and time
                        'objChart.pens.Item(intForLoop).StartTime = CStr(Format(dtToday, "Short Date") + " " + Format(guPenPropertiesArray(intForLoop - 1).FixedTime, "Short Time"))
                        objChart.pens.Item(intForLoop).StartTime = CStr(dtToday) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    Else
                        'hj012308 Changed to not format the date and time
                        'objChart.pens.Item(intForLoop).StartTime = CStr(Format(guPenPropertiesArray(intForLoop - 1).FixedDate, "Short Date") + " " + Format(guPenPropertiesArray(intForLoop - 1).FixedTime, "Short Time"))
                        objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    End If
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintRelative And guPenPropertiesArray(intForLoop - 1).StartDateType = gintRelative Then
                    objChart.pens.Item(intForLoop).StartTime = CGW_SetDate(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintRelative And guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    'mr050108 Port from 4.0 hj121807 If the FixedDate is the system default date (12/30/1899), don't use it. Instead, use today's date.
                    If guPenPropertiesArray(intForLoop - 1).FixedDate = gdtSystemDefaultDate Then
                        objChart.pens.Item(intForLoop).StartTime = CGW_SetRelativeTime(dtToday, guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                    Else
                        objChart.pens.Item(intForLoop).StartTime = CGW_SetRelativeTime(guPenPropertiesArray(intForLoop - 1).FixedDate, guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                    End If
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed And guPenPropertiesArray(intForLoop - 1).StartDateType = gintRelative Then
                    objChart.pens.Item(intForLoop).StartTime = CStr(CGW_SetRelativeDate(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow)) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                End If
                    
                'mr050108 Port from 4.0 hj121807 Set these two properties if necessary
                If guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    objChart.pens.Item(intForLoop).FixedDate = objChart.pens.Item(intForLoop).StartTime
                End If
                If guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed Then
                    objChart.pens.Item(intForLoop).FixedTime = objChart.pens.Item(intForLoop).StartTime
                End If
                
                objChart.pens.Item(intForLoop).IntervalMilliseconds = guPenPropertiesArray(intForLoop - 1).Interval Mod 1000
                objChart.pens.Item(intForLoop).Interval = (guPenPropertiesArray(intForLoop - 1).Interval - objChart.pens.Item(intForLoop).IntervalMilliseconds) / 1000
                objChart.pens.Item(intForLoop).DisplayMilliseconds = guPenPropertiesArray(intForLoop - 1).DisplayMS
                objChart.pens.Item(intForLoop).DaylightSavingsTime = guPenPropertiesArray(intForLoop - 1).DaylightSavingsTime
                If guPenPropertiesArray(intForLoop - 1).TimeZone > 2 Then
                    'hj111204 Should set the property with Long type of data instead of Integer type
                    'objChart.pens.Item(intForLoop).TimeZoneBiasRelative = 3
                    'objChart.pens.Item(intForLoop).TimeZoneBiasExplicit = guPenPropertiesArray(intForLoop - 1).TimeZone - 3
                    objChart.pens.Item(intForLoop).TimeZoneBiasRelative = CLng(3)
                    objChart.pens.Item(intForLoop).TimeZoneBiasExplicit = CLng(guPenPropertiesArray(intForLoop - 1).TimeZone - 3)
                Else
                    'hj111204 Should set the property with Long type of data instead of Integer type
                    'objChart.pens.Item(intForLoop).TimeZoneBiasRelative = guPenPropertiesArray(intForLoop - 1).TimeZone
                    objChart.pens.Item(intForLoop).TimeZoneBiasRelative = CLng(guPenPropertiesArray(intForLoop - 1).TimeZone)
                End If
            Next intForLoop
        
            For intForLoop = objChart.pens.Count + 1 To UBound(guPenPropertiesArray) + 1
                'siva 12/16/2011 - Sending newly added pen's name instead of pen's number
                'objChart.addpen (intForLoop)
                'objChart.addpen (objChart.pens.Item(intForLoop - 1).Source)
                'MTK 04/14/2014 addpen will result in error when addpen is empty(1-3487849151).
                'Adding pens to chart from guPenPropertiesArray.
                objChart.addpen (guPenPropertiesArray(intForLoop - 1).Source)
                
                'If source exists, use SetSource, if it DNE, use Source
                On Error GoTo ErrorHandler
                'Validate Source is not working right now. With a source that existing, but
                'not in the locally loaded database, ValidateSource returns a Tag Group Object
                'instead of an error.
                'hj052704
                objChart.pens.Item(intForLoop).FetchPenLimits = guPenPropertiesArray(intForLoop - 1).FetchPenLimits
                'If source exists, use SetSource, if it DNE, use Source
                If guPenPropertiesArray(intForLoop - 1).Source <> "" Then
                    objChart.pens.Item(intForLoop).SetSource guPenPropertiesArray(intForLoop - 1).Source, True
                End If
                'find out if the source is a use anyway or if it is real, if useanyway
                'then set its CurrentValue to 0
                System.ValidateSource guPenPropertiesArray(intForLoop - 1).Source, lngStatus, objRetObject, strPropName
                If lngStatus <> 0 Then
                    objChart.pens.Item(intForLoop).legend.legendvalue = "****"
                    objChart.pens.Item(intForLoop).legend.legenddesc = ""
                    If objChart.pens.Item(intForLoop).CurrentValue <> 0 And objChart.pens.Item(intForLoop).legend.legendvalue <> "****" Then
                        objChart.pens.Item(intForLoop).CurrentValue = 0
                    End If
                End If
                objChart.pens.Item(intForLoop).SetProperty "DaysBeforeNow", guPenPropertiesArray(intForLoop - 1).DaysBeforeNow
                objChart.pens.Item(intForLoop).SetProperty "Duration", guPenPropertiesArray(intForLoop - 1).Duration
                
                'mr050108 Ported from 4.0 thc070907  1-208005699 Overwrite hi & lo limit with values from file only if FetchPenLimits is False
                'If FetchPenLimits is True, the limits were fetched earlier when SetSource was called.
                If guPenPropertiesArray(intForLoop - 1).FetchPenLimits = False Then
                    'JPB010203 port hj110802 Use AllowValueAxisReset to make sure we can set the high and low limits into the pen
                    bAllowValueAxisReset = objChart.pens.Item(intForLoop).AllowValueAxisReset
                    objChart.pens.Item(intForLoop).AllowValueAxisReset = True
                    objChart.pens.Item(intForLoop).SetProperty "HiLimit", guPenPropertiesArray(intForLoop - 1).HiLimit
                    objChart.pens.Item(intForLoop).SetProperty "LoLimit", guPenPropertiesArray(intForLoop - 1).LoLimit
                    'JPB010203 port hj110802
                    objChart.pens.Item(intForLoop).AllowValueAxisReset = bAllowValueAxisReset
                End If
                
                objChart.pens.Item(intForLoop).SetProperty "FetchPenLimits", guPenPropertiesArray(intForLoop - 1).FetchPenLimits
                'mr050108 Ported from 4.0 thc050707 Set StartTime later
                'objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                
                'mr050108 Ported from 4.0 thc050707 Added 2 cases to this If to handle DateFixed/TimeRelative and DateRelative/TimeFixed
                If guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed And guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    'hj121807 Ported from 3.5 - hj022304 If the FixedDate is the system default date (12/30/1899), don't use it. Instead, use today's date.
                    'objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    ''hj121401 Need set these two properties if using FixedStartTime
                    'objChart.pens.Item(intForLoop).FixedDate = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    'objChart.pens.Item(intForLoop).FixedTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    If guPenPropertiesArray(intForLoop - 1).FixedDate = gdtSystemDefaultDate Then
                        'hj012308 Changed to not format the date and time
                        'objChart.pens.Item(intForLoop).StartTime = CStr(Format(dtToday, "Short Date") + " " + Format(guPenPropertiesArray(intForLoop - 1).FixedTime, "Short Time"))
                        objChart.pens.Item(intForLoop).StartTime = CStr(dtToday) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    Else
                        'hj012308 Changed to not format the date and time
                        'objChart.pens.Item(intForLoop).StartTime = CStr(Format(guPenPropertiesArray(intForLoop - 1).FixedDate, "Short Date") + " " + Format(guPenPropertiesArray(intForLoop - 1).FixedTime, "Short Time"))
                        objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    End If
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintRelative And guPenPropertiesArray(intForLoop - 1).StartDateType = gintRelative Then
                    objChart.pens.Item(intForLoop).StartTime = CGW_SetDate(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintRelative And guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    'mr050108 Ported from 4.0 hj121807 If the FixedDate is the system default date (12/30/1899), don't use it. Instead, use today's date.
                    If guPenPropertiesArray(intForLoop - 1).FixedDate = gdtSystemDefaultDate Then
                        objChart.pens.Item(intForLoop).StartTime = CGW_SetRelativeTime(dtToday, guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                    Else
                        objChart.pens.Item(intForLoop).StartTime = CGW_SetRelativeTime(guPenPropertiesArray(intForLoop - 1).FixedDate, guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                    End If
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed And guPenPropertiesArray(intForLoop - 1).StartDateType = gintRelative Then
                    objChart.pens.Item(intForLoop).StartTime = CStr(CGW_SetRelativeDate(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow)) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                End If
                
                'mr050108 Ported from 4.0 hj121807 Set these two properties if necessary
                If guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    objChart.pens.Item(intForLoop).FixedDate = objChart.pens.Item(intForLoop).StartTime
                End If
                If guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed Then
                    objChart.pens.Item(intForLoop).FixedTime = objChart.pens.Item(intForLoop).StartTime
                End If
               
                objChart.pens.Item(intForLoop).SetProperty "HistoricalSampleType", guPenPropertiesArray(intForLoop - 1).HistoricalSampleType
                objChart.pens.Item(intForLoop).SetProperty "MarkerChar", guPenPropertiesArray(intForLoop - 1).MarkerChar
                objChart.pens.Item(intForLoop).SetProperty "MarkerStyle", guPenPropertiesArray(intForLoop - 1).MarkerStyle
                objChart.pens.Item(intForLoop).SetProperty "PenLineColor", guPenPropertiesArray(intForLoop - 1).PenLineColor
                objChart.pens.Item(intForLoop).SetProperty "PenLineStyle", CLng(guPenPropertiesArray(intForLoop - 1).PenLineStyle)
                objChart.pens.Item(intForLoop).SetProperty "PenLineWidth", guPenPropertiesArray(intForLoop - 1).PenLineWidth
                objChart.pens.Item(intForLoop).SetProperty "StartDateType", guPenPropertiesArray(intForLoop - 1).StartDateType
                objChart.pens.Item(intForLoop).SetProperty "StartTimeType", guPenPropertiesArray(intForLoop - 1).StartTimeType
                'mr050108 Ported from 4.0 thc050707 TimeBeforeNow on chart object needs to be 24 hours or less.  So take the seconds
                ' associated with the DaysBeforeNow out.
                'objChart.pens.Item(intForLoop).SetProperty "TimeBeforeNow", guPenPropertiesArray(intForLoop - 1).TimeBeforeNow
                objChart.pens.Item(intForLoop).TimeBeforeNow = CGW_CalcRelativeTime(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                
                objChart.pens.Item(intForLoop).SetProperty "IntervalMilliSeconds", guPenPropertiesArray(intForLoop - 1).Interval Mod 1000
                objChart.pens.Item(intForLoop).SetProperty "Interval", (guPenPropertiesArray(intForLoop - 1).Interval - (guPenPropertiesArray(intForLoop - 1).Interval Mod 1000)) / 1000
                objChart.pens.Item(intForLoop).SetProperty "DisplayMilliseconds", guPenPropertiesArray(intForLoop - 1).DisplayMS
                objChart.pens.Item(intForLoop).SetProperty "DaylightSavingsTime", guPenPropertiesArray(intForLoop - 1).DaylightSavingsTime
                If guPenPropertiesArray(intForLoop - 1).TimeZone > 2 Then
                    'hj111204 Should set the property with Long type of data instead of Integer type
                    'objChart.pens.Item(intForLoop).SetProperty "TimeZoneBiasRelative", 3
                    'objChart.pens.Item(intForLoop).SetProperty "TimeZoneBiasExplicit", CInt(guPenPropertiesArray(intForLoop - 1).TimeZone - 3)
                    objChart.pens.Item(intForLoop).SetProperty "TimeZoneBiasRelative", CLng(3)
                    objChart.pens.Item(intForLoop).SetProperty "TimeZoneBiasExplicit", CLng(guPenPropertiesArray(intForLoop - 1).TimeZone - 3)
                Else
                    'hj111204 Should set the property with Long type of data instead of Integer type
                    'objChart.pens.Item(intForLoop).SetProperty "TimeZoneBiasRelative", CInt(guPenPropertiesArray(intForLoop - 1).TimeZone)
                    objChart.pens.Item(intForLoop).SetProperty "TimeZoneBiasRelative", CLng(guPenPropertiesArray(intForLoop - 1).TimeZone)
                End If
            Next intForLoop
   
        ElseIf UBound(guPenPropertiesArray) < objChart.pens.Count - 1 Then
            For intForLoop = 1 To UBound(guPenPropertiesArray) + 1
                'hj052704
                objChart.pens.Item(intForLoop).FetchPenLimits = guPenPropertiesArray(intForLoop - 1).FetchPenLimits
                'If source exists, use SetSource, if it DNE, use Source
                If guPenPropertiesArray(intForLoop - 1).Source <> "" Then
                    objChart.pens.Item(intForLoop).SetSource guPenPropertiesArray(intForLoop - 1).Source, True
                End If
                System.ValidateSource guPenPropertiesArray(intForLoop - 1).Source, lngStatus, objRetObject, strPropName
                If lngStatus <> 0 Then
                    objChart.pens.Item(intForLoop).legend.legendvalue = "****"
                    objChart.pens.Item(intForLoop).legend.legenddesc = ""
                    If objChart.pens.Item(intForLoop).CurrentValue <> 0 And objChart.pens.Item(intForLoop).legend.legendvalue <> "****" Then
                        objChart.pens.Item(intForLoop).CurrentValue = 0
                    End If
                End If
                objChart.pens.Item(intForLoop).DaysBeforeNow = guPenPropertiesArray(intForLoop - 1).DaysBeforeNow
                objChart.pens.Item(intForLoop).Duration = guPenPropertiesArray(intForLoop - 1).Duration
                
                'thc070907  1-208005699 Overwrite hi & lo limit with values from file only if FetchPenLimits is False
                'If FetchPenLimits is True, the limits were fetched earlier when SetSource was called.
                If guPenPropertiesArray(intForLoop - 1).FetchPenLimits = False Then
                    'JPB010203 port hj110802 Use AllowValueAxisReset to make sure we can set the high and low limits into the pen
                    bAllowValueAxisReset = objChart.pens.Item(intForLoop).AllowValueAxisReset
                    objChart.pens.Item(intForLoop).AllowValueAxisReset = True
                    objChart.pens.Item(intForLoop).HiLimit = guPenPropertiesArray(intForLoop - 1).HiLimit
                    objChart.pens.Item(intForLoop).LoLimit = guPenPropertiesArray(intForLoop - 1).LoLimit
                    'JPB010203 port hj110802
                    objChart.pens.Item(intForLoop).AllowValueAxisReset = bAllowValueAxisReset
                End If
                
                objChart.pens.Item(intForLoop).FetchPenLimits = guPenPropertiesArray(intForLoop - 1).FetchPenLimits
                ' thc050707 Set StartTime later
                'objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                objChart.pens.Item(intForLoop).HistoricalSampleType = guPenPropertiesArray(intForLoop - 1).HistoricalSampleType
                objChart.pens.Item(intForLoop).MarkerChar = guPenPropertiesArray(intForLoop - 1).MarkerChar
                objChart.pens.Item(intForLoop).MarkerStyle = guPenPropertiesArray(intForLoop - 1).MarkerStyle
                objChart.pens.Item(intForLoop).PenLineColor = guPenPropertiesArray(intForLoop - 1).PenLineColor
                objChart.pens.Item(intForLoop).PenLineStyle = CLng(guPenPropertiesArray(intForLoop - 1).PenLineStyle)
                objChart.pens.Item(intForLoop).PenLineWidth = guPenPropertiesArray(intForLoop - 1).PenLineWidth
                objChart.pens.Item(intForLoop).StartDateType = guPenPropertiesArray(intForLoop - 1).StartDateType
                objChart.pens.Item(intForLoop).StartTimeType = guPenPropertiesArray(intForLoop - 1).StartTimeType
                
                'thc050707 Added 2 cases to this If to handle DateFixed/TimeRelative and DateRelative/TimeFixed
                If guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed And guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    'hj121807 Ported from 3.5 - hj022304 If the FixedDate is the system default date (12/30/1899), don't use it. Instead, use today's date.
                    'objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    ''hj121401 Need set these two properties if using FixedStartTime
                    'objChart.pens.Item(intForLoop).FixedDate = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    'objChart.pens.Item(intForLoop).FixedTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    If guPenPropertiesArray(intForLoop - 1).FixedDate = gdtSystemDefaultDate Then
                        'hj012308 Changed to not format the date and time
                        'objChart.pens.Item(intForLoop).StartTime = CStr(Format(dtToday, "Short Date") + " " + Format(guPenPropertiesArray(intForLoop - 1).FixedTime, "Short Time"))
                        objChart.pens.Item(intForLoop).StartTime = CStr(dtToday) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    Else
                        'hj012308 Changed to not format the date and time
                        'objChart.pens.Item(intForLoop).StartTime = CStr(Format(guPenPropertiesArray(intForLoop - 1).FixedDate, "Short Date") + " " + Format(guPenPropertiesArray(intForLoop - 1).FixedTime, "Short Time"))
                        objChart.pens.Item(intForLoop).StartTime = CStr(guPenPropertiesArray(intForLoop - 1).FixedDate) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                    End If
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintRelative And guPenPropertiesArray(intForLoop - 1).StartDateType = gintRelative Then
                    objChart.pens.Item(intForLoop).StartTime = CGW_SetDate(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintRelative And guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    'hj121807 If the FixedDate is the system default date (12/30/1899), don't use it. Instead, use today's date.
                    If guPenPropertiesArray(intForLoop - 1).FixedDate = gdtSystemDefaultDate Then
                        objChart.pens.Item(intForLoop).StartTime = CGW_SetRelativeTime(dtToday, guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                    Else
                        objChart.pens.Item(intForLoop).StartTime = CGW_SetRelativeTime(guPenPropertiesArray(intForLoop - 1).FixedDate, guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                    End If
                ElseIf guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed And guPenPropertiesArray(intForLoop - 1).StartDateType = gintRelative Then
                    objChart.pens.Item(intForLoop).StartTime = CStr(CGW_SetRelativeDate(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow)) + " " + CStr(guPenPropertiesArray(intForLoop - 1).FixedTime)
                End If
                
                'hj121807 Set these two properties if necessary
                If guPenPropertiesArray(intForLoop - 1).StartDateType = gintFixed Then
                    objChart.pens.Item(intForLoop).FixedDate = objChart.pens.Item(intForLoop).StartTime
                End If
                If guPenPropertiesArray(intForLoop - 1).StartTimeType = gintFixed Then
                    objChart.pens.Item(intForLoop).FixedTime = objChart.pens.Item(intForLoop).StartTime
                End If
               
                 'thc050707 TimeBeforeNow on chart object needs to be 24 hours or less.  So take the seconds
                ' associated with the DaysBeforeNow out.
                'objChart.pens.Item(intForLoop).TimeBeforeNow = guPenPropertiesArray(intForLoop - 1).TimeBeforeNow
                objChart.pens.Item(intForLoop).TimeBeforeNow = CGW_CalcRelativeTime(guPenPropertiesArray(intForLoop - 1).DaysBeforeNow, guPenPropertiesArray(intForLoop - 1).TimeBeforeNow)
                
                objChart.pens.Item(intForLoop).IntervalMilliseconds = guPenPropertiesArray(intForLoop - 1).Interval Mod 1000
                objChart.pens.Item(intForLoop).Interval = (guPenPropertiesArray(intForLoop - 1).Interval - objChart.pens.Item(intForLoop).IntervalMilliseconds) / 1000
                objChart.pens.Item(intForLoop).DisplayMilliseconds = guPenPropertiesArray(intForLoop - 1).DisplayMS
                objChart.pens.Item(intForLoop).DaylightSavingsTime = guPenPropertiesArray(intForLoop - 1).DaylightSavingsTime
                If guPenPropertiesArray(intForLoop - 1).TimeZone > 2 Then
                    'hj111204 Should set the property with Long type of data instead of Integer type
                    'objChart.pens.Item(intForLoop).TimeZoneBiasRelative = 3
                    'objChart.pens.Item(intForLoop).TimeZoneBiasExplicit = guPenPropertiesArray(intForLoop - 1).TimeZone - 3
                    objChart.pens.Item(intForLoop).TimeZoneBiasRelative = CLng(3)
                    objChart.pens.Item(intForLoop).TimeZoneBiasExplicit = CLng(guPenPropertiesArray(intForLoop - 1).TimeZone - 3)
                Else
                    'hj111204 Should set the property with Long type of data instead of Integer type
                    'objChart.pens.Item(intForLoop).TimeZoneBiasRelative = guPenPropertiesArray(intForLoop - 1).TimeZone
                    objChart.pens.Item(intForLoop).TimeZoneBiasRelative = CLng(guPenPropertiesArray(intForLoop - 1).TimeZone)
                End If
            Next intForLoop
            'delete the pens that are not needed anymore.
            For intForLoop = (UBound(guPenPropertiesArray) + 2) To objChart.pens.Count
                objChart.deletepen UBound(guPenPropertiesArray) + 2
            Next intForLoop
        
        End If
   
        objChart.RefreshChartData
        objChart.Refresh
    End If

    For Each objSubChart In objChart.ContainedObjects
      CGW_ApplyChartGroupSettings objSubChart
    Next objSubChart
    blnErrorDisplayed = False
    Exit Sub
    
ErrorHandler:
    If Err.Number = -2147200604 Then
    'User does not have ThisNode enabled, give a message
    'this error is for the BETA release, and can be removed for the Final Release.
        If Not blnErrorDisplayed Then
            strError = objStrMgr.GetNLSStr(CLng(NLS_INVALIDPEN1)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_INVALIDPEN2)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_INVALIDPEN3))
            strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
            MsgBox strError, , strTitle
            blnErrorDisplayed = True
        End If
        Resume Next
    Else
        HandleError
    End If
End Sub


Public Sub CGW_OpenChartGroupFileSettings()
    Dim intLocalIndex As Integer
  
    On Error GoTo ErrorHandler
    frmChartGroupForm.lstDataSource.Clear
    
    For intLocalIndex = 0 To UBound(guPenPropertiesArray)
          frmChartGroupForm.lstDataSource.AddItem (guPenPropertiesArray(intLocalIndex).Source)
    Next intLocalIndex
        
    If UBound(guPenPropertiesArray) >= 0 Then
        frmChartGroupForm.ctrlColorButton.Color = guPenPropertiesArray(0).PenLineColor
        frmChartGroupForm.txtDuration = CGW_SetDuration(guPenPropertiesArray(0).Duration)
        frmChartGroupForm.chkFetchPenLimits = guPenPropertiesArray(0).FetchPenLimits
        'jes clarify #235605
        'frmChartGroupForm.txtFixedStartTime = Format(guPenPropertiesArray(0).FixedDate, "mm/dd/yy") + " " + Format(guPenPropertiesArray(0).FixedTime, "hh:mm:ss AMPM")
        'hj121807 Ported from 3.5 - hj022304 If the FixedDate is the system default date (12/30/1899), don't use it, i.e. only use the FixedTime in the FixedStartTime box
        If guPenPropertiesArray(0).FixedDate = gdtSystemDefaultDate Then
            frmChartGroupForm.txtFixedStartTime = Format(guPenPropertiesArray(0).FixedTime, "Short Time")
        Else
            frmChartGroupForm.txtFixedStartTime = Format(guPenPropertiesArray(0).FixedDate, "Short Date") + " " + Format(guPenPropertiesArray(0).FixedTime, "Short Time")
        End If
        frmChartGroupForm.txtHiLimit = Format(guPenPropertiesArray(0).HiLimit, "##0.00")
            
        If guPenPropertiesArray(0).StartDateType = gintRelative Then
            frmChartGroupForm.optTimeBeforeNow = True
        ElseIf guPenPropertiesArray(0).StartDateType = gintFixed Then
            frmChartGroupForm.optFixedStartTime = True
        End If
            
        If frmChartGroupForm.optTimeBeforeNow = True Then
            'although it is selected, see if it is enabled
            If frmChartGroupForm.optTimeBeforeNow.Enabled = True Then
                frmChartGroupForm.cboTimeBeforeNow.Enabled = True
            Else
                frmChartGroupForm.cboTimeBeforeNow.Enabled = False
            End If
            frmChartGroupForm.txtFixedStartTime.Enabled = False
        ElseIf frmChartGroupForm.optFixedStartTime.Enabled = True Then
            'although it is selected, see if it is enabled
            If frmChartGroupForm.optTimeBeforeNow.Enabled = True Then
                frmChartGroupForm.txtFixedStartTime.Enabled = True
            Else
                frmChartGroupForm.txtFixedStartTime.Enabled = False
            End If
            frmChartGroupForm.cboTimeBeforeNow.Enabled = False
        End If
        frmChartGroupForm.chkApplyAllPens.Value = guPenPropertiesArray(0).ApplyToAll
        frmChartGroupForm.cboHistoricalMode.Text = frmChartGroupForm.cboHistoricalMode.List(guPenPropertiesArray(0).HistoricalSampleType)
        frmChartGroupForm.txtLowLimit = Format(guPenPropertiesArray(0).LoLimit, "##0.00")
            
        If guPenPropertiesArray(0).DaysBeforeNow > 99 Then
            frmChartGroupForm.cboTimeBeforeNow = CStr(guPenPropertiesArray(0).DaysBeforeNow) + ":" + CGW_SetTimeBeforeNow(guPenPropertiesArray(0).TimeBeforeNow)
        ElseIf guPenPropertiesArray(0).DaysBeforeNow <= 99 And guPenPropertiesArray(0).DaysBeforeNow > 9 Then
            frmChartGroupForm.cboTimeBeforeNow = "0" + CStr(guPenPropertiesArray(0).DaysBeforeNow) + ":" + CGW_SetTimeBeforeNow(guPenPropertiesArray(0).TimeBeforeNow)
        ElseIf guPenPropertiesArray(0).DaysBeforeNow <= 9 Then
            frmChartGroupForm.cboTimeBeforeNow = "00" + CStr(guPenPropertiesArray(0).DaysBeforeNow) + ":" + CGW_SetTimeBeforeNow(guPenPropertiesArray(0).TimeBeforeNow)
        End If
            
        frmChartGroupForm.txtMarkerChar = guPenPropertiesArray(0).MarkerChar
        frmChartGroupForm.cboMarkerStyle.Text = frmChartGroupForm.cboMarkerStyle.List(guPenPropertiesArray(0).MarkerStyle)
        frmChartGroupForm.txtLineWidth = guPenPropertiesArray(0).PenLineWidth
        frmChartGroupForm.cboLineStyle.Text = frmChartGroupForm.cboLineStyle.List(guPenPropertiesArray(0).PenLineStyle)
        frmChartGroupForm.txtInterval = CGW_SetInterval(guPenPropertiesArray(0).Interval)
        frmChartGroupForm.chkDisplayMS.Value = guPenPropertiesArray(0).DisplayMS
        frmChartGroupForm.chkAdjustForDST.Value = guPenPropertiesArray(0).DaylightSavingsTime
        frmChartGroupForm.cboTimeZone.Text = frmChartGroupForm.cboTimeZone.List(guPenPropertiesArray(0).TimeZone)
    End If
        
    frmChartGroupForm.lstDataSource.AddItem ""
    frmChartGroupForm.lstDataSource.Selected(0) = True
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub


Public Sub CGW_Buttons()
    On Error GoTo ErrorHandler
    ' hj072903
    ' Is this script running in the workspace or background?
    If TypeName(Application) = "CFixApp" Then
        ' running in the workspace
        Set AppObj = Application
    Else
        ' running in the background
            
        ' see if we can get the workspace object
        Set AppObj = GetObject(, "Workspace.Application")
        
        If AppObj Is Nothing Then
            Exit Sub
        End If
    End If
    'If in Configure mode then enable the Save and Save as buttons, disable the Apply button
    If AppObj.Mode = 1 Then ' hj072903
        frmChartGroupForm.cmdSaveAs.Enabled = True
        frmChartGroupForm.cmdApply.Enabled = False
        frmChartGroupForm.cmdSave.Enabled = True
    'If in Run Mode, all 3 buttons should be enabled
    ElseIf AppObj.Mode = 4 Then ' hj072903
        frmChartGroupForm.cmdSaveAs.Enabled = True
        frmChartGroupForm.cmdSave.Enabled = True
        frmChartGroupForm.cmdApply.Enabled = True
    End If
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub

'***************************************************************************************
'Function: CGW_SetDate(intDays As Integer, lngTime As Long) As Variant
'
'Purpose:   This function accepts days before now from the
'           guPenPropertiesArray(intForLoop).DaysBeforeNow and it accepts time before now from the
'           guPenPropertiesArray(intForLoop).TimeBeforeNow.  Then it converts this into a date format
'           This applied to chart object's StartTime property.
'
'Inputs:    intDays:  Gets info from guPenProperteisArray(intForLoop).DaysBeforeNow
'           lngTime:  Gets info from guPenPropertiesArray(intForLoop).TimeBeforeNow
'***************************************************************************************

Function CGW_SetDate(intDays As Integer, lngTime As Long) As Variant
    Dim vntTimeDateBeforeNow As Variant
    Dim RelTime As Long
    
    On Error GoTo ErrorHandler
    
    'mr050108 Ported from 4.0 thc050707 Need to use just the time portion of TimeBeforeNow
    RelTime = CGW_CalcRelativeTime(intDays, lngTime)
    
    'Use the DateAdd Function to calculate the exact date and time by subracting
    vntTimeDateBeforeNow = DateAdd("s", CDbl(-RelTime), Now)
    vntTimeDateBeforeNow = DateAdd("d", CDbl(-intDays), vntTimeDateBeforeNow)
    CGW_SetDate = vntTimeDateBeforeNow
    Exit Function
    
ErrorHandler:
    HandleError
End Function
'***************************************************************************************
'Function: CGW_SetRelativeDate(intDays As Integer) As Variant
'
'Purpose:   This function accepts days before now from the
'           guPenPropertiesArray(intForLoop).DaysBeforeNow. Then it converts this into a date format.
'
'Inputs:    intDays:  Gets info from guPenProperteisArray(intForLoop).DaysBeforeNow
'***************************************************************************************

Function CGW_SetRelativeDate(intDays As Integer) As Variant
    Dim vntTimeDate As Variant
    On Error GoTo ErrorHandler
    'Use the DateAdd Function to calculate the exact date and time by subracting
    vntTimeDate = DateAdd("d", CDbl(-intDays), Date)
    CGW_SetRelativeDate = vntTimeDate
    Exit Function
    
ErrorHandler:
    HandleError
End Function

'***************************************************************************************
'Function: CGW_SetRelativeTime (lngTime As Long) As Variant
'
'Purpose:   This function accepts time before now from the
'           guPenPropertiesArray(intForLoop).TimeBeforeNow.  Then it converts this into a date format
'
'Inputs:    lngTime:  Gets info from guPenPropertiesArray(intForLoop).TimeBeforeNow
'
'***************************************************************************************

Function CGW_SetRelativeTime(FixedDate As Date, intDays As Integer, lngTime As Long) As Variant
    Dim vntTimeDate As Variant
    Dim RelTime As Long
    Dim strFixedDateNow As String
        
    On Error GoTo ErrorHandler
    
    strFixedDateNow = CStr(Format(FixedDate, "Short Date")) + " " + CStr(Format(time, "Long Time"))
    vntTimeDate = CDate(strFixedDateNow)
    RelTime = CGW_CalcRelativeTime(intDays, lngTime)
    CGW_SetRelativeTime = DateAdd("s", CDbl(-RelTime), vntTimeDate)
    
    Exit Function
    
ErrorHandler:
    HandleError
End Function

'***************************************************************************************
'Function: CGW_CalcRelativeTime(intDays As Integer, lngTime As Long) As Long
'
'Purpose:   This function accepts days before now from the
'           guPenPropertiesArray(intForLoop).DaysBeforeNow and it accepts time before now from the
'           guPenPropertiesArray(intForLoop).TimeBeforeNow.  Then it calculates the relative time
'           before now (removes the seconds associated with the DaysBeforeNow).
'
'Inputs:    intDays:  Gets info from guPenProperteisArray(intForLoop).DaysBeforeNow
'           lngTime:  Gets info from guPenPropertiesArray(intForLoop).TimeBeforeNow
'***************************************************************************************

Function CGW_CalcRelativeTime(intDays As Integer, lngTime As Long) As Long

    Dim DaySeconds As Long
    
    On Error GoTo ErrorHandler
    DaySeconds = intDays * 86400

    If lngTime >= DaySeconds Then
        CGW_CalcRelativeTime = lngTime - DaySeconds
    Else
        CGW_CalcRelativeTime = lngTime
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError
End Function
'*************************************************************************************
'
'Sub Routine CGW_ChartGroupFileVersion (strFileName As String)
'
'Purpose to detect chart group version from the file
'
'Inputs:   strFileName: contains full path of the file to be read from the disk
'
'*************************************************************************************

Public Sub CGW_ChartGroupFileVersion(strFileName As String)

    Dim intFileHandle As String
    Dim strPropertiesNames(22) As Variant
    Dim strError As String
    Dim strTitle As String
    
    On Error GoTo ErrorHandler
    
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    intFileHandle = FreeFile
    Open strFileName For Input As #intFileHandle
    Input #1, strPropertiesNames(0), strPropertiesNames(1), strPropertiesNames(2), strPropertiesNames(3), strPropertiesNames(4), strPropertiesNames(5), strPropertiesNames(6), strPropertiesNames(7), strPropertiesNames(8), strPropertiesNames(9), strPropertiesNames(10), strPropertiesNames(11), strPropertiesNames(12), strPropertiesNames(13), strPropertiesNames(14), strPropertiesNames(15), strPropertiesNames(16), strPropertiesNames(17), strPropertiesNames(18), strPropertiesNames(19), strPropertiesNames(20), strPropertiesNames(21), strPropertiesNames(22)
    If strPropertiesNames(19) = "&&" Then
        giVer = 0
    Else
        giVer = 1
    End If
    Close #intFileHandle
    
    Exit Sub
ErrorHandler:
    gblnBadFileName = True
    If Err.Number = 76 Then
        strError = objStrMgr.GetNLSStr(CLng(NLS_CHARTDNE))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
        MsgBox strError, , strTitle
    ElseIf Err.Number = 65 Then
        strError = objStrMgr.GetNLSStr(CLng(NLS_BADFORMAT1)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_BADFORMAT2))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
        MsgBox strError, vbExclamation, strTitle
    Else
        HandleError
    End If
    Close #intFileHandle

End Sub

Public Function CGW_SetInterval(lngInterval As Long) As String
    Dim strHours As String
    Dim strMinutes As String
    Dim strSeconds As String
    Dim strMilliSeconds As String
    Dim strTempSeconds As String
    Dim lngx As Long
    Dim lngy As Long
    Dim bHist As Boolean
    
    On Error GoTo ErrorHandler
    lngy = 60
    lngx = 1000
    bHist = CGW_IsIHistorianMode
    
    strHours = (lngInterval - lngInterval Mod (lngx * lngy * lngy)) / (lngx * lngy * lngy)
    If Len(strHours) = 1 Then
      strHours = "0" + strHours
    End If
    strMinutes = ((lngInterval Mod (lngx * lngy * lngy)) - ((lngInterval Mod (lngx * lngy * lngy)) Mod (lngx * lngy))) / (lngx * lngy)
    If Len(strMinutes) = 1 Then
      strMinutes = "0" + strMinutes
    End If
    strTempSeconds = ((lngInterval Mod (lngx * lngy * lngy)) Mod (lngx * lngy))
    strSeconds = (strTempSeconds - (strTempSeconds Mod lngx)) / lngx
    If Len(strSeconds) = 1 Then
      strSeconds = "0" + strSeconds
    End If
    strMilliSeconds = strTempSeconds Mod lngx
    If Len(strMilliSeconds) = 1 Then
      strMilliSeconds = "00" + strMilliSeconds
    ElseIf Len(strMilliSeconds) = 2 Then
      strMilliSeconds = "0" + strMilliSeconds
    End If
    If bHist = True Then
        CGW_SetInterval = Trim(strHours) + ":" + Trim(strMinutes) + ":" + Trim(strSeconds) + ":" + Trim(strMilliSeconds)
    Else
        CGW_SetInterval = Trim(strHours) + ":" + Trim(strMinutes) + ":" + Trim(strSeconds)
    End If
    Exit Function
    
ErrorHandler:
    HandleError
End Function

Public Function CGW_GetInterval(strInterval As String) As Long
    Dim strHours As String
    Dim strMinutes As String
    Dim strSeconds As String
    Dim strMilliSeconds As String
    Dim strChecktxtInterval As String
    Dim lngMSInaSecond As Long
    Dim lngSecondsInaMinute As Long
    Dim strError As String
    Dim strTitle As String
    Dim bHist As Boolean
    
    lngSecondsInaMinute = 60
    lngMSInaSecond = 1000
    bHist = False
    bHist = CGW_IsIHistorianMode
    If bHist = True Then
        On Error GoTo ErrorHandler
    Else
        On Error GoTo ErrorHandler2
    End If
    
    If bHist = True Then
        strInterval = Trim(strInterval)
        If Len(strInterval) <> 12 Then
            GoTo ErrorHandler
        End If

        strHours = Left(strInterval, 2)
        If (Not IsNumeric(strHours)) Or (CLng(strHours) >= 24) Or (CLng(strHours) < 0) Then
            GoTo ErrorHandler
        End If

        strChecktxtInterval = Left(Right(strInterval, Len(strInterval) - 2), 1)
        If strChecktxtInterval <> ":" Then
            GoTo ErrorHandler
        End If

        strMinutes = Left(Right(strInterval, Len(strInterval) - 3), 2)
        If (Not IsNumeric(strMinutes)) Or (CLng(strMinutes) >= 60) Or (CLng(strMinutes) < 0) Then
            GoTo ErrorHandler
        End If

        strChecktxtInterval = Left(Right(strInterval, Len(strInterval) - 5), 1)
        If strChecktxtInterval <> ":" Then
            GoTo ErrorHandler
        End If

        strSeconds = Left(Right(strInterval, Len(strInterval) - 6), 2)
        If (Not IsNumeric(strSeconds)) Or (CLng(strSeconds) >= 60) Or (CLng(strSeconds) < 0) Then
            GoTo ErrorHandler
        End If

        strChecktxtInterval = Left(Right(strInterval, Len(strInterval) - 8), 1)
        If strChecktxtInterval <> ":" Then
            GoTo ErrorHandler
        End If

        strMilliSeconds = Left(Right(strInterval, Len(strInterval) - 9), 3)
        If (Not IsNumeric(strMilliSeconds)) Or (CLng(strMilliSeconds) >= 1000) Or (CLng(strMilliSeconds) < 0) Then
            GoTo ErrorHandler
        End If
    
        CGW_GetInterval = CLng(strHours) * lngSecondsInaMinute * lngSecondsInaMinute * lngMSInaSecond + CLng(strMinutes) * lngSecondsInaMinute * lngMSInaSecond + CLng(strSeconds) * lngMSInaSecond + CLng(strMilliSeconds)
    Else 'bHist = False
        strInterval = Trim(strInterval)
        If Len(strInterval) <> 8 Then
            GoTo ErrorHandler2
        End If

        strHours = Left(strInterval, 2)
        If (Not IsNumeric(strHours)) Or (CLng(strHours) >= 24) Or (CLng(strHours) < 0) Then
            GoTo ErrorHandler2
        End If

        strChecktxtInterval = Left(Right(strInterval, Len(strInterval) - 2), 1)
        If strChecktxtInterval <> ":" Then
            GoTo ErrorHandler2
        End If

        strMinutes = Left(Right(strInterval, Len(strInterval) - 3), 2)
        If (Not IsNumeric(strMinutes)) Or (CLng(strMinutes) >= 60) Or (CLng(strMinutes) < 0) Then
            GoTo ErrorHandler2
        End If

        strChecktxtInterval = Left(Right(strInterval, Len(strInterval) - 5), 1)
        If strChecktxtInterval <> ":" Then
            GoTo ErrorHandler2
        End If

        strSeconds = Left(Right(strInterval, Len(strInterval) - 6), 2)
        If (Not IsNumeric(strSeconds)) Or (CLng(strSeconds) >= 60) Or (CLng(strSeconds) < 0) Then
            GoTo ErrorHandler2
        End If
        
        CGW_GetInterval = CLng(strHours) * lngSecondsInaMinute * lngSecondsInaMinute * lngMSInaSecond + CLng(strMinutes) * lngSecondsInaMinute * lngMSInaSecond + CLng(strSeconds) * lngMSInaSecond
    End If 'endif bHist = True
    Exit Function

ErrorHandler:
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    strError = objStrMgr.GetNLSStr(CLng(NLS_INTERVALERROR))
    strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
    MsgBox strError, , strTitle
    frmChartGroupForm.txtInterval = "00:00:10:000"
    Exit Function
ErrorHandler2:
    Set objStrMgr = CreateObject("iFix_CGW.ResMgr")
    strError = objStrMgr.GetNLSStr(CLng(NLS_INTERVALERROR2))
    strTitle = objStrMgr.GetNLSStr(CLng(NLS_TITLE))
    MsgBox strError, , strTitle
    frmChartGroupForm.txtInterval = "00:00:10"
End Function
Public Function CGW_IsIHistorianMode() As Boolean
    Dim lngSize As Long
    Dim intPos As Integer
    Dim strReturnedString As String
        
    ' read the setting from the INI file
    ' note we use toolbar path to find INI file
    ' but that is always LOCPATH
    lngSize = 100
    strReturnedString = Space(lngSize)
    GetPrivateProfileString HISTORIAN_INI_SECTION, HISTORIAN_INI_ENTRY, CLASSIC_INI_VAL, strReturnedString, lngSize&, System.ToolbarPath + "\FixUserPreferences.ini"
    
    ' clean up the value
    LTrim (strReturnedString)
    RTrim (strReturnedString)
    intPos = InStr(strReturnedString, Chr(0))
    strReturnedString = Left(strReturnedString, intPos - 1)
    
    ' compare the value and set the mode
    If UCase(strReturnedString) = UCase(IHISTORIAN_INI_VAL) Then
        CGW_IsIHistorianMode = True
    Else
        ' treat anything else as classic
        CGW_IsIHistorianMode = False
    End If
End Function

Public Function CGW_HasEscKeySent() As Boolean
    CGW_HasEscKeySent = bHasSentEsc
End Function

Public Sub CGW_EscKeyHasBeenSent(bSetFlag As Boolean)
    bHasSentEsc = bSetFlag
End Sub
'ab03172007 Windows Vista does not support SendKeys
Public Function SendKeysA(ByVal vKey As Integer, Optional booDown As Boolean = False)
Dim GInput(0) As GENERALINPUT
Dim KInput As KEYBDINPUT
KInput.wVk = vKey
If Not booDown Then
   KInput.dwFlags = KEYEVENTF_KEYUP
End If
GInput(0).dwType = INPUT_KEYBOARD
CopyMemory GInput(0).xi(0), KInput, Len(KInput)
Call SendInput(1, GInput(0), Len(GInput(0)))
End Function


