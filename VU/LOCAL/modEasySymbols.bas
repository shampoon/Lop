Attribute VB_Name = "modEasySymbols"
Option Explicit
' Version 2.1 for iFIX 2.1 ---- 07/09/99

'MOD LOG
'
'Ver    Date        By      Bug #       Notes
'----   --------    ---     --------    ---------------------------------------------------
'3.5    02/07/03    JPB                 Using Tracker database "iFIX 3.5"
'       03/14/03    JPB     T865        Added IsAlphanumericStartIsAlpha() function and call it from
'                                       BDW_UpdateGroupName() to properly validate dynamo name.
'       03/14/03    JPB     T18         Added new line to error message when invalid dynamo name is entered.
'       03/14/03    JPB     T422        Added tooltips for user prompt labels that are greater than 30
'                                       characters, since there isn't enough room to display long strings
'                                       in the Edit Dynamo form.
'       03/17/03    JPB     T18         Added new line of error message (above) to BDW_HasNoConnections()
'       04/24/03    JPB     T1328       Check if the object type is FixGlobalSysInfo (as is the case when
'                                       the object is a date or time link) before checking it's category,
'                                       since that object doesn't support the category property.  Take the same
'                                       actions as would be taken if category was not equal to "Animation" since
'                                       Time and Date links are not animations (though their text is animated).
'       02/05/04    hj      C283379     Added BDW_SubstituteNodeTag() and modified BDW_FillInEditDynamoArrayFromForm()
'                                       to make multiple occurences of partial substitution in a single property work.
'4.0    09/27/2005  PBH     T1529       Integrated SIM 290242
'    >> 05/25/04    hj      C290242     Modified BDW_AddToSymbolList() to handle decimal points in the expression.
'4.5    10/20/2006  PBH                 Modified CreateDynamo() to create a Dynamo object and not a group.
'       11/09/2006  PBH                 Modified Create & Edit code & forms to handle Dynamo object and it's new Text_Name property
'       01/03/2007  PBH                 Changed Text_Name to Dynamo_Description
'4.5    06/11/2008  jes     1-441590991 Changed labels to textboxes on form for current setting (frx changes)
'                                       the labels clip anything after a dash or underscore, when the length is exceeded.
'5.0    05/07/2008  Priya   T5441       Port cmk 1-99874598  only display prompt for *.blink if there is a datasource associated with it
'       07/25/2008  kei     T6617       Support new charts
'       04/01/2009  kei     T7133       Button layouts for Build Dynamo when contained objects have no connection
'5.1    08/25/2009  ab                  If iHistorian, let the Historian tab show in expression editor
'All string constants used in this module
Const mLOOKUP_CLASS = "Lookup"
Const mTEXT_CLASS = "Text"
Const mLINEAR_CLASS = "Linear"
Const mFORMAT_CLASS = "Format"
Const mCHART_CLASS = "Chart"
Const mVARIABLE_CLASS = "Variable"
Const mFIXEVENT_CLASS = "FixEvent"
Const mPEN_CLASS = "Pen"
Const mOLEOBJECT_CLASS = "OleObject"
'kei072508 iFix5.0 T6617
Const mREALTIMEDS_CLASS = "RealTimeDataSet"
Const mSPCDS_CLASS = "RealTimeSPCDataSet"
Const mHISTDS_CLASS = "HistoricalDataSet"

Const mMAX_SYMBOL_SIZE = 200
Const mMAX_PROMPT_SIZE = 200
Const mNO_PROMPT = "{[No prompt defined]}"

'The Error Strings
Const mERROR1 = 2800 '"No objects selected"
Const NLS_mERROR1 = 2080   'The 5 error messages have to do with Illegal Dynamo Name
Const NLS_mERROR1a = 2081
Const NLS_mERROR1b = 2082
Const NLS_mERROR1c = 2083
Const NLS_mERROR1d = 2084
Const NLS_mTITLE = 2090     'Title of the message box
Const NLS_mPrompt = 2091     'kei090302 #257738 - Prompt message
Const mERROR2 = 2085 'cannot use a single OLE Object
Const NLS_mSTARTCHARS = 11113   'JPB031403 T865 added for dynamo name validation
Const NLS_mRESTOFCHARS = 11114  'JPB031403 T865 added for dynamo name validation
Const NLS_mERROR1e = 2086       'JPB031403 T18  added another reason to 'bad name' error
Const NLS_mDYN_DESC_LEN_ERROR = 11116 ' PBH 11/06/2006 - Error message for when the text string is too long

Public gblnNoSource As Boolean
Public gintNumberInArray As Integer 'Used in the Edit Dynamo Form
Public gintNumberOfUniquePrompts As Integer
Public gintTempNumberOfUniquePrompts As Integer

'Array to temporarily hold any modified information in the CreateDynamo form
Private BDW_CreateDynamoArray
'Module level variable to hold the number of indeces in the BDW_CreateDynamoArray
Private BDW_mintCDANumOfIndeces As Integer
'Array to temporarily hold any modified information in the EditDynamo form
Private BDW_EditDynamoArray
'Module level variable to hold the number of indeces in the BDW_EditDynamoArray
Private BDW_mintEDANumOfIndeces As Integer

Type SYMBOL_LIST
    strName As String                     'The Name of the Connection
    strFullName As String                 'The Fully Qualified Name of the Connection
    strPrompt As String                   'The User Prompt, surrounded by {}
    strContent As String                  'The Node.Tag, Caption, or Initial Value
    strField As String                    'The field *if any
    blnIsUsingSubstitution As Boolean     'Flag for if the user is using partial substitution
    strPropertyName As String             'The Property it is connected to.
End Type

'DataType array to hold all of the information in creating the Dynamo
Public BDW_gudtSymbolData(mMAX_SYMBOL_SIZE) As SYMBOL_LIST 'Array of the SYMBOL_LIST data type.
Private mintSymbolLines As Integer          'number of indeces in BDW_gudtSymbolData

Type PROMPT_LIST
    strPrompt As String                    'The User Prompt
    strContent As String                   'The Node.Tag, Caption, or Initial Value
    strField As String                     'The field *if any
    blnIsUsingSubstitution As Boolean      'Flag for if the user is using partial substitution
End Type

Type TestArray
    strProperty As String
    strSource As String
End Type

Private MyTestArray() As TestArray  'Array For several Obj to Obj Connections on One Object
'DataType to hold all of the information in editing the Dynamo
Private BDW_mudtPromptData(mMAX_PROMPT_SIZE) As PROMPT_LIST 'Array of the PROMPT_LIST data type.
Private mintPromptLines As Integer          'number of indeces in BDW_mudtPromptData

'NLS Object
Private objStrMgr As Object
'Used in BDW_HasNoConnections
Private mstrBadName As String    'The illegal name the Developer entered
Dim mblnError As Boolean     'Set to true if the Developer enters an Illegal Name
'The name of the object that will become a dynamo
Public mobjParentObject As Object
Public mstrParentName As String
'Boolean that is set to true if a new group is created.
Private mblnNewGroupCreated As Boolean
'Public variable for creating an instance of the frmEditDynamo
Public NewEditDynamoForm As frmEditDynamo
Public mstrPictureName As String
Private mblnGroupAlreadyDone As Boolean
Private Const HISTORIAN_INI_SECTION$ = "Historian"
Private Const HISTORIAN_INI_ENTRY$ = "CurrentHistorian"
' these are the possible values for current historian
Private Const IHISTORIAN_INI_VAL$ = "iHistorian"
Private Const CLASSIC_INI_VAL$ = "Classic"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub GetFormEditDynamo()
'******************************************************************************************
'PURPOSE: To create an instance of the EditDynamo form so it can be accessed
'from User Globals by other projects.
'******************************************************************************************

    Set NewEditDynamoForm = New frmEditDynamo
End Sub

Function BDW_GetInitValueOfVar(objMainObject As Object, strConnectionProperty As String, strVariableName As String) As String
'******************************************************************************************
'PURPOSE: To search an object, and all of its contained objects, and check if it already has an EasyDynamo Variable attached.
'If it does, BDW_GetInitValueOfVar retrieves the User Prompt from the variable's Initial Value.
'INPUTS:
'   objMainObject:  The Parent Object we are searching
'   strConnectionProperty:   What property the Variable would be connected to
'   strVariableName:   The Variable Name we are looking for within the Parent Object and its children.
'RETURNS:   The Name of the User Prompt which is stored in the Variable's Initial Value (*if any)
'******************************************************************************************
    Dim intPosition As Integer             'The position where the equal sign is in the InitialValue
    Dim intStartPosition As Integer        'Two spaces after the equal sign is the actual User Prompt, the numeric value of this position is stored in this variable
    Dim strParameter As String        'The whole value stored in the Variable's Initial Value
    Dim strUserPrompt As String           'The User Prompt after strConnectionProperty and the = is taken from the string
    Dim objSubObject As Object        'Variable to hold each object in the For Loop
    
    On Error GoTo ErrorHandler
    BDW_GetInitValueOfVar = ""
    If strVariableName = "EasyDynamo" Then
        For Each objSubObject In objMainObject.ContainedObjects
        'if it is an Easy Dynamo Variable (and not a togglesource variable)
            If Left(objSubObject.Name, 10) = "EasyDynamo" Then
                If InStr(1, objSubObject.InitialValue, "ToggleSource") = 0 Then
                    strParameter = objSubObject.InitialValue
                    objSubObject.EnableAsVBAControl = False
                    Exit For
                End If
            End If
        Next objSubObject
    ElseIf strVariableName = "EasyDynamoToggleSource" Then
        For Each objSubObject In objMainObject.ContainedObjects
            If Left(objSubObject.Name, 22) = "EasyDynamoToggleSource" Then
                strParameter = objSubObject.InitialValue
                objSubObject.EnableAsVBAControl = False
                Exit For
            End If
        Next objSubObject
    End If
    
    intPosition = InStr(strParameter, strConnectionProperty & "=" & Chr$(34))
    
    'Fix for Version 1.0. Was writing "Source = {prompt}" to all Easy Dynamo
    'Variables in Version 1.0.
    'We need to retrieve that, and re-write to the variable a new initial value, which is
    '"its Connected Property = {prompt}"
    If intPosition = 0 Then
        intPosition = BDW_CheckForOldVariable(strParameter)
        '8 is for Len("Source =")
        intStartPosition = intPosition + 8
    Else
        '2 is for the equals sign and the space before it
        intStartPosition = intPosition + 2 + Len(strConnectionProperty)
    End If
    
    'if the connection property or the word "Source" was found in the variable's initial value,
    'intPosition will be > 0
    If intPosition Then
        intPosition = InStr(intStartPosition, strParameter, Chr$(34))
        While intPosition <= Len(strParameter) And intPosition
            If Mid(strParameter, intPosition, 1) = Chr$(34) And Mid(strParameter, intPosition + 1, 1) <> Chr$(34) Then
                strUserPrompt = Mid(strParameter, intStartPosition, intPosition - intStartPosition)
                BDW_GetInitValueOfVar = strUserPrompt
                Exit Function
            End If
            intPosition = InStr(intPosition + 2, strParameter, Chr$(34))
        Wend
    End If
    Exit Function

ErrorHandler:
    HandleError
End Function

Function BDW_CheckForOldVariable(strParameter As String) As Integer
'******************************************************************************************
'PURPOSE: To check if the initial value of the variable contains the words "Source ="
'           In version 1.0, that was written to all variables; "Source = {prompt}".
'           Now that we support object to object connections, we may have to write several
'           prompts to a variable. To make them unique for re-entrance,
'           we now write their connection property instead of the word "Source"
'INPUTS:
'           strParameter: the initial value of the variable
'
'RETURNS:
'           The position in the initial value where the strSource exists.
'           Will return 0 if strSource is not found.
'******************************************************************************************
    Dim strSource As String     'The string we are searching for
    
    strSource = "Source"
    BDW_CheckForOldVariable = InStr(strParameter, strSource & "=" & Chr$(34))

End Function

Sub BDW_SetVarInitialValue(objMainObject As Object, strConnectionProperty As String, strUserPrompt As String, strVariableName As String, blnObjToObj As Boolean)
'******************************************************************************************
'PURPOSE: To set the Initial Value of the EasyDynamo Variable equal to a string.
'The string consists of:
'   The type of property we are connecting to (ie: Source, Caption etc), an equal sign,
'   then the User Prompt.
'INPUTS:
'   objMainObject: The Parent Object we are searching
'   strConnectionProperty: The Property we are connecting the Variable to
'   strUserPrompt: The User Prompt
'   strVariableName:    The Variable Name
'RETURNS:
'******************************************************************************************
    Dim strParameter As String       'The whole string to set the Initial Value of the Variable to.
    Dim objSubObject As Object       'Variable to hold each object in the For Loop
    Dim strTempString As String     'string to hold the initial value of the variable when removing an old prompt
    Dim intStartPosition As Integer 'the position in the string where the old prompt starts
    Dim intPipePosition As Integer  'the position in the variable where the old prompt ends
    Dim intLengthOfVarValue As Integer  'the length of the initial value of the variable.
    
    On Error GoTo ErrorHandler
    
    If strVariableName = "EasyDynamo" Then
        strParameter = Trim(strParameter) & " " & strConnectionProperty & "=" & Chr$(34) & strUserPrompt & Chr$(34)
        For Each objSubObject In objMainObject.ContainedObjects
            If Left(objSubObject.Name, 10) = "EasyDynamo" And blnObjToObj = False Then
                objSubObject.InitialValue = strParameter
                'added for Clarify Case FD170230
                objSubObject.EnableAsVBAControl = False
                Exit For
            ElseIf Left(objSubObject.Name, 10) = "EasyDynamo" And blnObjToObj Then
                'this is for obj to obj connections, we need to add this info to the variable's current value
                'first we need to see if there is anything there, if no, put strParameter in there
                'if there is stuff in the varaible's initialvalue, first we need to see if this
                'connection property was already configured and remove its old prompt information
                'if it was not configured, but has stuff in it already, separate the initialvalue's
                'with a |
                If objSubObject.InitialValue = "" Then
                    objSubObject.InitialValue = strParameter
                    'added for Clarify Case FD170230
                    objSubObject.EnableAsVBAControl = False
                    Exit For
                Else
                    If InStr(1, objSubObject.InitialValue, strConnectionProperty) > 0 Then
                        'if that property had a prompt previously for this connection property,
                        'we need to remove it from the variable object's initial variable before adding
                        'the new prompt
                        intLengthOfVarValue = Len(objSubObject.InitialValue)
                        intStartPosition = InStr(1, objSubObject.InitialValue, strConnectionProperty)
                        intPipePosition = InStr(intStartPosition, objSubObject.InitialValue, "|")
                        If intPipePosition = 0 Then
                        'it is the only obj to obj connection, need to remove it.
                            objSubObject.InitialValue = strParameter
                            'added for Clarify Case FD170230
                            objSubObject.EnableAsVBAControl = False
                        Else
                            'get the left side of the old prompt
                            strTempString = LTrim(Left(objSubObject.InitialValue, intStartPosition - 1))
                            'get the right side of the old prompt
                            strTempString = strTempString & RTrim(Right(objSubObject.InitialValue, intLengthOfVarValue - intPipePosition))
                            objSubObject.InitialValue = strTempString & "|" & strParameter
                            'added for Clarify Case FD170230
                            objSubObject.EnableAsVBAControl = False
                            Exit For
                        End If
                    Else
                        'this connection property has not been configured yet at all
                        'but there are prompts already in the variable
                        objSubObject.InitialValue = objSubObject.InitialValue & "|" & strParameter
                        'added for Clarify Case FD170230
                        objSubObject.EnableAsVBAControl = False
                        Exit For
                    End If
                End If
            End If
        Next objSubObject
    ElseIf strVariableName = "EasyDynamoToggleSource" Then
        strParameter = Trim(strParameter) & " " & strConnectionProperty & ".Blink=" & Chr$(34) & strUserPrompt & Chr$(34)
        For Each objSubObject In objMainObject.ContainedObjects
            If Left(objSubObject.Name, 22) = "EasyDynamoToggleSource" Then
                objSubObject.InitialValue = strParameter
                'added for Clarify Case FD170230
                objSubObject.EnableAsVBAControl = False
                Exit For
            End If
        Next objSubObject
    End If
    Exit Sub
    
ErrorHandler:
    If Err.Number = 438 Then
        MsgBox "Object does not support this property.  Make sure you have the correct   SIM installed.  See Build Dynamo Wizard Release Notes."
        Exit Sub
    Else
        HandleError
    End If
End Sub

Sub BDW_SetProcedure(objMainObject As Object, strProcedureName As String, strScriptText As String)
'******************************************************************************************
'PURPOSE: Writes the call to the Dynamo's Edit Event that will launch the Dynamo form when
'the User drags and drops or double clicks the dynamo.
'INPUTS:
'   objMainObject: The Parent Object we are searching
'   strProcedureName: Name of the Procedure Object to write to.
'   strScriptText: The Script text to write to the Procedure Object
'******************************************************************************************
    Dim lngIndex As Long  'Numerical index of the procedures position in the existing collection.
    Dim lngFound As Long  'Returns 1 if an event procedure is present, 0 if no event procedure is present.
    
    On Error GoTo ErrorHandler
    objMainObject.Procedures.GetEventHandlerIndex strProcedureName, lngIndex, lngFound
    'If there is an Edit Procedure present, remove it.
    If lngFound = 1 Then
        objMainObject.Procedures.Remove lngIndex
    End If
    If objMainObject.ClassName <> "OleObject" Then
        objMainObject.Procedures.AddEventHandler strProcedureName, strScriptText, objMainObject.Procedures.Count + 1
        objMainObject.Commit
    End If
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub

Sub EditDynamo()
'******************************************************************************************
'PURPOSE: This initializes the Edit Dynamo form
'******************************************************************************************
    On Error GoTo ErrorHandler
    'NLS Object
    Set objStrMgr = CreateObject("FD_BDW.ResMgr")
    mintSymbolLines = 0
    BDW_CleanPromptList
    gblnNoSource = False
    gintNumberInArray = 0
    gintNumberOfUniquePrompts = 0
    gintTempNumberOfUniquePrompts = 0
    
    'Loop through all of the objects
    Set mobjParentObject = Application.ActiveDocument.Page.SelectedShapes.Item(1)
    mstrParentName = mobjParentObject.Name
    
    BDW_GetObjectInformation mobjParentObject, ""
    BDW_AddToPromptList
    BDW_PrepareFrmEditDynamo mstrParentName
    frmEditDynamo.Show
    'If the user did not hit Cancel
    If Not frmEditDynamo.mblnCancel Then
        BDW_ReplacePropertyWithNewSetting mobjParentObject, ""
    End If
    Unload frmEditDynamo
    'clean out the EditDynamo temporary array
    BDW_CleanEditDynamoArray
    'release all objects
    Set mobjParentObject = Nothing
    Set objStrMgr = Nothing
    
    Exit Sub

ErrorHandler:
    HandleError
End Sub
Sub BDW_CleanCreateDynamoArray()
'******************************************************************************************
'PURPOSE: To remove all data from the BDW_CreateDynamoArray.
'RETURNS: A "Clean" Data Type
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_CreateDynamoArray
    
    For intIndex = 0 To BDW_mintCDANumOfIndeces
        BDW_CreateDynamoArray(intIndex, 0) = ""
        BDW_CreateDynamoArray(intIndex, 1) = ""
        BDW_CreateDynamoArray(intIndex, 2) = ""
    Next intIndex
    
End Sub

Sub BDW_CleanEditDynamoArray()
'******************************************************************************************
'PURPOSE: To remove all data from the BDW_EditDynamoArray.
'RETURNS: A "Clean" Data Type
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_EditDynamoArray
    
    For intIndex = 0 To BDW_mintEDANumOfIndeces
        BDW_EditDynamoArray(intIndex, 0) = ""
        BDW_EditDynamoArray(intIndex, 1) = ""
    Next intIndex
    
End Sub

Sub BDW_CleanSymbolList()
'******************************************************************************************
'PURPOSE: To remove all data from the BDW_gudtSymbolData Data Type.
'RETURNS: A "Clean" Data Type
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    
    For intIndex = 0 To 200
        With BDW_gudtSymbolData(intIndex)
            .blnIsUsingSubstitution = False
            .strContent = ""
            .strField = ""
            .strFullName = ""
            .strName = ""
            .strPrompt = ""
            .strPropertyName = ""
        End With
    Next intIndex
 
End Sub
Sub BDW_CleanPromptList()
'******************************************************************************************
'PURPOSE: To remove all data from the BDW_mudtPromptData Data Type.
'RETURNS: A "Clean" Data Type
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.

    For intIndex = 0 To 100
        With BDW_mudtPromptData(intIndex)
            .blnIsUsingSubstitution = False
            .strContent = ""
            .strField = ""
            .strPrompt = ""
        End With
    Next intIndex

End Sub

Sub CreateDynamo()
'******************************************************************************************
'PURPOSE: To initialize the configuration form for the creation of a Dynamo.
'******************************************************************************************
    Dim intIndex As Integer      'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    Dim strError As String       'Error message
    Dim strTitle As String       'Title for Error message box
    Dim bCameFromHasNoConnections As Boolean ' Indicates we called BDW_HasNoConnections
    
    'NLS Object
    Set objStrMgr = CreateObject("FD_BDW.ResMgr")

    BDW_CleanSymbolList
    On Error GoTo ErrorHandler
    mintSymbolLines = 0
    BDW_mintCDANumOfIndeces = 0
    mblnNewGroupCreated = False
    mblnGroupAlreadyDone = False
    
    'If there is no object selected
    If Application.ActiveDocument.Page.SelectedShapes.Count = 0 Then
        GoTo NO_OBJECTS
    End If
    
    'If Application.ActiveDocument.Page.SelectedShapes.Count > 1 Then
    If Application.ActiveDocument.Page.SelectedShapes.Count > 0 Then 'New dynamo won't just rename the group, it will create a dynamo to contain any selected objects.
        'Application.ActiveDocument.Page.Group
        Application.ActiveDocument.Page.Create_Dynamo_By_Grouping
        
        mblnNewGroupCreated = True
    End If
    'Name of the picture, needed later
    mstrPictureName = Application.ActiveDocument.Name
    Set mobjParentObject = Application.ActiveDocument.Page.SelectedShapes.Item(1)
    mstrParentName = mobjParentObject.Name
    If mobjParentObject.ClassName = "OleObject" Then
        strError = objStrMgr.GetNLSStr(CLng(mERROR2))
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
        MsgBox strError, , strTitle
        Exit Sub
    End If
        
    BDW_GetObjectInformation mobjParentObject, ""
    BDW_PrepareFrmCreateDynamo mstrParentName
    
    'If the selected object does not have any connections
    If BDW_gudtSymbolData(0).strName = "" Then
        Call BDW_HasNoConnections(mobjParentObject)
        bCameFromHasNoConnections = True
    Else
    'Has Connections
        bCameFromHasNoConnections = False
        frmCreateDynamo.Show
    End If
    'If the User did not hit Cancel
    If (Not frmCreateDynamo.mblnCancel) And (False = bCameFromHasNoConnections) Then ' Only do this stuff if the user didnt cancel and if there are connections to do stuff with
        For intIndex = 0 To mintSymbolLines - 1
            'For partial subst.
            If (Left(BDW_gudtSymbolData(intIndex).strPrompt, 1) = "{") And (Right(BDW_gudtSymbolData(intIndex).strPrompt, 3) = "}.*") Then
                GoTo NextOne
            'If no Left and no Right Bracket
            ElseIf BDW_gudtSymbolData(intIndex).strPrompt <> "" And _
               (InStr(BDW_gudtSymbolData(intIndex).strPrompt, "{") Or _
                    InStr(BDW_gudtSymbolData(intIndex).strPrompt, "}")) = 0 Then
                BDW_gudtSymbolData(intIndex).strPrompt = "{" & BDW_gudtSymbolData(intIndex).strPrompt & "}"
            'Has Left but no Right bracket
            ElseIf BDW_gudtSymbolData(intIndex).strPrompt <> "" And _
                Left(BDW_gudtSymbolData(intIndex).strPrompt, 1) = "{" And _
                    InStr(BDW_gudtSymbolData(intIndex).strPrompt, "}") = 0 Then
                        BDW_gudtSymbolData(intIndex).strPrompt = BDW_gudtSymbolData(intIndex).strPrompt & "}"
            'Has Right but no Left Bracket
            ElseIf BDW_gudtSymbolData(intIndex).strPrompt <> "" And _
                Right(BDW_gudtSymbolData(intIndex).strPrompt, 1) = "}" And _
                    InStr(BDW_gudtSymbolData(intIndex).strPrompt, "{") = 0 Then
                        BDW_gudtSymbolData(intIndex).strPrompt = "{" + BDW_gudtSymbolData(intIndex).strPrompt
            End If
NextOne:
        Next intIndex
        BDW_CreateVarAndSetItsValue mobjParentObject, ""
        BDW_SetProcedure mobjParentObject, "Edit", "EditDynamo"
    ElseIf (frmCreateDynamo.mblnCancel) And (mblnNewGroupCreated) Then
    'User hit Cancel, and a Group was created out of their Selected Shapes.
        Application.ActiveDocument.Page.ungroup
    End If
    Set mobjParentObject = Nothing
    Unload frmCreateDynamo
    
    'clean out the CreateDynamo temporary array
    BDW_CleanCreateDynamoArray
    Exit Sub
    
NO_OBJECTS:
    strError = objStrMgr.GetNLSStr(CLng(mERROR1))
    strTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
    MsgBox strError, , strTitle
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub

Private Sub BDW_HasNoConnections(objMainObject As Object)
'******************************************************************************************
'PURPOSE: To allow the developer to name their new Dynamo when there are no connected properties
'INPUTS:
'   objMainObject: The Parent Object we are naming
'******************************************************************************************
'    Dim strDynamoName As String 'What the Developer wants to name their Dynamo
'    Dim strHelpPath As String   'The Path where the help file is located
'    Dim strError As String      'Error message
'    Dim strTitle As String      'Title for Error message box
'    Dim strPrompt As String     'kei090302 Clarify #257738 - Prompt message
    
'    If mblnError <> True Then
'        mstrBadName = ""
'        mblnError = False
'    End If
    
    On Error GoTo ErrorHandler
'    strHelpPath = System.NlsPath & "\BuildDynamoWizard.hlp"
    'kei090302 Clarify #257738 NLS
'    strTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
'    strPrompt = objStrMgr.GetNLSStr(CLng(NLS_mPrompt))
'    If mblnError = True Then
'        'kei090302 Clarify #257738 NLS
'        'strDynamoName = InputBox("Enter the Dynamo Name", "Build Dynamo Wizard", mstrBadName, , , strHelpPath, 999)
'        strDynamoName = InputBox(strPrompt, strTitle, mstrBadName, , , strHelpPath, 999)
'        mblnError = False
'    Else
'        'kei090302 Clarify #257738 NLS
'        'strDynamoName = InputBox("Enter the Dynamo Name", "Build Dynamo Wizard", objMainObject.Name, , , strHelpPath, 999)
'        strDynamoName = InputBox(strPrompt, strTitle, objMainObject.Name, , , strHelpPath, 999)
'    End If
'    ' If the developer hits cancel
'    If strDynamoName = "" Then
'        Exit Sub
'    Else
'        If objMainObject.Name <> strDynamoName Then
'            'if there is a space in the dynamo name
'            If InStr(1, strDynamoName, " ") = 0 Then
'                objMainObject.Name = strDynamoName
'            Else
'                strTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
'                'JPB031703  Tracker #18 added new line to message
'                strError = objStrMgr.GetNLSStr(CLng(NLS_mERROR1)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1a)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1e)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1b)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1c)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1d))
'                MsgBox strError, , strTitle
'                mstrBadName = strDynamoName
'                mblnError = True
'                BDW_HasNoConnections objMainObject
'            End If
'        End If
'    End If

    mstrParentName = mobjParentObject.Name
    BDW_PrepareFrmCreateDynamo mstrParentName, True

    frmCreateDynamo.Show
    Exit Sub
    
ErrorHandler:
    'Invalid Syntax for the Dynamo Name
    If Err.Number = 440 Then
        Dim strTitle As String
        Dim strError As String
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
        'JPB031703  Tracker #18 added new line to message
        strError = objStrMgr.GetNLSStr(CLng(NLS_mERROR1)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1a)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1e)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1b)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1c)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1d))
        MsgBox strError, , strTitle
        'mstrBadName = strDynamoName
        'mblnError = True
        BDW_HasNoConnections objMainObject
    Else
        HandleError
    End If
End Sub

Private Sub BDW_CreateVarAndSetItsValue(objMainObject As Object, strConnectionProperty As String)
'******************************************************************************************
'PURPOSE: For each Object in the Dynamo, if the object's Classname equals one of the cases,
'call BDW_CreateEasyDynamoVariable subroutine to build a new Variable object on the current object
'ASSUMPTIONS: Will only build a variable Object if the Connection was given a User Prompt.
'INPUTS:
'   objMainObject: The Parent Object whose Contained Objects we are looping through
'******************************************************************************************
    Dim objSubObject As Object        'Used in the For Loop
    Dim intContainedCount As Integer    'integer holding the number of Contained Objects of objMainObject
    Dim lngConnectedCount As Long       'Long holding the number of Connected objects to objMainObject
    Dim blnAlreadyDone As Boolean
    Dim intCounter As Integer
    Dim strPropertyName As String
    Dim strSource As String
    Dim strFullyQualifiedSource As String
    Dim vntSourceObjects
    Dim VarToObj As Object
    
    blnAlreadyDone = False
    'If (objMainObject.ClassName = "Group") Or (objMainObject.ClassName = "Chart") Or (objMainObject.ClassName = "Dynamo") Then
    'kei072508 iFix5.0 T6617
    If (objMainObject.ClassName = "Group") Or _
        (objMainObject.ClassName = "Chart") Or _
        (objMainObject.ClassName = "LineChart") Or _
        (objMainObject.ClassName = "SPCBarChart") Or _
        (objMainObject.ClassName = "HistogramChart") Or _
        (objMainObject.ClassName = "Dynamo") Then
        
        mblnGroupAlreadyDone = False
    End If
    objMainObject.connectedpropertyCount lngConnectedCount
    intContainedCount = objMainObject.ContainedObjects.Count

    On Error GoTo ErrorHandler
    Select Case objMainObject.ClassName
        Case mLOOKUP_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, strConnectionProperty, BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Source"), "EasyDynamo", False
            BDW_CreateEasyDynamoToggleSourceVariable objMainObject
            BDW_SetVarInitialValue objMainObject, strConnectionProperty, BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".ToggleSource"), "EasyDynamoToggleSource", False
        Case mTEXT_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, "Caption", BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Caption"), "EasyDynamo", False
        Case mLINEAR_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, strConnectionProperty, BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Source"), "EasyDynamo", False
        Case mFORMAT_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, strConnectionProperty, BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Source"), "EasyDynamo", False
        Case mFIXEVENT_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, strConnectionProperty, BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Source"), "EasyDynamo", False
        Case mVARIABLE_CLASS
            If Left(objMainObject.Name, 10) <> "EasyDynamo" And Left(objMainObject.Name, 22) <> "EasyDynamoToggleSource" Then
                blnAlreadyDone = True
                BDW_CreateEasyDynamoVariable objMainObject
                BDW_SetVarInitialValue objMainObject, "InitialValue", BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".InitialValue"), "EasyDynamo", False
            End If
        Case mOLEOBJECT_CLASS
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, "Caption", BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Caption"), "EasyDynamo", False
        Case mPEN_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, "Source", BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Source"), "EasyDynamo", False
        'kei072508 iFix5.0 T6617
        Case mREALTIMEDS_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, "Source", BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Source"), "EasyDynamo", False
        Case mSPCDS_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, "Source", BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Source"), "EasyDynamo", False
        Case mHISTDS_CLASS
            blnAlreadyDone = True
            BDW_CreateEasyDynamoVariable objMainObject
            BDW_SetVarInitialValue objMainObject, "Source", BDW_GetUserPrompt(objMainObject.Name, strConnectionProperty & ".Source"), "EasyDynamo", False
            
    End Select
    'if the object is a group, loop through its connected properties
    'If (objMainObject.ClassName = "Group") Or (objMainObject.ClassName = "Chart") Or (objMainObject.ClassName = "Dynamo") Then
    'kei072508 iFix5.0 T6617
    If (objMainObject.ClassName = "Group") Or _
        (objMainObject.ClassName = "Chart") Or _
        (objMainObject.ClassName = "LineChart") Or _
        (objMainObject.ClassName = "SPCBarChart") Or _
        (objMainObject.ClassName = "HistogramChart") Or _
        (objMainObject.ClassName = "Dynamo") Then
        
        'if the group was not done yet
        If mblnGroupAlreadyDone = False Then
            'if it has a connected property
            If lngConnectedCount > 0 Then
                'it has some type of animation, call get group connections on it.
                For intCounter = 1 To lngConnectedCount
                    objMainObject.Getconnectioninformation intCounter, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
                    'error check for unresolved dynamos created in 2.0
                    If (TypeName(vntSourceObjects) <> "Empty") Then
                        'error check for unresolved pictures created in 2.0
                        If (TypeName(vntSourceObjects(0)) <> "Nothing") Then
                            'JPB042403  Tracker #1328  don't check category if object is FixGlobalSysInfo (time or date link)
                            'since that object doesn't have a category property.
                            If TypeName(vntSourceObjects(0)) = "FixGlobalSysInfo" Then
                                BDW_CreateEasyDynamoVariable objMainObject
                                BDW_SetVarInitialValue objMainObject, strPropertyName, BDW_GetUserPrompt(objMainObject.Name, objMainObject.Name & "." & strPropertyName & ".Source"), "EasyDynamo", True
                            ElseIf vntSourceObjects(0).Category <> "Animation" Then
                                BDW_CreateEasyDynamoVariable objMainObject
                                BDW_SetVarInitialValue objMainObject, strPropertyName, BDW_GetUserPrompt(objMainObject.Name, objMainObject.Name & "." & strPropertyName & ".Source"), "EasyDynamo", True
                            Else
                                Set VarToObj = vntSourceObjects(0)
                                BDW_CreateVarAndSetItsValue VarToObj, strPropertyName
                            End If
                        End If
                    End If
                Next intCounter
                blnAlreadyDone = True
                mblnGroupAlreadyDone = True
            End If
        End If
     End If
     'if it is a group or a chart, need to loop through its contained objects
    'If (intContainedCount > 0 And objMainObject.ClassName = "Group") Or (intContainedCount > 0 And objMainObject.ClassName = "Chart") Or (intContainedCount > 0 And objMainObject.ClassName = "Dynamo") Then
    'kei072508 iFix5.0 T6617
    If (intContainedCount > 0 And objMainObject.ClassName = "Group") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "Chart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "LineChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "SPCBarChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "HistogramChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "Dynamo") Then
        
        For Each objSubObject In objMainObject.ContainedObjects
        'if it is not an animation object (those were done above)
            If objSubObject.ClassName <> mLOOKUP_CLASS And objSubObject.ClassName <> mLINEAR_CLASS And objSubObject.ClassName <> mFORMAT_CLASS Then
                BDW_CreateVarAndSetItsValue objSubObject, strConnectionProperty & "." & objSubObject.Name
                'BDW_CreateVarAndSetItsValue objSubObject, strPropertyName
                blnAlreadyDone = False
            End If
        Next objSubObject
    ElseIf blnAlreadyDone = False Or objMainObject.ClassName = mVARIABLE_CLASS Or objMainObject.ClassName = mOLEOBJECT_CLASS Or objMainObject.ClassName = mTEXT_CLASS Then
        'There are more connected properties than contained properties to the shape...We'll use the connected count
        'This means there are Object to Object connections
        For intCounter = 1 To lngConnectedCount
            objMainObject.Getconnectioninformation intCounter, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
            'error check for unresolved dynamos created in 2.0
            If TypeName(vntSourceObjects) <> "Empty" Then
                'error check for unresolved pictures created in 2.0
                If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                    'JPB042403  Tracker #1328  don't check category if object is FixGlobalSysInfo (time or date link)
                    'since that object doesn't have a category property.
                    If TypeName(vntSourceObjects(0)) = "FixGlobalSysInfo" Then
                        BDW_CreateEasyDynamoVariable objMainObject
                        'this gets around the problem of differentiating between the Caption of a text object
                        'and an obj to obj anim on the caption of a text object
                        If objMainObject.ClassName = mTEXT_CLASS And strPropertyName = "Caption" Then
                            strPropertyName = "AnimatedCaption"
                        End If
                        BDW_SetVarInitialValue objMainObject, strPropertyName, BDW_GetUserPrompt(objMainObject.Name, objMainObject.Name & "." & strPropertyName & ".Source"), "EasyDynamo", True
                    ElseIf vntSourceObjects(0).Category <> "Animation" Then
                        BDW_CreateEasyDynamoVariable objMainObject
                        'this gets around the problem of differentiating between the Caption of a text object
                        'and an obj to obj anim on the caption of a text object
                        If objMainObject.ClassName = mTEXT_CLASS And strPropertyName = "Caption" Then
                            strPropertyName = "AnimatedCaption"
                        End If
                        BDW_SetVarInitialValue objMainObject, strPropertyName, BDW_GetUserPrompt(objMainObject.Name, objMainObject.Name & "." & strPropertyName & ".Source"), "EasyDynamo", True
                    Else
                        Set VarToObj = vntSourceObjects(0)
                        BDW_CreateVarAndSetItsValue VarToObj, strPropertyName
                    End If
                End If
            End If
        Next intCounter
    End If
    Exit Sub

ErrorHandler:
    HandleError
End Sub

Public Sub BDW_GetObjectInformation(objMainObject As Object, strConnectionProperty As String, Optional IndexNumber As Integer)
'******************************************************************************************
'PURPOSE: For each Object in the Dynamo, if the object's Classname equals one of the cases,
'call BDW_AddToSymbolList to add the FullName of the Connection,
'the Current Data Source, and whether it is using substitution to the BDW_gudtSymbolData.
'INPUTS:
'   objMainObject: The Parent Object whose Contained Objects we are looping through
'   strConnectionProperty: Name of the Connected Object.
'******************************************************************************************
    Dim objMyOwnerObject As Object  'The Owner Object
    Dim objSubObject As Object      'Used in the For Loop
    Dim strPropertyName As String   'Property Name returned from GetConnectionInformation
    Dim strSource As String 'Source returned from GetConnectionInformation
    Dim strFullyQualifiedSource As String   'FullyQualifiedSource returned from GetConnectionInformation
    Dim vntSourceObjects    'Variant returned from GetConnectionInformation
    Dim intCounter As Integer   'Counter Used in a For Loop
    Dim intContainedCount As Integer    'number of Contained objects to objMainObject
    Dim lngConnectedCount As Long       'number of connected objects to objMainObject
    Dim blnAlreadyDone As Boolean
    
    'VBA6.0 changes..
    'VBA6.0 compiler is much more strict with object types than VBA5.0
    Dim sourceObject As Object
    
    On Error GoTo ErrorHandler
    blnAlreadyDone = False
    'if the object is a Fix Shape, and not the Parent Object, we need to find out how many contained objects and how many connected
    'properties it has.
    
    'kei072508 iFix5.0 T6617
    'If (objMainObject.ClassName = "Group") Or (objMainObject.ClassName = "Chart") Or (objMainObject.ClassName = "Dynamo") Then
    If (objMainObject.ClassName = "Group") Or _
        (objMainObject.ClassName = "Chart") Or _
        (objMainObject.ClassName = "LineChart") Or _
        (objMainObject.ClassName = "SPCBarChart") Or _
        (objMainObject.ClassName = "HistogramChart") Or _
        (objMainObject.ClassName = "Dynamo") Then
        
        mblnGroupAlreadyDone = False
    End If
    
    'get the number of contained and connected property counts
    objMainObject.connectedpropertyCount lngConnectedCount
    intContainedCount = objMainObject.ContainedObjects.Count
    'get the objects owner name
    Set objMyOwnerObject = objMainObject.owner

    Select Case objMainObject.ClassName
        Case mLOOKUP_CLASS
                blnAlreadyDone = True
                BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Source", BDW_GetInitValueOfVar(objMainObject, strConnectionProperty, "EasyDynamo"), objMainObject.Source
                'CMK 1-99874598  only display prompt for *.blink if there is a datasource associated with it
                If Len(objMainObject.ToggleSource) > 0 Then
                    BDW_AddToSymbolList objMainObject, strConnectionProperty & ".ToggleSource", BDW_GetInitValueOfVar(objMainObject, strConnectionProperty & ".Blink", "EasyDynamoToggleSource"), objMainObject.ToggleSource
                End If
        Case mTEXT_CLASS
            If ((objMainObject.owner.ClassName <> "Pen") And (objMainObject.owner.ClassName <> "Legend") And (objMainObject.owner.ClassName <> "TimeAxis") And (objMainObject.owner.ClassName <> "ValueAxis")) Then
                blnAlreadyDone = True
                BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Caption", BDW_GetInitValueOfVar(objMainObject, "Caption", "EasyDynamo"), objMainObject.Caption
            End If
        Case mLINEAR_CLASS
            blnAlreadyDone = True
            BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Source", BDW_GetInitValueOfVar(objMainObject, strConnectionProperty, "EasyDynamo"), objMainObject.Source
        Case mFORMAT_CLASS
            blnAlreadyDone = True
            BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Source", BDW_GetInitValueOfVar(objMainObject, strConnectionProperty, "EasyDynamo"), objMainObject.Source
        Case mFIXEVENT_CLASS
            blnAlreadyDone = True
            BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Source", BDW_GetInitValueOfVar(objMainObject, strConnectionProperty, "EasyDynamo"), objMainObject.Source
        Case mOLEOBJECT_CLASS
            On Error Resume Next
            BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Caption", BDW_GetInitValueOfVar(objMainObject, "Caption", "EasyDynamo"), objMainObject.Caption
        Case mVARIABLE_CLASS
            'If the variable is NOT an EasyDynamo Variable
            If Left(objMainObject.Name, 10) <> "EasyDynamo" And Left(objMainObject.Name, 22) <> "EasyDynamoToggleSource" Then
                blnAlreadyDone = True
                BDW_AddToSymbolList objMainObject, strConnectionProperty & ".InitialValue", BDW_GetInitValueOfVar(objMainObject, "InitialValue", "EasyDynamo"), objMainObject.InitialValue
            End If
        Case mPEN_CLASS
            blnAlreadyDone = True
            BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Source", BDW_GetInitValueOfVar(objMainObject, "Source", "EasyDynamo"), objMainObject.Source
        'kei072508 iFix5.0 T6617
        Case mREALTIMEDS_CLASS
            blnAlreadyDone = True
            BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Source", BDW_GetInitValueOfVar(objMainObject, "Source", "EasyDynamo"), objMainObject.Source
        Case mSPCDS_CLASS
            blnAlreadyDone = True
            BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Source", BDW_GetInitValueOfVar(objMainObject, "Source", "EasyDynamo"), objMainObject.Source
        Case mHISTDS_CLASS
            blnAlreadyDone = True
            BDW_AddToSymbolList objMainObject, strConnectionProperty & ".Source", BDW_GetInitValueOfVar(objMainObject, "Source", "EasyDynamo"), objMainObject.Source
            
    End Select
    'if the object is a group, loop through its connected properties
    'kei072508 iFix5.0 T6617
    'If (objMainObject.ClassName = "Group") Or (objMainObject.ClassName = "Chart") Or (objMainObject.ClassName = "Dynamo") Then
    If (objMainObject.ClassName = "Group") Or _
        (objMainObject.ClassName = "Chart") Or _
        (objMainObject.ClassName = "LineChart") Or _
        (objMainObject.ClassName = "SPCBarChart") Or _
        (objMainObject.ClassName = "HistogramChart") Or _
        (objMainObject.ClassName = "Dynamo") Then
        
        'if the group was not done yet
        If mblnGroupAlreadyDone = False Then
            'if it has a connected property
            If lngConnectedCount > 0 Then
                'it has some type of animation, call get group connections on it.
                For intCounter = 1 To lngConnectedCount
                    objMainObject.Getconnectioninformation intCounter, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
                    'error check for unresolved dynamos created in 2.0
                    If TypeName(vntSourceObjects) <> "Empty" Then
                        'error check for unresolved pictures created in 2.0
                        If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                            'JPB042403  Tracker #1328  don't check category if object is FixGlobalSysInfo (time or date link)
                            'since that object doesn't have a category property.
                            If TypeName(vntSourceObjects(0)) = "FixGlobalSysInfo" Then
                                Call BDW_GetObjectToObjectConnections(objMainObject, strPropertyName, objMainObject.Name & "." & strPropertyName & ".Source", BDW_GetInitValueOfVar(objMainObject, strPropertyName, "EasyDynamo"), strFullyQualifiedSource)
                            ElseIf vntSourceObjects(0).Category <> "Animation" Then
                                Call BDW_GetObjectToObjectConnections(objMainObject, strPropertyName, objMainObject.Name & "." & strPropertyName & ".Source", BDW_GetInitValueOfVar(objMainObject, strPropertyName, "EasyDynamo"), strFullyQualifiedSource)
                            Else
                                'VBA6.0 changes
                                Set sourceObject = vntSourceObjects(0)
                                BDW_GetObjectInformation sourceObject, strPropertyName, intCounter
                                Set sourceObject = Nothing
                                
                            End If
                        End If
                    End If
                Next intCounter
                blnAlreadyDone = True
                mblnGroupAlreadyDone = True
            End If
        End If
     End If
    'if it is a group or a chart, need to look at its contained objects
    'kei072508 iFix5.0 T6617
    'If (intContainedCount > 0 And objMainObject.ClassName = "Group") Or (intContainedCount > 0 And objMainObject.ClassName = "Chart") Or (intContainedCount > 0 And objMainObject.ClassName = "Dynamo") Then
    If (intContainedCount > 0 And objMainObject.ClassName = "Group") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "Chart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "LineChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "SPCBarChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "HistogramChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "Dynamo") Then
        
        For Each objSubObject In objMainObject.ContainedObjects
        'if it is not an animation object (those were done above)
        If objSubObject.ClassName <> mLOOKUP_CLASS And objSubObject.ClassName <> mLINEAR_CLASS And objSubObject.ClassName <> mFORMAT_CLASS Then
            On Error Resume Next
            BDW_GetObjectInformation objSubObject, strConnectionProperty & "." & objSubObject.Name
            blnAlreadyDone = False
        End If
        Next objSubObject
    ElseIf blnAlreadyDone = False Or objMainObject.ClassName = mVARIABLE_CLASS Or objMainObject.ClassName = mOLEOBJECT_CLASS Or objMainObject.ClassName = mTEXT_CLASS Then
    'There are more connected properties than contained properties to the shape...We'll use the connected intCounter
        'This means there are Object to Object connections
        For intCounter = 1 To lngConnectedCount
            objMainObject.Getconnectioninformation intCounter, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
                'error check for unresolved dynamos created in 2.0
                If TypeName(vntSourceObjects) <> "Empty" Then
                    'error check for unresolved pictures created in 2.0
                    If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                        'JPB042403  Tracker #1328  don't check category if object is FixGlobalSysInfo (time or date link)
                        'since that object doesn't have a category property.
                        If TypeName(vntSourceObjects(0)) = "FixGlobalSysInfo" Then
                            If objMainObject.ClassName = mTEXT_CLASS And strPropertyName = "Caption" Then
                                strPropertyName = "AnimatedCaption"
                            End If
                            Call BDW_GetObjectToObjectConnections(objMainObject, strPropertyName, objMainObject.Name & "." & strPropertyName & ".Source", BDW_GetInitValueOfVar(objMainObject, strPropertyName, "EasyDynamo"), strFullyQualifiedSource)
                        ElseIf vntSourceObjects(0).Category <> "Animation" Then
                            'this gets around the problem of differentiating between the Caption of a text object
                            'and an obj to obj anim on the caption of a text object
                            If objMainObject.ClassName = mTEXT_CLASS And strPropertyName = "Caption" Then
                                strPropertyName = "AnimatedCaption"
                            End If
                            Call BDW_GetObjectToObjectConnections(objMainObject, strPropertyName, objMainObject.Name & "." & strPropertyName & ".Source", BDW_GetInitValueOfVar(objMainObject, strPropertyName, "EasyDynamo"), strFullyQualifiedSource)
                        Else
                            'VBA 6.0 changes
                            Set sourceObject = vntSourceObjects(0)
                            BDW_GetObjectInformation sourceObject, strPropertyName, intCounter
                            Set sourceObject = Nothing
                            
                        End If
                    End If
                End If
                'End If
            vntSourceObjects = Empty
        Next intCounter
    End If
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub

Sub BDW_CreateEasyDynamoVariable(objMainObject As Object)
'******************************************************************************************
'PURPOSE: A lot of information is stored in the BDW_gudtSymbolData Data Type. Needed a way to acces that
'information. By creating a variable object, and setting its Initial Value equal to the
'User Prompt, we have a way to access the rest of the information in data type.
'ie: If variable.InitialValue = guidtSymbolData(a).strPrompt then
'ASSUMPTIONS: A User Prompt was assigned to the Connection.
'EFFECTS: If there was an EasyDynamo Variable object previously created, and the user prompt
'for that connection was now removed, the variable object will be destroyed.
'INPUTS:
'   objMainObject: The Parent Object whose Contained Objects we are looping through
'******************************************************************************************
    Dim blnExists As Boolean   'Set to True if an EasyDynamo Variable exists
    Dim objSubObject As Object    'Used in the For Loop, then used as the new variable object
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    Dim objVariableObject As Object 'A previously created Variable Object
    
    On Error GoTo ErrorHandler
    blnExists = False
    For Each objSubObject In objMainObject.ContainedObjects
        If Left(objSubObject.Name, 10) = "EasyDynamo" Then
            Set objVariableObject = objSubObject
            blnExists = True
        End If
    Next objSubObject
    
    For intIndex = 0 To 200
        With BDW_gudtSymbolData(intIndex)
            If (.strName = objMainObject.Name) Then
                If (blnExists = False) Then
                    If .strPrompt <> "" Then
                        Set objSubObject = objMainObject.BuildObject("Variable")
                        objSubObject.Name = "EasyDynamo"
                        objSubObject.VariableType = vbString
                        objSubObject.InitialValue = ""
                        'added for Clarify Case FD170230
                        objSubObject.EnableAsVBAControl = False
                        Exit For
                    End If
                ElseIf blnExists = True And .strPrompt = "" Then
                    'need to destroy the variable
                    objVariableObject.DestroyObject
                    Exit For
                End If
                Exit For
            End If
        End With
    Next intIndex
    Exit Sub
    
ErrorHandler:
    If Err.Number = 438 Then
        MsgBox "Object does not support this property.  Make sure you have the correct   SIM installed.  See Build Dynamo Wizard Release Notes."
        Exit Sub
    Else
        HandleError
    End If
End Sub

Sub BDW_CreateEasyDynamoToggleSourceVariable(objMainObject As Object)
'******************************************************************************************
'PURPOSE: To loop through each Contained Object in the Dynamo, and check if it
'currently has an EasyDynamoToggleSource Variable connected to it. If not, create one.
'ASSUMPTIONS: A User Prompt was assigned to the Connection.
'EFFECTS: If there was a EasyDynamoToggleSource Variable object previously created, and the
'user prompt for that connection was now removed, the variable object will be destroyed.
'INPUTS:
'   objMainObject: The Parent Object whose Contained Objects we are looping through
'******************************************************************************************
    Dim blnExists As Boolean   'Set to True if an EasyDynamoToggleSource Variable exists
    Dim objSubObject As Object    'Used in the For Loop, then used as the new variable object
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    Dim objVariableObject As Object 'A previously created Variable Object
    
    On Error GoTo ErrorHandler
    blnExists = False
    For Each objSubObject In objMainObject.ContainedObjects
        If Left(objSubObject.Name, 22) = "EasyDynamoToggleSource" Then
            Set objVariableObject = objSubObject
            blnExists = True
        End If
    Next objSubObject
    
    For intIndex = 0 To 200
        With BDW_gudtSymbolData(intIndex)
            If (.strName = objMainObject.Name) Then
                If (blnExists = False) Then
                    If .strPrompt <> "" Then
                        Set objSubObject = objMainObject.BuildObject("Variable")
                        objSubObject.Name = "EasyDynamoToggleSource"
                        objSubObject.VariableType = vbString
                        objSubObject.InitialValue = ""
                        'added for Clarify Case FD170230
                        objSubObject.EnableAsVBAControl = False
                        Exit For
                    End If
                ElseIf blnExists = True And .strPrompt = "" Then
                    'need to destroy the variable
                    objVariableObject.DestroyObject
                    Exit For
                End If
                Exit For
            End If
        End With
    Next intIndex
    Exit Sub

ErrorHandler:
    If Err.Number = 438 Then
        MsgBox "Object does not support this property.  Make sure you have the correct   SIM installed.  See Build Dynamo Wizard Release Notes."
        Exit Sub
    Else
        HandleError
    End If
End Sub
Private Sub BDW_ReplacePropertyWithNewSetting(objMainObject As Object, strConnectionProperty As String)
'******************************************************************************************
'PURPOSE:For each Object in the Dynamo, if the object's Classname equals one of the cases,
'go to their index in BDW_gudtSymbolData (done in the BDW_GetCurrentSetting subroutine),
'and get its new Data Source (Caption, or Current Value) which is stored as .strContent.
'Then assign it to the object's Source (Caption, InitialValue) Property.
'INPUTS:
'   objMainObject: The Parent Object whose Contained Objects we are looping through
'   strConnectionProperty: Name of the connection whose current Data Source (Caption, or Value)
'   the user may have changed
'******************************************************************************************
    Dim objSubObject As Object      'Used in the For Loop
    Dim strObjectName As String     'The Parent Object's Name
    Dim strData As String           'Return value from BDW_GetCurrentSetting
    Dim intContainedCount As Integer    'number of Contained objects to objMainObject
    Dim lngConnectedCount As Long       'number of connected objects to objMainObject
    Dim lngStatus As Long           'Status returned from the Connect Method
    Dim strSource As String         'Source returned from GetConnectionInformation
    Dim strFullyQualifiedSource As String   'FullyQualified Source returned from GetConnectionInformation
    Dim strPropertyName As String       'property name returned from GetConnectionInformation
    Dim vntSourceObjects     ''Variant object returned from GetConnectionInformation
    Dim blnAlreadyDone As Boolean
    Dim intCounter As Integer
    Dim blnHasConnection As Boolean
    Dim lngIndex As Long
    Dim VarToObj As Object  'ce111700 bugfix #896
    
    On Error GoTo ErrorHandler
    
    blnAlreadyDone = False
    'If (objMainObject.ClassName = "Group") Or (objMainObject.ClassName = "Chart") Or (objMainObject.ClassName = "Dynamo") Then
    'kei072508 iFix5.0 T6617
    If (objMainObject.ClassName = "Group") Or _
        (objMainObject.ClassName = "Chart") Or _
        (objMainObject.ClassName = "LineChart") Or _
        (objMainObject.ClassName = "SPCBarChart") Or _
        (objMainObject.ClassName = "HistogramChart") Or _
        (objMainObject.ClassName = "Dynamo") Then
        mblnGroupAlreadyDone = False
    End If
    objMainObject.connectedpropertyCount lngConnectedCount
    intContainedCount = objMainObject.ContainedObjects.Count
    
    strObjectName = objMainObject.Name
    Select Case objMainObject.ClassName
        Case mLOOKUP_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Source")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Source = strData
                End If
            End If
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".ToggleSource")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.ToggleSource = strData
                End If
            End If
        Case mTEXT_CLASS
            If ((objMainObject.owner.ClassName <> "Pen") And (objMainObject.owner.ClassName <> "Legend") And (objMainObject.owner.ClassName <> "TimeAxis") And (objMainObject.owner.ClassName <> "ValueAxis")) Then
                strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Caption")
                blnAlreadyDone = True
                If strData <> mNO_PROMPT Then
                    If Not (strData = "") Then
                        objMainObject.Caption = strData
                    End If
                End If
            End If
         Case mLINEAR_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Source")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Source = strData
                End If
            End If
         Case mFORMAT_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Source")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Source = strData
                End If
            End If
         Case mFIXEVENT_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Source")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Source = strData
                End If
            End If
        Case mOLEOBJECT_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Caption")
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Caption = strData
                End If
            End If
        Case mVARIABLE_CLASS
            If Left(objMainObject.Name, 10) <> "EasyDynamo" And Left(objMainObject.Name, 22) <> "EasyDynamoToggleSource" Then
                strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".InitialValue")
                blnAlreadyDone = True
                If strData <> mNO_PROMPT Then
                    If Not (strData = "") Then
                        objMainObject.InitialValue = strData
                    End If
                End If
            Else
                blnAlreadyDone = True
            End If
        Case mPEN_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Source")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Source = strData
                End If
            End If
            
        'kei072508 iFix5.0 T6617
        Case mREALTIMEDS_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Source")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Source = strData
                End If
            End If
        Case mSPCDS_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Source")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Source = strData
                End If
            End If
        Case mHISTDS_CLASS
            strData = BDW_GetCurrentSetting(objMainObject.Name, strConnectionProperty & ".Source")
            blnAlreadyDone = True
            If strData <> mNO_PROMPT Then
                If Not (strData = "") Then
                    objMainObject.Source = strData
                End If
            End If
    End Select

    'If (objMainObject.ClassName = "Group") Or (objMainObject.ClassName = "Chart") Or (objMainObject.ClassName = "Dynamo") Then
    'kei072508 iFix5.0 T6617
    If (objMainObject.ClassName = "Group") Or _
        (objMainObject.ClassName = "Chart") Or _
        (objMainObject.ClassName = "LineChart") Or _
        (objMainObject.ClassName = "SPCBarChart") Or _
        (objMainObject.ClassName = "HistogramChart") Or _
        (objMainObject.ClassName = "Dynamo") Then
        
        'if the group was not done yet
        If mblnGroupAlreadyDone = False Then
            'if it has a connected property
            If lngConnectedCount > 0 Then
                'it has some type of animation, call get group connections on it.
                ReDim MyTestArray(lngConnectedCount)
                For intCounter = 1 To lngConnectedCount
                    'need to do this as a separate For Loop, as opposed to calling the Connect item each time
                    'because the GetConnectionInformation was returning the same information
                    'for different indeces because of the Connect Method
                    objMainObject.Getconnectioninformation intCounter, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
                    'error check for unresolved dynamos created in 2.0
                    If TypeName(vntSourceObjects) <> "Empty" Then
                        'error check for unresolved pictures created in 2.0
                        If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                            MyTestArray(intCounter).strProperty = strPropertyName
                            MyTestArray(intCounter).strSource = strSource
                        End If
                    End If
                Next intCounter
                intCounter = 0
                For intCounter = 1 To lngConnectedCount
                    objMainObject.IsConnected MyTestArray(intCounter).strProperty, blnHasConnection, lngIndex, lngStatus
                    objMainObject.Getconnectioninformation lngIndex, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
                    'error check for unresolved dynamos created in 2.0
                    If TypeName(vntSourceObjects) <> "Empty" Then
                        'error check for unresolved pictures created in 2.0
                        If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                            'JPB042403  Tracker #1328  don't check category if object is FixGlobalSysInfo (time or date link)
                            'since that object doesn't have a category property.
                            If TypeName(vntSourceObjects(0)) = "FixGlobalSysInfo" Then
                                strData = BDW_GetCurrentSetting(objMainObject.Name, objMainObject.Name & "." & strPropertyName & ".Source")
                                If strData <> mNO_PROMPT Then
                                    If Not (strData = "") Then
                                        objMainObject.Disconnect strPropertyName
                                        objMainObject.Connect strPropertyName, strData, lngStatus
                                    End If
                                End If
                            ElseIf vntSourceObjects(0).Category <> "Animation" Then
                                strData = BDW_GetCurrentSetting(objMainObject.Name, objMainObject.Name & "." & strPropertyName & ".Source")
                                If strData <> mNO_PROMPT Then
                                    If Not (strData = "") Then
                                        objMainObject.Disconnect strPropertyName
                                        objMainObject.Connect strPropertyName, strData, lngStatus
                                    End If
                                End If
                            Else
                                'BDW_ReplacePropertyWithNewSetting vntSourceObjects(0), strConnectionProperty & "." & vntSourceObjects(0).Name
                                'ce111700 Bugfix 896
                                'Convert the variant into an object for VBA 6.0
                                Set VarToObj = vntSourceObjects(0)
                                BDW_ReplacePropertyWithNewSetting VarToObj, strPropertyName
                            End If
                        End If
                    End If
                Next intCounter
                blnAlreadyDone = True
                mblnGroupAlreadyDone = True
            End If
        End If
     End If
    'if it is a group or a chart, need to look at its contained objects
    'If (intContainedCount > 0 And objMainObject.ClassName = "Group") Or (intContainedCount > 0 And objMainObject.ClassName = "Chart") Or (intContainedCount > 0 And objMainObject.ClassName = "Dynamo") Then
    'kei072508 iFix5.0 T6617
    If (intContainedCount > 0 And objMainObject.ClassName = "Group") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "Chart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "LineChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "SPCBarChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "HistogramChart") Or _
        (intContainedCount > 0 And objMainObject.ClassName = "Dynamo") Then
        
        For Each objSubObject In objMainObject.ContainedObjects
        'if it is not an animation object (those were done above)
        If objSubObject.ClassName <> mLOOKUP_CLASS And objSubObject.ClassName <> mLINEAR_CLASS And objSubObject.ClassName <> mFORMAT_CLASS Then
            BDW_ReplacePropertyWithNewSetting objSubObject, strConnectionProperty & "." & objSubObject.Name
            blnAlreadyDone = False
        End If
        Next objSubObject
    ElseIf blnAlreadyDone = False Or objMainObject.ClassName = mVARIABLE_CLASS Or objMainObject.ClassName = mOLEOBJECT_CLASS Or objMainObject.ClassName = mTEXT_CLASS Then
        'There are more connected properties than contained properties to the shape...We'll use the connected count
        'This means there are Object to Object connections
        Erase MyTestArray
        ReDim MyTestArray(lngConnectedCount)
        For intCounter = 1 To lngConnectedCount
            'need to do this as a separate For Loop, as opposed to calling the Connect item each time
            'because the GetConnectionInformation was returning the same information
            'for different indeces because of the Connect Method
            objMainObject.Getconnectioninformation intCounter, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
            'error check for unresolved dynamos created in 2.0
            If TypeName(vntSourceObjects) <> "Empty" Then
                'error check for unresolved pictures created in 2.0
                If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                    MyTestArray(intCounter).strProperty = strPropertyName
                    MyTestArray(intCounter).strSource = strSource
                End If
            End If
        Next intCounter
        intCounter = 0
        For intCounter = 1 To lngConnectedCount
            objMainObject.IsConnected MyTestArray(intCounter).strProperty, blnHasConnection, lngIndex, lngStatus
            objMainObject.Getconnectioninformation lngIndex, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
            'error check for unresolved dynamos created in 2.0
            If TypeName(vntSourceObjects) <> "Empty" Then
                'error check for unresolved pictures created in 2.0
                If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                    'JPB042403  Tracker #1328  don't check category if object is FixGlobalSysInfo (time or date link)
                    'since that object doesn't have a category property.
                    If TypeName(vntSourceObjects(0)) = "FixGlobalSysInfo" Then
                        'this gets around the problem of differentiating between the Caption of a text object
                        'and an obj to obj anim on the caption of a text object
                        If objMainObject.ClassName = mTEXT_CLASS And strPropertyName = "Caption" Then
                            strPropertyName = "AnimatedCaption"
                        End If
                        strData = BDW_GetCurrentSetting(objMainObject.Name, objMainObject.Name & "." & strPropertyName & ".Source")
                        'this gets around the problem of differentiating between the Caption of a text object
                        'and an obj to obj anim on the caption of a text object
                        If objMainObject.ClassName = mTEXT_CLASS And strPropertyName = "AnimatedCaption" Then
                            strPropertyName = "Caption"
                        End If
                        If strData <> mNO_PROMPT Then
                            If Not (strData = "") Then
                                objMainObject.Disconnect strPropertyName
                                objMainObject.Connect strPropertyName, strData, lngStatus
                            End If
                        End If
                    ElseIf vntSourceObjects(0).Category <> "Animation" Then
                        'this gets around the problem of differentiating between the Caption of a text object
                        'and an obj to obj anim on the caption of a text object
                        If objMainObject.ClassName = mTEXT_CLASS And strPropertyName = "Caption" Then
                            strPropertyName = "AnimatedCaption"
                        End If
                        strData = BDW_GetCurrentSetting(objMainObject.Name, objMainObject.Name & "." & strPropertyName & ".Source")
                        'this gets around the problem of differentiating between the Caption of a text object
                        'and an obj to obj anim on the caption of a text object
                        If objMainObject.ClassName = mTEXT_CLASS And strPropertyName = "AnimatedCaption" Then
                            strPropertyName = "Caption"
                        End If
                        If strData <> mNO_PROMPT Then
                            If Not (strData = "") Then
                                objMainObject.Disconnect strPropertyName
                                objMainObject.Connect strPropertyName, strData, lngStatus
                            End If
                        End If
                    Else
                        'BDW_ReplacePropertyWithNewSetting vntSourceObjects(0), strConnectionProperty & "." & vntSourceObjects(0).Name
                        'ce111700 Bugfix 896
                        'Convert the variant into an object for VBA 6.0
                        Set VarToObj = vntSourceObjects(0)
                        BDW_ReplacePropertyWithNewSetting VarToObj, strPropertyName
                    End If
                End If
            End If
        Next intCounter
        
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub

Private Sub BDW_AddToPromptList()
'******************************************************************************************
'PURPOSE: For each member of the public BDW_gudtSymbolData user defined type array, get its prompt
'and field, check to see if substitution is being used, and set the substitution flag to True is there is.
'Then, add their User Prompt, and Current Value (Data Source, Caption etc) into the
'BDW_mudtPromptData Data Type array.
'******************************************************************************************
    Dim intIndex As Integer                 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    Dim strPrompt As String                 'The User Prompt
    Dim strContent As String                'The Node and Tag, Caption, or Initial Value.
    Dim strField As String                  'The Field of the Data Source, if there is one
    Dim intStartPosition As Integer         'The start position of the User Prompt
    Dim intEndPosition As Integer           'The end position of the User Prompt
    Dim intRightCut As Integer              'The User Prompt, preceded by what the connection is to, and an equals sign
    Dim blnIsUsingSubstitution As Boolean   'Set to True if the user is using Substitution.

    On Error GoTo ErrorHandler
    mintPromptLines = 0
    'Loop through all of the indices in the data type
    For intIndex = 0 To mintSymbolLines - 1
        strContent = BDW_gudtSymbolData(intIndex).strContent
        'The indices for both Data Types are the same for every property.
        'If there is not a User Prompt filled in, the BDW_mudtPromptData Data Type index is left blank.
        If BDW_gudtSymbolData(intIndex).strPrompt <> "" Then
            strPrompt = BDW_gudtSymbolData(intIndex).strPrompt
            strField = BDW_gudtSymbolData(intIndex).strField
            If InStr(1, strPrompt, ".*") Then
                'using partial subst.
                blnIsUsingSubstitution = True
                BDW_gudtSymbolData(intIndex).blnIsUsingSubstitution = blnIsUsingSubstitution
            Else
                blnIsUsingSubstitution = False
                BDW_gudtSymbolData(intIndex).blnIsUsingSubstitution = blnIsUsingSubstitution
            End If
            intStartPosition = InStr(strPrompt, "{")
            strPrompt = Mid(strPrompt, intStartPosition + 1)
            intEndPosition = InStr(strPrompt, "}")
            If intEndPosition > 0 Then
                strPrompt = VBA.Left(strPrompt, intEndPosition - 1)
            End If
            If mintPromptLines < mMAX_PROMPT_SIZE Then
                BDW_mudtPromptData(mintPromptLines).strPrompt = strPrompt
                strPrompt = BDW_gudtSymbolData(intIndex).strPrompt
                intRightCut = Len(strPrompt) - intEndPosition - intStartPosition
                If (blnIsUsingSubstitution = False) Then
                    BDW_mudtPromptData(mintPromptLines).blnIsUsingSubstitution = False
                    If UCase(VBA.Left(strPrompt, intStartPosition - 1)) = UCase(VBA.Left(strContent, intStartPosition - 1)) And _
                        UCase(VBA.Right(strPrompt, intRightCut)) = UCase(VBA.Right(strContent, intRightCut)) Then
                        strContent = Mid(strContent, intStartPosition)
                        strContent = VBA.Left(strContent, Len(strContent) - intRightCut)
                    End If
                Else
                    BDW_mudtPromptData(mintPromptLines).blnIsUsingSubstitution = True
                End If
                BDW_mudtPromptData(mintPromptLines).strContent = strContent
                BDW_mudtPromptData(mintPromptLines).strField = strField
            End If
            mintPromptLines = mintPromptLines + 1
        Else
            mintPromptLines = mintPromptLines + 1
        End If
    Next intIndex
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub

Private Function BDW_GetCurrentSetting(strNameOfObject As String, strConnectionProperty As String)
'******************************************************************************************
'PURPOSE: To retrieve the current Data Source (Caption, or Value) of the Connection Property
'passed in.
'INPUTS:
'   strConnectionProperty: The Full Name of the connection
'RETURNS: The current Data Source, Caption, or Value (.strContent) of strConnectionProperty
'******************************************************************************************
    Dim intIndex As Integer             'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    Dim strPrompt As String             'User Prompt
    Dim strPrefix As String             'any text to the left of the {
    Dim strField As String              'the Field of the Data Source (*if any)
    Dim strContent As String            'The Node.Tag, InitialValue, or Caption
    Dim intLengthOfContent As Integer   'Length of strContent
    
    On Error GoTo ErrorHandler
    For intIndex = 0 To mintSymbolLines - 1
        If (strConnectionProperty = BDW_gudtSymbolData(intIndex).strFullName) And (strNameOfObject = BDW_gudtSymbolData(intIndex).strName) Then
            If BDW_gudtSymbolData(intIndex).strPrompt = "" Then
                BDW_GetCurrentSetting = mNO_PROMPT
                Exit Function
            Else
                strContent = BDW_gudtSymbolData(intIndex).strContent
                strField = BDW_gudtSymbolData(intIndex).strField
                strPrompt = BDW_gudtSymbolData(intIndex).strPrompt
                strPrefix = VBA.Left(strPrompt, InStr(strPrompt, "{") - 1)
                strPrompt = Mid(strPrompt, InStr(strPrompt, "{") + 1)
                strPrompt = VBA.Left(strPrompt, InStr(strPrompt, "}") - 1)
                If BDW_gudtSymbolData(intIndex).blnIsUsingSubstitution = True Then
                    BDW_GetCurrentSetting = strPrefix & BDW_mudtPromptData(intIndex).strContent & strField
                Else
                'if the old and new Tag and Fields do not match in Edit Dynamo, or the old and new content fields do not match, and the field is not empty
                    If ((strContent <> BDW_mudtPromptData(intIndex).strContent) And (strField <> "")) Or ((strField <> BDW_mudtPromptData(intIndex).strField) And (strField <> "")) Then
                    'there was a change, empty .strfield and .strcontent and fill them back in.
                        BDW_mudtPromptData(intIndex).strField = ""
                        intLengthOfContent = Len(BDW_mudtPromptData(intIndex).strContent)
                        If InStr(1, BDW_mudtPromptData(intIndex).strContent, ".") > 0 Then
                            Do While (intLengthOfContent > 0)
                                If Mid(BDW_mudtPromptData(intIndex).strContent, intLengthOfContent, 1) = "." Then
                                    BDW_mudtPromptData(intIndex).strField = Mid(BDW_mudtPromptData(intIndex).strContent, intLengthOfContent, 1) + BDW_mudtPromptData(intIndex).strField
                                    BDW_mudtPromptData(intIndex).strContent = Left(BDW_mudtPromptData(intIndex).strContent, intLengthOfContent - 1)
                                    Exit Do
                                Else
                                    BDW_mudtPromptData(intIndex).strField = Mid(BDW_mudtPromptData(intIndex).strContent, intLengthOfContent, 1) + BDW_mudtPromptData(intIndex).strField
                                    intLengthOfContent = intLengthOfContent - 1
                                End If
                            Loop
                        End If
                    End If
                End If
                BDW_GetCurrentSetting = strPrefix & BDW_mudtPromptData(intIndex).strContent & BDW_mudtPromptData(intIndex).strField
                Exit Function
            End If
        End If
    Next intIndex
    Exit Function
    
ErrorHandler:
    HandleError
End Function

Private Function BDW_GetUserPrompt(strNameOfObject As String, strConnectionProperty As String) As String
'******************************************************************************************
'PURPOSE: Given the Connection Property, find the index in BDW_gudtSymbolData whose .strFullName
'matches it, and return its User Prompt.
'INPUTS:
'   strConnectionProperty: The Connection Property whose User Prompt you wish to return.
'RETURNS: The User Prompt
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    
    For intIndex = 0 To mintSymbolLines
        If (strConnectionProperty = BDW_gudtSymbolData(intIndex).strFullName) And (strNameOfObject = BDW_gudtSymbolData(intIndex).strName) Then
            BDW_GetUserPrompt = BDW_gudtSymbolData(intIndex).strPrompt
            Exit Function
        End If
    Next intIndex
    
End Function

Private Function GetMyFamily(objMainObject As Object, strPropertyName As String) As String
Dim strName As String
Dim ImmediateOwnerName As String
Dim strTempOwner As String
Dim ImmediateOwnerObject As Object
Dim objTempOwner As Object

    Set ImmediateOwnerObject = objMainObject.owner
    ImmediateOwnerName = ImmediateOwnerObject.Name
    strName = objMainObject.Name
    
    If strName = mstrParentName Then
        GetMyFamily = strPropertyName
    'groups within groups
    ElseIf ImmediateOwnerName = mstrParentName Then
        GetMyFamily = strName & "." & strPropertyName
    Else
        strTempOwner = ImmediateOwnerName
        'Loop through until you hit the Parent Name
        Do While ImmediateOwnerName <> mstrParentName
            Set objTempOwner = ImmediateOwnerObject.owner
            If objTempOwner.Name = mstrParentName Then
                Exit Do
            End If
            strTempOwner = objTempOwner.Name & "." & strTempOwner
            ImmediateOwnerName = objTempOwner.Name
            Set ImmediateOwnerObject = objTempOwner
        Loop
        ImmediateOwnerName = strTempOwner
        GetMyFamily = ImmediateOwnerName & "." & strName & "." & strPropertyName
    End If
End Function

Private Sub BDW_AddToSymbolList(objName As Object, strFullName As String, strPrompt As String, strContent As String)
'******************************************************************************************
'PURPOSE: To fill in the BDW_gudtSymbolData array.
'INPUTS:
'   objName: Object who has the connection
'   strFullName: The Connectioned Property's Full Name
'   strPrompt: The User Prompt for that Property
'   strContent: The Current Data Source, Caption, or Initial Value
'******************************************************************************************
Dim intLengthOfContent As Integer   'Length of strContent
Dim intLengthofDDS As Integer       'Length of the Default Data System
Dim strName As String               'The Object's Name
Dim objMyOwnerObject As Object      'The Owner Object
Dim strMyOwnerName As String        'The Owner Object's Name
Dim intForLoopCounter As Integer    'Counter for the For Loop
Dim intCount As Integer             'The number of Contained Objects the owner has
'The following are used for the method GetConnectionInformation
Dim intMyIndex As Integer           'The connection index
Dim strPropertyName As String       'Returns the name of the property for the connection index
Dim strSource As String             'Returns the Data Source Object Name
Dim strFullyQualifiedSource As String   'Returns the fully qualified Data Source Name
Dim vntSourceObjects                'Returns the array of tokenized object parameters
Dim blnGotPropertyNameAlready As Boolean
Dim lngConnectedCount As Long
Dim ImmediateOwnerName As String        'the immediate owners name
Dim ImmediateOwnerObject As Object      'the immediate owner object
Dim objTempOwner As Object              'temp object to used in the Do...While Loop
Dim strTempOwner As String              'temp string used in the Do...While Loop
Dim LenOfOwner As Integer
Dim strTempName As String

    blnGotPropertyNameAlready = False
    Set objMyOwnerObject = objName.owner
    strMyOwnerName = objMyOwnerObject.Name
    strName = objName.Name
    If objName.ClassName = mVARIABLE_CLASS Then
        blnGotPropertyNameAlready = True
        strTempName = GetMyFamily(objName, "InitialValue")
        strPropertyName = strTempName
    ElseIf objName.ClassName = mOLEOBJECT_CLASS Then
        blnGotPropertyNameAlready = True
        strTempName = GetMyFamily(objName, "Caption")
        strPropertyName = strTempName
    ElseIf objName.ClassName = mTEXT_CLASS Then
         blnGotPropertyNameAlready = True
        strTempName = GetMyFamily(objName, "Caption")
        strPropertyName = strTempName
    ElseIf objName.ClassName = mPEN_CLASS Then
        blnGotPropertyNameAlready = True
        strTempName = GetMyFamily(objName, "Source")
        strPropertyName = strTempName
    'kei072508 iFix5.0 T6617
    ElseIf objName.ClassName = mREALTIMEDS_CLASS Then
        blnGotPropertyNameAlready = True
        strTempName = GetMyFamily(objName, "Source")
        strPropertyName = strTempName
    ElseIf objName.ClassName = mSPCDS_CLASS Then
        blnGotPropertyNameAlready = True
        strTempName = GetMyFamily(objName, "Source")
        strPropertyName = strTempName
    ElseIf objName.ClassName = mHISTDS_CLASS Then
        blnGotPropertyNameAlready = True
        strTempName = GetMyFamily(objName, "Source")
        strPropertyName = strTempName
        
    Else
        'need to get the index of this property for GetConnectionInformation
        'For objects within a group with several properties
        objMyOwnerObject.connectedpropertyCount lngConnectedCount
        If lngConnectedCount > 1 Then
            For intForLoopCounter = 1 To lngConnectedCount
                objMyOwnerObject.Getconnectioninformation intForLoopCounter, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
                'error check for unresolved dynamos created in 2.0
                If TypeName(vntSourceObjects) <> "Empty" Then
                    'error check for unresolved pictures created in 2.0
                    If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                        If vntSourceObjects(0).ClassName <> "COpcDataItem" Then
                            If strName = vntSourceObjects(0).Name Then
                                intMyIndex = intForLoopCounter
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next intForLoopCounter
        ElseIf lngConnectedCount = 1 Then
            intMyIndex = 1
        Else
            'there are no connected properties
            Exit Sub
        End If
        
        objMyOwnerObject.Getconnectioninformation intMyIndex, strPropertyName, strSource, strFullyQualifiedSource, vntSourceObjects
        'error check for unresolved dynamos created in 2.0
        If TypeName(vntSourceObjects) <> "Empty" Then
            'error check for unresolved pictures created in 2.0
            If TypeName(vntSourceObjects(0)) <> "Nothing" Then
                If InStr(1, strFullName, "ToggleSource") > 0 Then
                    If strPropertyName = "" Then
                        blnGotPropertyNameAlready = True
                        strPropertyName = strFullName & ".Blink"
                    Else
                        strPropertyName = strPropertyName & ".Blink"
                    End If
                End If
            End If
        End If
    End If
    'If there is a "." as the first character, remove it.
    If Left(strPropertyName, 1) = "." Then
        strPropertyName = Right(strPropertyName, Len(strPropertyName) - 1)
    End If
    'if GetobjectInformation is not returning anything for a propertyName
    'for animated Pens, need to give it a value.
    If strPropertyName = "" Then
        strPropertyName = strFullName
        If Left(strPropertyName, 1) = "." Then
            blnGotPropertyNameAlready = True
            strPropertyName = Right(strPropertyName, Len(strPropertyName) - 1)
        End If
    End If
    
    If mintSymbolLines < mMAX_SYMBOL_SIZE Then
        With BDW_gudtSymbolData(mintSymbolLines)
            On Error Resume Next
            Set ImmediateOwnerObject = objMyOwnerObject.owner
            ImmediateOwnerName = objMyOwnerObject.owner.Name
            'one object, no group
            If blnGotPropertyNameAlready = False Then
                If strMyOwnerName = mstrParentName Then
                    .strPropertyName = strPropertyName
                'groups within groups
                ElseIf ImmediateOwnerName <> mstrParentName Then
                    strTempOwner = ImmediateOwnerName
                    'Loop through until you hit the Parent Name
                    Do While ImmediateOwnerName <> mstrParentName
                        Set objTempOwner = ImmediateOwnerObject.owner
                        If objTempOwner.Name = mstrParentName Then
                            Exit Do
                        End If
                        strTempOwner = objTempOwner.Name & "." & strTempOwner
                        ImmediateOwnerName = objTempOwner.Name
                        Set ImmediateOwnerObject = objTempOwner
                    Loop
                    ImmediateOwnerName = strTempOwner
                    .strPropertyName = ImmediateOwnerName & "." & strMyOwnerName & "." & strPropertyName
                'objects within groups
                Else
                    LenOfOwner = Len(strMyOwnerName)
                    If Left(strPropertyName, LenOfOwner) <> strMyOwnerName Then
                        .strPropertyName = strMyOwnerName & "." & strPropertyName
                    Else
                        .strPropertyName = strPropertyName
                    End If
                End If
            Else
            'PropertyName was found earlier
                .strPropertyName = strPropertyName
            End If
            .strName = strName
            .strFullName = strFullName
            .strPrompt = strPrompt
            .strContent = strContent
            .strField = ""
            intLengthofDDS = Len(System.DefaultDataSystem)
            If (Left(.strContent, intLengthofDDS) = System.DefaultDataSystem) Then
                intLengthOfContent = Len(.strContent)
                Do While (intLengthOfContent > 0)
                    'hj052504 Check the character after the dot. If the character is a number, that means
                    'that the dot is a decimal point and so we have to keep looking.
                    'If Mid(.strContent, intLengthOfContent, 1) = "." Then
                    If (Mid(.strContent, intLengthOfContent, 1) = ".") And (IsNumeric(Left(.strField, 1)) = False) Then
                        .strField = Mid(.strContent, intLengthOfContent, 1) + .strField
                        .strContent = Left(.strContent, intLengthOfContent - 1)
                        Exit Do
                    Else
                        .strField = Mid(.strContent, intLengthOfContent, 1) + .strField
                        intLengthOfContent = intLengthOfContent - 1
                    End If
                Loop
            End If
        End With
        mintSymbolLines = mintSymbolLines + 1
    End If
    
End Sub
Private Sub BDW_GetObjectToObjectConnections(objMyOwnerObject As Object, strPropertyName As String, strFullName As String, strPrompt As String, strContent As String)
'******************************************************************************************
'PURPOSE: To create an array called BDW_GroupAnimsArray, and fill it with all of the Group
'Animation's Property Names.
'INPUTS:
'   objMyOwnerObject: Name of the Group Object
'   strPropertyName: The name of the connected Property
'******************************************************************************************
    Dim strMyOwnerName As String        'The Owner Objects Name
    Dim ImmediateOwnerName As String        'the immediate owners name
    Dim ImmediateOwnerObject As Object      'the immediate owner object
    Dim objTempOwner As Object              'temp object to used in the Do...While Loop
    Dim strTempOwner As String              'temp string used in the Do...While Loop
    Dim intLengthOfContent As Integer   'Length of strContent
    Dim intLengthofDDS As Integer       'Length of the Default Data System

    strMyOwnerName = objMyOwnerObject.Name
    'if the parent is not the same as the main object then
    If strMyOwnerName <> mstrParentName Then
        Set ImmediateOwnerObject = objMyOwnerObject.owner
        ImmediateOwnerName = objMyOwnerObject.owner.Name
        'groups within groups
        If (ImmediateOwnerName <> mstrParentName) And (ImmediateOwnerName <> mstrPictureName) Then
            strTempOwner = ImmediateOwnerName
            'Loop through until you hit the Parent Name
            Do While ImmediateOwnerName <> mstrParentName
                Set objTempOwner = ImmediateOwnerObject.owner
                If objTempOwner.Name = mstrParentName Then
                    Exit Do
                End If
                strTempOwner = objTempOwner.Name & "." & strTempOwner
                ImmediateOwnerName = objTempOwner.Name
                Set ImmediateOwnerObject = objTempOwner
            Loop
            ImmediateOwnerName = strTempOwner
            strPropertyName = ImmediateOwnerName & "." & strMyOwnerName & "." & strPropertyName
        Else
            strPropertyName = strMyOwnerName & "." & strPropertyName
        End If
    End If

    If mintSymbolLines < mMAX_SYMBOL_SIZE Then
        With BDW_gudtSymbolData(mintSymbolLines)
            On Error Resume Next
            .strPropertyName = strPropertyName
            .strName = strMyOwnerName
            .strFullName = strFullName
            .strPrompt = strPrompt
            .strContent = strContent
            .strField = ""
            intLengthofDDS = Len(System.DefaultDataSystem)
            If (Left(.strContent, intLengthofDDS) = System.DefaultDataSystem) Then
                intLengthOfContent = Len(.strContent)
                Do While (intLengthOfContent > 0)
                    If Mid(.strContent, intLengthOfContent, 1) = "." Then
                        .strField = Mid(.strContent, intLengthOfContent, 1) + .strField
                        .strContent = Left(.strContent, intLengthOfContent - 1)
                        Exit Do
                    Else
                        .strField = Mid(.strContent, intLengthOfContent, 1) + .strField
                        intLengthOfContent = intLengthOfContent - 1
                    End If
                Loop
            End If
        End With
        mintSymbolLines = mintSymbolLines + 1
    End If
    
End Sub
Private Sub BDW_PrepareFrmEditDynamo(strObjectName As String)
'******************************************************************************************
'PURPOSE: To prepare the EditDynamo form, add its name to the Name field, and
'check whether it will need the scroll bar enabled. Then calls BDW_FillInEditDynamoForm to fill in
'the fields in the form.
'INPUTS:
'   strObjectName: Name of the Selected Object
'******************************************************************************************
    Dim MyMax As Integer
    Dim intTempMyMax As Integer
    
    frmEditDynamo.mblnCancel = False
    frmEditDynamo.txtObjName = strObjectName
    
    BDW_InitializeEditDynamoArray
    BDW_FillInEditDynamoArrayWithPromptData
    BDW_FillInEditDynamoForm (-1)
    gintNumberOfUniquePrompts = gintTempNumberOfUniquePrompts
    
    If gintNumberOfUniquePrompts > 8 Then
        'hj082101 Changed the way of calculating frmEditDynamo.scrLine.Max
        'MyMax = gintNumberOfUniquePrompts - 7
        'frmEditDynamo.scrLine.Visible = True
        'If BDW_mintEDANumOfIndeces > gintNumberOfUniquePrompts Then
        '    intTempMyMax = BDW_mintEDANumOfIndeces - gintNumberOfUniquePrompts
        'End If
        'this is true if the user gives half or less of their connected properties
        'a prompt
        'If intTempMyMax > MyMax Then
        '   MyMax = intTempMyMax
        'End If
        
        'If (gintNumberOfUniquePrompts Mod 2 = 1) Then
            'if odd number
        '    frmEditDynamo.scrLine.Max = MyMax
        'Else
            'if even number
        '    frmEditDynamo.scrLine.Max = MyMax - 1  'even number needs one less
        'End If
        frmEditDynamo.scrLine.Visible = True
        frmEditDynamo.scrLine.Max = BDW_mintEDANumOfIndeces - 7
    Else
        frmEditDynamo.scrLine.Visible = False
    End If
    If mobjParentObject.ClassName = "Dynamo" Then ' Need to be careful since we still support the old type of Dynamo
        frmEditDynamo.txtboxDynamoDesc.MaxLength = mobjParentObject.Max_Dynamo_Desc_Length
        frmEditDynamo.txtboxDynamoDesc.Text = mobjParentObject.Dynamo_Description
    Else
        frmEditDynamo.txtboxDynamoDesc.Visible = False
        frmEditDynamo.lblDynDesc.Visible = False
        frmEditDynamo.lblDynamoProperty.Top = frmEditDynamo.lblDynamoProperty.Top - 25
        frmEditDynamo.lblCurrentSetting.Top = frmEditDynamo.lblCurrentSetting.Top - 25
        frmEditDynamo.fra1.Top = frmEditDynamo.fra1.Top - 25
        frmEditDynamo.cmdOK.Top = frmEditDynamo.cmdOK.Top - 25
        frmEditDynamo.cmdCancel.Top = frmEditDynamo.cmdCancel.Top - 25
        frmEditDynamo.cmdHelp.Top = frmEditDynamo.cmdHelp.Top - 25
        frmEditDynamo.Height = (frmEditDynamo.Height - 25)
    End If
End Sub
Sub BDW_InitializeCreateDynamoArray()
'******************************************************************************************
'PURPOSE: We do not want information saved until the developer hits OK within the CreateDynamo
'form. So, we create an array called the CreateDynamo Array, with 3 columns for the Property Name,
'Current Setting, and User Prompt. This information will be saved back into the BDW_gudtSymbolData
'only if the Developer clicks OK within the CreateDynamo form.
'The array size is based upon how many Properties the Dynamo has.
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop, and Index for the BDW_gudtSymbolData Data Type.

    For intIndex = 0 To 200
        With BDW_gudtSymbolData(intIndex)
            If .strName = "" Then
                Exit For
            End If
        End With
    Next intIndex
    'This is a module level variable set for the number of indeces in the BDW_CreateDynamoArray.
    'The we Redim the array BDW_CreateDynamoArray to that size.
    If intIndex > 0 Then
        BDW_mintCDANumOfIndeces = intIndex - 1
        ReDim BDW_CreateDynamoArray(BDW_mintCDANumOfIndeces, 2) As String
    Else
        BDW_mintCDANumOfIndeces = -1
    End If
End Sub

Sub BDW_InitializeEditDynamoArray()
'******************************************************************************************
'PURPOSE: We do not want information saved until the developer hits OK within the EditDynamo
'form. So,we create an array called the EditDynamo Array, with 3 columns for the User Prompt,
'the Current Setting, and whether it will need an expression editor or a textbox. This information will be saved back into the BDW_mudtPromptData
'only if the Developer clicks OK within the EditDynamo form.
'The array size is based upon how many Properties the Dynamo has.
'******************************************************************************************
    'This is a module level variable set for the number of indeces in the BDW_CreateDynamoArray.
    'The we Redim the array BDW_CreateDynamoArray to that size.
    BDW_mintEDANumOfIndeces = mintPromptLines - 1
    ReDim BDW_EditDynamoArray(BDW_mintEDANumOfIndeces, 2) As String
    
End Sub
Sub BDW_FillInCreateDynamoArrayFromForm(intIndex As Integer, txtbox As Control)
'******************************************************************************************
'PURPOSE: This is called from the CreateDynamo form.
'Each time the Developer changes their User Prompt, it is stored in the
'proper index within the BDW_CreateDynamoArray
'INPUTS:
'   intIndex: The textbox number that was changed
'   txtbox: The control that was changed
'******************************************************************************************
    Dim intArrayIndex As Integer    'Array index number
    
    intArrayIndex = intIndex + frmCreateDynamo.scrLine.Value
    BDW_CreateDynamoArray(intArrayIndex, 2) = txtbox.Text

End Sub

Sub BDW_FillInEditDynamoArrayFromForm(intIndex As Integer, txtbox As Control, txtPrompt As Control)
'******************************************************************************************
'PURPOSE: This is called from the EditDynamo form.
'Each time the Developer changes their Property's Current Setting, it is stored in the
'proper index within the BDW_EditDynamoArray
'INPUTS:
'   intIndex: The textbox number that was changed
'   blnIsEEControl: True if the control is an Expression Editor Control
'   txtbox: The control that was changed
'******************************************************************************************
    Dim intArrayIndex As Integer    'Array index number
    Dim blnBeenHere As Boolean      'Boolean Set to True If Partial Substitution is used and more than one object has the Prompt Passed in.
    Dim intLengthOfText As Integer  'Length of the New Current Setting
    Dim strField As String           'Temp string to hold the field
    Dim strExpression As String     'hj020504
    
    blnBeenHere = False
    strField = ""
    
    For intArrayIndex = 0 To BDW_mintEDANumOfIndeces
        If txtPrompt.Caption = BDW_EditDynamoArray(intArrayIndex, 0) Then
            If BDW_EditDynamoArray(intArrayIndex, 2) = "True" And (blnBeenHere = False) Then
                'Is an expression editor control, no partial sub, and field still on it
                If TypeName(txtbox) <> "TextBox" Then
                    BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.EditText
                Else
                    BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.Text
                End If
            ElseIf BDW_EditDynamoArray(intArrayIndex, 2) = "True" And (blnBeenHere = True) Then
                'Is an expression editor control, no partial sub, and field was previously removed
                If TypeName(txtbox) <> "TextBox" Then
                    BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.EditText & strField
                Else
                    BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.Text
                End If
            ElseIf BDW_EditDynamoArray(intArrayIndex, 2) = "TrueTrue" And (blnBeenHere = False) Then
                'Only gets in here once when partial sub is used.
                'We only need to strip the Field off the Source once, then re-use it.
                'Here we remove the Field because partial sub is used and we don't need to save the current Field into the Array.
                blnBeenHere = True
                intLengthOfText = Len(txtbox.EditText)
                 If txtbox.EditText Like "*.*.*.*" Then
                'If InStr(1, txtbox.EditText, ".") > 0 Then (JES 1/25/00 replaced with above If to resolve case197603-tab off field and text gets stripped)
                    Do While (intLengthOfText > 0)
                        If Mid(txtbox.EditText, intLengthOfText, 1) = "." Then
                            strField = Mid(txtbox.EditText, intLengthOfText, 1) + strField
                            txtbox.EditText = Left(txtbox.EditText, intLengthOfText - 1)
                            Exit Do
                        Else
                            strField = Mid(txtbox.EditText, intLengthOfText, 1) + strField
                            intLengthOfText = intLengthOfText - 1
                        End If
                    Loop
                End If
                If TypeName(txtbox) <> "TextBox" Then
                    'hj020504 Clarify #283379 Call BDW_SubstituteNodeTag to do substitution
                    'BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.EditText
                    strExpression = BDW_EditDynamoArray(intArrayIndex, 1)
                    BDW_SubstituteNodeTag strExpression, txtbox.EditText
                    BDW_EditDynamoArray(intArrayIndex, 1) = strExpression
                Else
                    BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.Text
                End If
            ElseIf BDW_EditDynamoArray(intArrayIndex, 2) = "TrueTrue" And (blnBeenHere = True) Then
                'Is an expression editor control, using partial sub, and field was already removed
                'hj020504 Clarify #283379 Call BDW_SubstituteNodeTag to do substitution
                'BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.EditText
                strExpression = BDW_EditDynamoArray(intArrayIndex, 1)
                BDW_SubstituteNodeTag strExpression, txtbox.EditText
                BDW_EditDynamoArray(intArrayIndex, 1) = strExpression
            Else    'Is a text box
                'Check to see if the prompt is in a text box or a EE
                If TypeName(txtbox) <> "TextBox" Then
                    BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.EditText
                Else
                    BDW_EditDynamoArray(intArrayIndex, 1) = txtbox.Text
                End If
            End If
        End If
    Next intArrayIndex
    
End Sub

Sub BDW_FillInCreateDynamoArrayWithSymbolData()
'******************************************************************************************
'PURPOSE: To retrieve the Property Name, Current Value, and User Prompt from the Data Type
'Array BDW_gudtSymbolData, and store it into the same index number in the BDW_CreateDynamoArray.
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    
    If BDW_mintCDANumOfIndeces >= 0 Then
        For intIndex = 0 To BDW_mintCDANumOfIndeces
            With BDW_gudtSymbolData(intIndex)
                If Left(.strPropertyName, 1) = "." Then
                    BDW_CreateDynamoArray(intIndex, 0) = Right(.strPropertyName, Len(.strPropertyName) - 1)
                Else
                    BDW_CreateDynamoArray(intIndex, 0) = .strPropertyName
                End If
                BDW_CreateDynamoArray(intIndex, 1) = .strContent & .strField
                BDW_CreateDynamoArray(intIndex, 2) = .strPrompt
            End With
        Next intIndex
    End If
End Sub

Sub BDW_FillInEditDynamoArrayWithPromptData()
'******************************************************************************************
'PURPOSE: To retrieve the User Prompt and Current Setting from the Data Type
'Array BDW_mudtPromptData, and store it into the same index number in the BDW_EditDynamoArray.
'******************************************************************************************
Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_mudtPromptData Data Type.
    
    For intIndex = 0 To BDW_mintEDANumOfIndeces
        With BDW_mudtPromptData(intIndex)
            BDW_EditDynamoArray(intIndex, 0) = .strPrompt
            If (Right(BDW_gudtSymbolData(intIndex).strFullName, 6) = "Source") Then
                'needs an EE
                If .blnIsUsingSubstitution = True Then
                    'Partial subst. is used, no need to show the field
                    BDW_EditDynamoArray(intIndex, 2) = "TrueTrue"
                    BDW_EditDynamoArray(intIndex, 1) = .strContent
                Else
                    BDW_EditDynamoArray(intIndex, 2) = "True"
                    BDW_EditDynamoArray(intIndex, 1) = .strContent & .strField
                End If
            Else
                'needs a text box
                BDW_EditDynamoArray(intIndex, 2) = "False"
                BDW_EditDynamoArray(intIndex, 1) = .strContent & .strField
            End If
        End With
    Next intIndex
    
End Sub

Sub BDW_FillInEditDynamoForm(Optional Value As Integer)
'******************************************************************************************
'PURPOSE: To fill in the necessary fields in the EditDynamo form with information retrieved
'from BDW_EditDynamoArray.
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_mudtPromptData Data Type.
    Dim intKeepingCount As Integer  'Counter for the Case Statement
    Dim blnPromptIsAlreadyShownOrBlank As Boolean  'Boolean set to True if the User Prompt is already shown in the form
    Dim intCount As Integer 'Counter for Second For Loop
    Dim intEDAIndex As Integer  'EditDynamoArray Index
    Dim intScrollBarValue As Integer    'The Scroll bars value
    
    intKeepingCount = 0
    gintTempNumberOfUniquePrompts = 0
    
    On Error GoTo ErrorHandler
    ''We want to make sure this loops enough times for the form to get filled in, so
    'use the max. number of indeces allowed for the BDW_mudtPromptData
        For intIndex = 0 To mMAX_PROMPT_SIZE
        'Index number in EDArray is the scroll position of the scroll bar plus the index from the for loop
        If Value <> -1 Then
            If Value > 8 Then
                intScrollBarValue = gintNumberOfUniquePrompts - 8
            Else
                intScrollBarValue = 0
            End If
            frmEditDynamo.scrLine.Value = intScrollBarValue
            intEDAIndex = intIndex + frmEditDynamo.scrLine.Value
        Else
            intEDAIndex = intIndex + frmEditDynamo.scrLine.Value
        End If
        blnPromptIsAlreadyShownOrBlank = False
        'Once all of the indeces have been displayed in the form, we want to set the visible property of the remaining
        'controls to False. This If stmt checks that.
        If intEDAIndex <= BDW_mintEDANumOfIndeces Then
            'If first index is blank need to do this because it will get skipped over otherwise
            'The For Loop below disregards Index 0
            If intEDAIndex = 0 And BDW_EditDynamoArray(intEDAIndex, 1) = "" Then
                If BDW_EditDynamoArray(intEDAIndex, 0) = "" Then
                    blnPromptIsAlreadyShownOrBlank = True
                End If
            End If
            'Check to see if Prompt is already shown by looping through the array. If the current
            'index matches one previously shown then set blnPromptIsAlreadyShownOrBlank to True
            For intCount = 0 To intEDAIndex - 1
                If BDW_EditDynamoArray(intEDAIndex, 0) = BDW_mudtPromptData(intCount).strPrompt Or BDW_EditDynamoArray(intEDAIndex, 0) = "" Then
                    blnPromptIsAlreadyShownOrBlank = True
                    Exit For
                End If
            Next intCount
            'If the prompt is already shown then go to the AlreadyVisibleOrBlank marker
            If blnPromptIsAlreadyShownOrBlank Then
                GoTo AlreadyVisibleOrBlank
            Else
                'we only need to count these once
                'Counter for how many unique prompts there are.
                gintTempNumberOfUniquePrompts = gintTempNumberOfUniquePrompts + 1
            End If
            Select Case intKeepingCount
            Case 0:
                If BDW_EditDynamoArray(intEDAIndex, 2) = True Or BDW_EditDynamoArray(intEDAIndex, 2) = "TrueTrue" Then
                    frmEditDynamo.txtContent0.Visible = True
                    frmEditDynamo.txtContent0.EditText = BDW_EditDynamoArray(intEDAIndex, 1)
                    If (IsIHistorian() = False) Then
                        frmEditDynamo.txtContent0.ShowHistoricalTab = False
                    End If
                    frmEditDynamo.txtContent0a.Visible = False
                Else
                    frmEditDynamo.txtContent0a.Visible = True
                    frmEditDynamo.txtContent0a.Text = BDW_EditDynamoArray(intEDAIndex, 1)
                    frmEditDynamo.txtContent0.Visible = False
                End If
                frmEditDynamo.txtPrompt0.Visible = True
                frmEditDynamo.txtPrompt0.Caption = BDW_EditDynamoArray(intEDAIndex, 0)
                'JPB031403  Tracker #422 add tooltip if text may be too long for display
                If (Len(frmEditDynamo.txtPrompt0.Caption) > 30) Then
                    frmEditDynamo.txtPrompt0.ControlTipText = frmEditDynamo.txtPrompt0.Caption
                Else
                    frmEditDynamo.txtPrompt0.ControlTipText = ""
                End If
                intKeepingCount = intKeepingCount + 1
            Case 1:
                If BDW_EditDynamoArray(intEDAIndex, 2) = True Or BDW_EditDynamoArray(intEDAIndex, 2) = "TrueTrue" Then
                    frmEditDynamo.txtContent1.Visible = True
                    frmEditDynamo.txtContent1.EditText = BDW_EditDynamoArray(intEDAIndex, 1)
                    If (IsIHistorian() = False) Then
                        frmEditDynamo.txtContent1.ShowHistoricalTab = False
                    End If
                    frmEditDynamo.txtContent1a.Visible = False
                Else
                    frmEditDynamo.txtContent1a.Visible = True
                    frmEditDynamo.txtContent1a.Text = BDW_EditDynamoArray(intEDAIndex, 1)
                    frmEditDynamo.txtContent1.Visible = False
                End If
                frmEditDynamo.txtPrompt1.Visible = True
                frmEditDynamo.txtPrompt1.Caption = BDW_EditDynamoArray(intEDAIndex, 0)
                'JPB031403  Tracker #422 add tooltip if text may be too long for display
                If (Len(frmEditDynamo.txtPrompt1.Caption) > 30) Then
                    frmEditDynamo.txtPrompt1.ControlTipText = frmEditDynamo.txtPrompt1.Caption
                Else
                    frmEditDynamo.txtPrompt1.ControlTipText = ""
                End If
                intKeepingCount = intKeepingCount + 1
            Case 2:
                If BDW_EditDynamoArray(intEDAIndex, 2) = True Or BDW_EditDynamoArray(intEDAIndex, 2) = "TrueTrue" Then
                    frmEditDynamo.txtContent2.Visible = True
                    frmEditDynamo.txtContent2.EditText = BDW_EditDynamoArray(intEDAIndex, 1)
                    If (IsIHistorian() = False) Then
                        frmEditDynamo.txtContent2.ShowHistoricalTab = False
                    End If
                    frmEditDynamo.txtContent2a.Visible = False
                Else
                    frmEditDynamo.txtContent2a.Visible = True
                    frmEditDynamo.txtContent2a.Text = BDW_EditDynamoArray(intEDAIndex, 1)
                    frmEditDynamo.txtContent2.Visible = False
                End If
                frmEditDynamo.txtPrompt2.Visible = True
                frmEditDynamo.txtPrompt2.Caption = BDW_EditDynamoArray(intEDAIndex, 0)
                'JPB031403  Tracker #422 add tooltip if text may be too long for display
                If (Len(frmEditDynamo.txtPrompt2.Caption) > 30) Then
                    frmEditDynamo.txtPrompt2.ControlTipText = frmEditDynamo.txtPrompt2.Caption
                Else
                    frmEditDynamo.txtPrompt2.ControlTipText = ""
                End If
                intKeepingCount = intKeepingCount + 1
            Case 3:
                If BDW_EditDynamoArray(intEDAIndex, 2) = True Or BDW_EditDynamoArray(intEDAIndex, 2) = "TrueTrue" Then
                    frmEditDynamo.txtContent3.Visible = True
                    frmEditDynamo.txtContent3.EditText = BDW_EditDynamoArray(intEDAIndex, 1)
                    If (IsIHistorian() = False) Then
                        frmEditDynamo.txtContent3.ShowHistoricalTab = False
                    End If
                    frmEditDynamo.txtContent3a.Visible = False
                Else
                    frmEditDynamo.txtContent3a.Visible = True
                    frmEditDynamo.txtContent3a.Text = BDW_EditDynamoArray(intEDAIndex, 1)
                    frmEditDynamo.txtContent3.Visible = False
                End If
                frmEditDynamo.txtPrompt3.Visible = True
                frmEditDynamo.txtPrompt3.Caption = BDW_EditDynamoArray(intEDAIndex, 0)
                'JPB031403  Tracker #422 add tooltip if text may be too long for display
                If (Len(frmEditDynamo.txtPrompt3.Caption) > 30) Then
                    frmEditDynamo.txtPrompt3.ControlTipText = frmEditDynamo.txtPrompt3.Caption
                Else
                    frmEditDynamo.txtPrompt3.ControlTipText = ""
                End If
                intKeepingCount = intKeepingCount + 1
            Case 4:
                If BDW_EditDynamoArray(intEDAIndex, 2) = True Or BDW_EditDynamoArray(intEDAIndex, 2) = "TrueTrue" Then
                    frmEditDynamo.txtContent4.Visible = True
                    frmEditDynamo.txtContent4.EditText = BDW_EditDynamoArray(intEDAIndex, 1) '.strContent & .strField
                    If (IsIHistorian() = False) Then
                        frmEditDynamo.txtContent4.ShowHistoricalTab = False
                    End If
                    frmEditDynamo.txtContent4a.Visible = False
                Else
                    frmEditDynamo.txtContent4a.Visible = True
                    frmEditDynamo.txtContent4a.Text = BDW_EditDynamoArray(intEDAIndex, 1)
                    frmEditDynamo.txtContent4.Visible = False
                End If
                frmEditDynamo.txtPrompt4.Visible = True
                frmEditDynamo.txtPrompt4.Caption = BDW_EditDynamoArray(intEDAIndex, 0)
                'JPB031403  Tracker #422 add tooltip if text may be too long for display
                If (Len(frmEditDynamo.txtPrompt4.Caption) > 30) Then
                    frmEditDynamo.txtPrompt4.ControlTipText = frmEditDynamo.txtPrompt4.Caption
                Else
                    frmEditDynamo.txtPrompt4.ControlTipText = ""
                End If
                intKeepingCount = intKeepingCount + 1
            Case 5:
                If BDW_EditDynamoArray(intEDAIndex, 2) = True Or BDW_EditDynamoArray(intEDAIndex, 2) = "TrueTrue" Then
                    frmEditDynamo.txtContent5.Visible = True
                    frmEditDynamo.txtContent5.EditText = BDW_EditDynamoArray(intEDAIndex, 1)
                    If (IsIHistorian() = False) Then
                        frmEditDynamo.txtContent5.ShowHistoricalTab = False
                    End If
                    frmEditDynamo.txtContent5a.Visible = False
                Else
                    frmEditDynamo.txtContent5a.Visible = True
                    frmEditDynamo.txtContent5a.Text = BDW_EditDynamoArray(intEDAIndex, 1)
                    frmEditDynamo.txtContent5.Visible = False
                End If
                frmEditDynamo.txtPrompt5.Visible = True
                frmEditDynamo.txtPrompt5.Caption = BDW_EditDynamoArray(intEDAIndex, 0)
                'JPB031403  Tracker #422 add tooltip if text may be too long for display
                If (Len(frmEditDynamo.txtPrompt5.Caption) > 30) Then
                    frmEditDynamo.txtPrompt5.ControlTipText = frmEditDynamo.txtPrompt5.Caption
                Else
                    frmEditDynamo.txtPrompt5.ControlTipText = ""
                End If
                intKeepingCount = intKeepingCount + 1
            Case 6:
                If BDW_EditDynamoArray(intEDAIndex, 2) = True Or BDW_EditDynamoArray(intEDAIndex, 2) = "TrueTrue" Then
                    frmEditDynamo.txtContent6.Visible = True
                    frmEditDynamo.txtContent6.EditText = BDW_EditDynamoArray(intEDAIndex, 1)
                    If (IsIHistorian() = False) Then
                        frmEditDynamo.txtContent6.ShowHistoricalTab = False
                    End If
                    frmEditDynamo.txtContent6a.Visible = False
                Else
                    frmEditDynamo.txtContent6a.Visible = True
                    frmEditDynamo.txtContent6a.Text = BDW_EditDynamoArray(intEDAIndex, 1)
                    frmEditDynamo.txtContent6.Visible = False
                End If
                frmEditDynamo.txtPrompt6.Visible = True
                frmEditDynamo.txtPrompt6.Caption = BDW_EditDynamoArray(intEDAIndex, 0)
                'JPB031403  Tracker #422 add tooltip if text may be too long for display
                If (Len(frmEditDynamo.txtPrompt6.Caption) > 30) Then
                    frmEditDynamo.txtPrompt6.ControlTipText = frmEditDynamo.txtPrompt6.Caption
                Else
                    frmEditDynamo.txtPrompt6.ControlTipText = ""
                End If
                intKeepingCount = intKeepingCount + 1
            Case 7:
                If BDW_EditDynamoArray(intEDAIndex, 2) = True Or BDW_EditDynamoArray(intEDAIndex, 2) = "TrueTrue" Then
                    frmEditDynamo.txtContent7.Visible = True
                    frmEditDynamo.txtContent7.EditText = BDW_EditDynamoArray(intEDAIndex, 1)
                    If (IsIHistorian() = False) Then
                        frmEditDynamo.txtContent7.ShowHistoricalTab = False
                    End If
                    frmEditDynamo.txtContent7a.Visible = False
                Else
                    frmEditDynamo.txtContent7a.Visible = True
                    frmEditDynamo.txtContent7a.Text = BDW_EditDynamoArray(intEDAIndex, 1)
                    frmEditDynamo.txtContent7.Visible = False
                End If
                frmEditDynamo.txtPrompt7.Visible = True
                frmEditDynamo.txtPrompt7.Caption = BDW_EditDynamoArray(intEDAIndex, 0)
                'JPB031403  Tracker #422 add tooltip if text may be too long for display
                If (Len(frmEditDynamo.txtPrompt7.Caption) > 30) Then
                    frmEditDynamo.txtPrompt7.ControlTipText = frmEditDynamo.txtPrompt7.Caption
                Else
                    frmEditDynamo.txtPrompt7.ControlTipText = ""
                End If
                intKeepingCount = intKeepingCount + 1
            End Select
        Else
            Select Case intKeepingCount
            Case 0:
                frmEditDynamo.txtContent0a.Visible = False
                frmEditDynamo.txtContent0.Visible = False
                frmEditDynamo.txtPrompt0.Visible = False
                intKeepingCount = intKeepingCount + 1
            Case 1:
                If frmEditDynamo.scrLine.Value = 0 Then
                    frmEditDynamo.scrLine.Visible = False
                End If
                frmEditDynamo.txtContent1a.Visible = False
                frmEditDynamo.txtContent1.Visible = False
                frmEditDynamo.txtPrompt1.Visible = False
                intKeepingCount = intKeepingCount + 1
            Case 2:
                If frmEditDynamo.scrLine.Value = 0 Then
                    frmEditDynamo.scrLine.Visible = False
                End If
                frmEditDynamo.txtContent2a.Visible = False
                frmEditDynamo.txtContent2.Visible = False
                frmEditDynamo.txtPrompt2.Visible = False
                intKeepingCount = intKeepingCount + 1
            Case 3:
                If frmEditDynamo.scrLine.Value = 0 Then
                    frmEditDynamo.scrLine.Visible = False
                End If
                frmEditDynamo.txtContent3a.Visible = False
                frmEditDynamo.txtContent3.Visible = False
                frmEditDynamo.txtPrompt3.Visible = False
                intKeepingCount = intKeepingCount + 1
            Case 4:
                If frmEditDynamo.scrLine.Value = 0 Then
                    frmEditDynamo.scrLine.Visible = False
                End If
                frmEditDynamo.txtContent4a.Visible = False
                frmEditDynamo.txtContent4.Visible = False
                frmEditDynamo.txtPrompt4.Visible = False
                intKeepingCount = intKeepingCount + 1
            Case 5:
                If frmEditDynamo.scrLine.Value = 0 Then
                    frmEditDynamo.scrLine.Visible = False
                End If
                frmEditDynamo.txtContent5a.Visible = False
                frmEditDynamo.txtContent5.Visible = False
                frmEditDynamo.txtPrompt5.Visible = False
                intKeepingCount = intKeepingCount + 1
            Case 6:
                If frmEditDynamo.scrLine.Value = 0 Then
                    frmEditDynamo.scrLine.Visible = False
                End If
                frmEditDynamo.txtContent6a.Visible = False
                frmEditDynamo.txtContent6.Visible = False
                frmEditDynamo.txtPrompt6.Visible = False
                intKeepingCount = intKeepingCount + 1
            Case 7:
                If frmEditDynamo.scrLine.Value = 0 Then
                    frmEditDynamo.scrLine.Visible = False
                End If
                frmEditDynamo.txtContent7a.Visible = False
                frmEditDynamo.txtContent7.Visible = False
                frmEditDynamo.txtPrompt7.Visible = False
                intKeepingCount = intKeepingCount + 1
            End Select
        
AlreadyVisibleOrBlank:
        End If
    Next intIndex
    Exit Sub
        
ErrorHandler:
    HandleError
End Sub

Private Sub BDW_PrepareFrmCreateDynamo(strObjectName As String, Optional bShowShortDialog As Boolean = False)
'******************************************************************************************
'PURPOSE: To prepare the CreateDynamo form, add its name to the Name field, and
'check whether it will need the scroll bar enabled. Then initializes the BDW_CreateDynamoArray,
'Fills in the Array with information from the Data Type Array BDW_gudtSymbolData, then
'uses that information to fill in the CreateDynamo form.
'INPUTS:
'   strObjectName: Name of the Selected Object
'******************************************************************************************
    If False = bShowShortDialog Then
        
        frmCreateDynamo.mblnCancel = False
        frmCreateDynamo.txtObjName = strObjectName
        
        If mintSymbolLines > 10 Then
            frmCreateDynamo.scrLine.Visible = True
            frmCreateDynamo.scrLine.Max = mintSymbolLines - 10
        Else
            frmCreateDynamo.scrLine.Visible = False
        End If
        
        BDW_InitializeCreateDynamoArray
        BDW_FillInCreateDynamoArrayWithSymbolData
        BDW_FillInCreateDynamoForm
    Else
        Dim strTitle As String      'Title for Error message box
        Dim strHelpPath As String   'The Path where the help file is located
        strHelpPath = System.NlsPath & "\BuildDynamoWizard.hlp"
        strTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
        frmCreateDynamo.Caption = strTitle
        frmCreateDynamo.fra1.Visible = False
        frmCreateDynamo.lblUserPrompt.Visible = False
        frmCreateDynamo.lblCurrentSetting.Visible = False
        frmCreateDynamo.lblPropertyName.Visible = False
        frmCreateDynamo.cmdOK.Top = 88
        frmCreateDynamo.cmdOK.Left = 162
        frmCreateDynamo.cmdHelp.Top = 88
        frmCreateDynamo.cmdHelp.Left = 310 'kei040209 T7133 270
        frmCreateDynamo.cmdCancel.Top = 88
        frmCreateDynamo.cmdCancel.Left = 236 'kei040209 T7133 216
        frmCreateDynamo.Width = 470
        frmCreateDynamo.Height = 150
    End If
    
    frmCreateDynamo.txtboxDynamoDesc.MaxLength = mobjParentObject.Max_Dynamo_Desc_Length

End Sub

Sub BDW_FillInCreateDynamoForm()
'******************************************************************************************
'PURPOSE: To retrieve information from the BDW_CreateDynamoArray, and fill in the fields
'in the CreateDynamo form with that information.
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    Dim intCDAIndex As Integer  'The CreateDynamoArray Index number
    
    On Error GoTo ErrorHandler
    frmCreateDynamo.lblObjectProperty0.Visible = True
    frmCreateDynamo.lblCurrentSetting0.Visible = True
    frmCreateDynamo.txtPrompt0.Visible = True
    frmCreateDynamo.lblObjectProperty1.Visible = True
    frmCreateDynamo.lblCurrentSetting1.Visible = True
    frmCreateDynamo.txtPrompt1.Visible = True
    frmCreateDynamo.lblObjectProperty2.Visible = True
    frmCreateDynamo.lblCurrentSetting2.Visible = True
    frmCreateDynamo.txtPrompt2.Visible = True
    frmCreateDynamo.lblObjectProperty3.Visible = True
    frmCreateDynamo.lblCurrentSetting3.Visible = True
    frmCreateDynamo.txtPrompt3.Visible = True
    frmCreateDynamo.lblObjectProperty4.Visible = True
    frmCreateDynamo.lblCurrentSetting4.Visible = True
    frmCreateDynamo.txtPrompt4.Visible = True
    frmCreateDynamo.lblObjectProperty5.Visible = True
    frmCreateDynamo.lblCurrentSetting5.Visible = True
    frmCreateDynamo.txtPrompt5.Visible = True
    frmCreateDynamo.lblObjectProperty6.Visible = True
    frmCreateDynamo.lblCurrentSetting6.Visible = True
    frmCreateDynamo.txtPrompt6.Visible = True
    frmCreateDynamo.lblObjectProperty7.Visible = True
    frmCreateDynamo.lblCurrentSetting7.Visible = True
    frmCreateDynamo.txtPrompt7.Visible = True
    frmCreateDynamo.lblObjectProperty8.Visible = True
    frmCreateDynamo.lblCurrentSetting8.Visible = True
    frmCreateDynamo.txtPrompt8.Visible = True
    frmCreateDynamo.lblObjectProperty9.Visible = True
    frmCreateDynamo.lblCurrentSetting9.Visible = True
    frmCreateDynamo.txtPrompt9.Visible = True
    
    For intIndex = 0 To 9 'This 0-9 is how many textboxes can be visible at one time in the CD form
        intCDAIndex = intIndex + frmCreateDynamo.scrLine.Value 'Index number in CDArray
        If (intCDAIndex <= BDW_mintCDANumOfIndeces) And (BDW_mintCDANumOfIndeces >= 0) Then
        'Once there are no more indeces in
        'the CreateDynamo array (intCDAIndex > BDW_mintCDANumOfIndeces),
        'we don't want to show any more of the textboxes
            Select Case intIndex
            Case 0:
                frmCreateDynamo.lblObjectProperty0.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting0.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt0.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 1:
                frmCreateDynamo.lblObjectProperty1.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting1.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt1.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 2:
                frmCreateDynamo.lblObjectProperty2.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting2.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt2.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 3:
                frmCreateDynamo.lblObjectProperty3.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting3.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt3.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 4:
                frmCreateDynamo.lblObjectProperty4.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting4.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt4.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 5:
                frmCreateDynamo.lblObjectProperty5.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting5.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt5.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 6:
                frmCreateDynamo.lblObjectProperty6.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting6.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt6.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 7:
                frmCreateDynamo.lblObjectProperty7.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting7.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt7.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 8:
                frmCreateDynamo.lblObjectProperty8.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting8.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt8.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            Case 9:
                frmCreateDynamo.lblObjectProperty9.Caption = BDW_CreateDynamoArray(intCDAIndex, 0)
                frmCreateDynamo.lblCurrentSetting9.Text = BDW_CreateDynamoArray(intCDAIndex, 1)
                frmCreateDynamo.txtPrompt9.Text = BDW_CreateDynamoArray(intCDAIndex, 2)
            End Select
        Else
            Select Case intIndex
            Case 0:
                frmCreateDynamo.lblObjectProperty0.Visible = False
                frmCreateDynamo.lblCurrentSetting0.Visible = False
                frmCreateDynamo.txtPrompt0.Visible = False
            Case 1:
                frmCreateDynamo.lblObjectProperty1.Visible = False
                frmCreateDynamo.lblCurrentSetting1.Visible = False
                frmCreateDynamo.txtPrompt1.Visible = False
            Case 2:
                frmCreateDynamo.lblObjectProperty2.Visible = False
                frmCreateDynamo.lblCurrentSetting2.Visible = False
                frmCreateDynamo.txtPrompt2.Visible = False
            Case 3:
                frmCreateDynamo.lblObjectProperty3.Visible = False
                frmCreateDynamo.lblCurrentSetting3.Visible = False
                frmCreateDynamo.txtPrompt3.Visible = False
            Case 4:
                frmCreateDynamo.lblObjectProperty4.Visible = False
                frmCreateDynamo.lblCurrentSetting4.Visible = False
                frmCreateDynamo.txtPrompt4.Visible = False
            Case 5:
                frmCreateDynamo.lblObjectProperty5.Visible = False
                frmCreateDynamo.lblCurrentSetting5.Visible = False
                frmCreateDynamo.txtPrompt5.Visible = False
            Case 6:
                frmCreateDynamo.lblObjectProperty6.Visible = False
                frmCreateDynamo.lblCurrentSetting6.Visible = False
                frmCreateDynamo.txtPrompt6.Visible = False
            Case 7:
                frmCreateDynamo.lblObjectProperty7.Visible = False
                frmCreateDynamo.lblCurrentSetting7.Visible = False
                frmCreateDynamo.txtPrompt7.Visible = False
            Case 8:
                frmCreateDynamo.lblObjectProperty8.Visible = False
                frmCreateDynamo.lblCurrentSetting8.Visible = False
                frmCreateDynamo.txtPrompt8.Visible = False
            Case 9:
                frmCreateDynamo.lblObjectProperty9.Visible = False
                frmCreateDynamo.lblCurrentSetting9.Visible = False
                frmCreateDynamo.txtPrompt9.Visible = False
            End Select
        End If
      Next intIndex

    Exit Sub

ErrorHandler:
    HandleError
End Sub
Function BDW_UpdateGroupName(vntNewGroupName As Variant) As Boolean
'******************************************************************************************
'PURPOSE: To assign the Selected Object its new name. The user can change it in either the
'CreateDynamo or EditDynamo form.
'INPUTS:
'   strNewGroupName: The new name for the Dynamo.
'******************************************************************************************
    Dim strTitle As String
    Dim strError As String
    
    On Error GoTo ErrorHandler
    
    'If the name has changed
    If mobjParentObject.Name <> vntNewGroupName Then
        'if there is a space in the dynamo name
        'JPB031403  Tracker #865  added IsAlpha... check for valid name
        If InStr(1, vntNewGroupName, " ") = 0 And IsAlphanumericStartIsAlpha(CStr(vntNewGroupName)) = True Then
            mobjParentObject.Name = vntNewGroupName
            BDW_UpdateGroupName = True
        Else
            strTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
            'JPB031403  Tracker #18  added another line to the message, explaining that spaces are illegal
            strError = objStrMgr.GetNLSStr(CLng(NLS_mERROR1)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1a)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1e)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1b)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1c)) & Chr(13) & objStrMgr.GetNLSStr(CLng(NLS_mERROR1d))
            MsgBox strError, , strTitle
            BDW_UpdateGroupName = False
        End If
    Else
        BDW_UpdateGroupName = True
    End If
    Exit Function
    
ErrorHandler:
    HandleError
End Function

Function BDW_Update_Object_Desc(sNewText As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sTitle As String
    Dim sError As String
    
    If mobjParentObject.Description <> sNewText Then
        If Len(sNewText) <= 64 Then
            mobjParentObject.Description = sNewText
            BDW_Update_Object_Desc = True
        Else
            sTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
            sError = objStrMgr.GetNLSStr(CLng(NLS_mDYN_DESC_LEN_ERROR)) & CStr(mobjParentObject.Max_Dynamo_Desc_Length)
            MsgBox sError, , sTitle
            BDW_Update_Object_Desc = False
        End If
    Else
        BDW_Update_Object_Desc = True
    End If
    
    Exit Function
ErrorHandler:
    HandleError
End Function
Function BDW_Update_Dynamo_Desc(sNewText As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sTitle As String
    Dim sError As String
    
    If mobjParentObject.Description <> sNewText Then
        If Len(sNewText) <= mobjParentObject.Max_Dynamo_Desc_Length Then
            mobjParentObject.Dynamo_Description = sNewText
            BDW_Update_Dynamo_Desc = True
        Else
            sTitle = objStrMgr.GetNLSStr(CLng(NLS_mTITLE))
            sError = objStrMgr.GetNLSStr(CLng(NLS_mDYN_DESC_LEN_ERROR)) & CStr(mobjParentObject.Max_Dynamo_Desc_Length)
            MsgBox sError, , sTitle
            BDW_Update_Dynamo_Desc = False
        End If
    Else
        BDW_Update_Dynamo_Desc = True
    End If
    
    Exit Function
ErrorHandler:
    HandleError
End Function


Sub BDW_UpdateSymbolDataWithCreateDynamoArray()
'******************************************************************************************
'PURPOSE: To update the member .strPrompt in the BDW_gudtSymbolData with the new User Prompt
'This is stored in the temporary array called BDW_CreateDynamoArray.
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_gudtSymbolData Data Type.
    
    On Error GoTo ErrorHandler
    For intIndex = 0 To BDW_mintCDANumOfIndeces
        With BDW_gudtSymbolData(intIndex)
            .strPrompt = BDW_CreateDynamoArray(intIndex, 2)
        End With
    Next intIndex
    Exit Sub

ErrorHandler:
    HandleError
End Sub
Sub BDW_UpdatePromptDataWithEditDynamoArray()
'******************************************************************************************
'PURPOSE: To update the member .strPrompt in the BDW_mudtPromptData with the new Current Setting
'This is stored in the temporary array called BDW_EditDynamoArray.
'******************************************************************************************
    Dim intIndex As Integer 'Counter for the For Loop and Index of the BDW_mudtPromptData Data Type.
    
    On Error GoTo ErrorHandler
    For intIndex = 0 To BDW_mintEDANumOfIndeces
        With BDW_mudtPromptData(intIndex)
            'if there is not a source, and there is a User prompt then set the Error Flag to True
            If BDW_EditDynamoArray(intIndex, 1) = "" And BDW_EditDynamoArray(intIndex, 0) <> "" Then
                gblnNoSource = True
                gintNumberInArray = intIndex
            Else
                .strContent = BDW_EditDynamoArray(intIndex, 1)
            End If
        End With
    Next intIndex
    Exit Sub

ErrorHandler:
    HandleError
End Sub
'  JPB031403  Tracker #865  Added this function to validate dynamo names
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IsAlphanumericStartIsAlpha()
'
'   determines whether filename is alphanumeric and starts with alpha
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsAlphanumericStartIsAlpha(strName As String) As Boolean
    Dim szName As String
    szName = strName    'make a copy so we don't modify it
       
    If Len(szName) <= 0 Then
        IsAlphanumericStartIsAlpha = False
        Exit Function
    End If

    'check for start with alpha, and the rest is alphanumeric
    'allow underscores
    If szName Like objStrMgr.GetNLSStr(CLng(NLS_mSTARTCHARS)) And _
                    Not Mid$(szName, 2) Like objStrMgr.GetNLSStr(CLng(NLS_mRESTOFCHARS)) Then
        IsAlphanumericStartIsAlpha = True
    Else
        IsAlphanumericStartIsAlpha = False
    End If
    
End Function

'hj020504 Clarify #283379
'The Sub allows multiple occurences of partial substitution in an expression
'
Private Sub BDW_SubstituteNodeTag(strExpression As String, strNodeTag As String)
    Dim strTemp As String
    Dim nSpacePosition, nDotPosition As Integer
    Dim nStart As Integer
    
    If Len(strExpression) <= 0 Or Len(strNodeTag) <= 0 Then
        Exit Sub
    End If
    
    strTemp = ""
    nStart = Len(strExpression)
    
    Do While nStart > 0
        strTemp = strNodeTag & strTemp
        'Check if there is a white space in the expression.
        'If yes, this may be a complex expression.
        nSpacePosition = InStrRev(strExpression, " ", nStart)
        nStart = nSpacePosition
        If nSpacePosition > 0 Then
            'This is a complex expression. Check if there is another NTF need to be substituted.
            nDotPosition = InStrRev(strExpression, ".", nSpacePosition)
            nStart = nDotPosition
            If nDotPosition > 0 Then
                'Handle the field part of NTF and operators
                strTemp = Mid(strExpression, nDotPosition, nSpacePosition - nDotPosition + 1) & strTemp
            Else
                'Only operators
                strTemp = Mid(strExpression, 1, nSpacePosition) & strTemp
            End If
        End If
    Loop
    
    strExpression = strTemp
    
End Sub
Public Function IsIHistorian() As Boolean
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
        IsIHistorian = True
    Else
        ' treat anything else as classic
        IsIHistorian = False
    End If
End Function
