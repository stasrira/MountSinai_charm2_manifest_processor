Attribute VB_Name = "mdlMain"
Option Explicit

Public Const Version = "1.01"

Private Const ShipmentWrkSheet = "shipment"
Private Const DemographicWrkSheet = "study_demographic"
Private Const ConfigWrkSheet = "config"
Private Const ManifestWrkSheet = "manifest"
Private Const MetadataWrkSheet = "metadata"
Private Const DictionaryWrkSheet = "dictionary"
Private Const LogsWrkSheet = "logs"
Private Const LabStatsWrkSheet = "lab_stats"

Public Enum LogMsgType
    Debugging = 0
    info = 1
    warning = 2
    Error = 3
End Enum

Public Enum LogMsgTypeColor
    Debugging = 22428 'RGB(156, 87, 0) DarkYellow
    info = 0 'RGB(0, 0, 0) Black
    warning = 16711680 'RGB(0, 0, 255) Blue
    Error = 255 'RGB(255, 0, 0) Red
End Enum

Public Enum ValidationResults
    OK = 0
    warning = 1
    Error = 2
End Enum

Dim firstFileLoaded As Boolean

'returns dictionary providing mapping between specimen prep values from "shipment" tab and the specimen prep abbreviations expected by the system
Private Function GetSpecimenAbbreviationsDict() As Scripting.dictionary
    Set GetSpecimenAbbreviationsDict = GetDictItems("SpecimenAbbreviation")
End Function

'returns dictionary providing mapping between specimen prep values from "shipment" tab and the specimen prep name expected by the system
Private Function GetSpecimenFullNameDict() As Scripting.dictionary
    Set GetSpecimenFullNameDict = GetDictItems("SpecimenFullName")
End Function

'returns dictionary providing mapping between specimen prep values from "shipment" tab and the aliquot volume expected by the system for such specimen
Private Function GetVolumeBySpecPrepDict() As Scripting.dictionary
    Set GetVolumeBySpecPrepDict = GetDictItems("VolumeBySpecPrep")
End Function

'returns dictionary providing mapping between specimen prep values from "shipment" tab and the comments expected by the system for such specimen
Private Function GetCommentBySpecPrepDict() As Scripting.dictionary
    Set GetCommentBySpecPrepDict = GetDictItems("CommentsBySpecPrep")
End Function

'returns dictionary providing mapping between specimen prep values from "shipment" tab and the tissue name expected by the system for such specimen
Private Function GetTissueBySpecPrepDict() As Scripting.dictionary
    Set GetTissueBySpecPrepDict = GetDictItems("TissueBySpecPrep")
End Function

Private Function GetDictItems(dict_name As String) As Scripting.dictionary
    Dim col_num As Integer
    Dim dict As New Scripting.dictionary

    col_num = FindColNumberOfDictCategory(dict_name)
    If col_num > 0 Then
        Set dict = GetDictItemsPerColNum(col_num)
    Else
        Set dict = Nothing
    End If
    
    Set GetDictItems = dict
End Function

'gets value of the cell (from "shipment" tab) based on the given config parameter name (that identifies the column) and given row number
Private Function GetCellValuePerRow(row_num As Integer, cfg_param_name As String) As String
    Dim cfg_row As Integer
    Dim column_letter As String
    Dim out_val As String
    
    'get config value to identify column letter on the shipment tab
    column_letter = GetConfigParameterValue(cfg_param_name)
    
    'get value
    out_val = Worksheets(ShipmentWrkSheet).Range(column_letter & CStr(row_num))
    GetCellValuePerRow = out_val
End Function

Private Function GetConfigParameterValueByColumn(cfg_param_name As String, column_letter As String, Optional wb As Workbook = Nothing) As String
    Dim cfg_row As Integer
    Dim out_val As String
    Dim ws_cfg As Worksheet
    
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(ConfigWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(ConfigWrkSheet)
    End If
    
    'get config value to identify column letter on the shipment tab
    cfg_row = FindRowNumberOfConfigParam(cfg_param_name, wb)
    If cfg_row > 0 Then
        'get configuration value
        out_val = ws_cfg.Range(column_letter & CStr(cfg_row))
    Else
        out_val = ""
    End If
    GetConfigParameterValueByColumn = out_val
End Function

Private Function SetConfigParameterValueByColumn(cfg_param_name As String, column_letter As String, value_to_set As String, Optional wb As Workbook = Nothing)
    Dim cfg_row As Integer
    Dim out_val As String
    Dim ws_cfg As Worksheet
    
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(ConfigWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(ConfigWrkSheet)
    End If
    
    'get config value to identify column letter on the shipment tab
    cfg_row = FindRowNumberOfConfigParam(cfg_param_name, wb)
    If cfg_row > 0 Then
        'set configuration value
        ws_cfg.Range(column_letter & CStr(cfg_row)).Value = value_to_set
    End If
End Function

Private Function GetConfigParameterValue(cfg_param_name As String, Optional wb As Workbook = Nothing) As String
    'retrieve value from column B of the config tab
    GetConfigParameterValue = GetConfigParameterValueByColumn(cfg_param_name, "B", wb)
End Function

Private Function GetConfigParameterValue_SheetAssignment(cfg_param_name As String, Optional wb As Workbook = Nothing) As String
    'retrieve value from column B of the config tab
    GetConfigParameterValue_SheetAssignment = GetConfigParameterValueByColumn(cfg_param_name, "C", wb)
End Function

Private Function SetConfigParameterValue(cfg_param_name As String, cfg_param_new_value As String, Optional wb As Workbook = Nothing) As Boolean
    Dim cfg_row As Integer
    Dim out_val As Boolean
    Dim ws_cfg As Worksheet
    
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(ConfigWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(ConfigWrkSheet)
    End If
    
    'get config value to identify column letter on the shipment tab
    cfg_row = FindRowNumberOfConfigParam(cfg_param_name, wb)
    If cfg_row > 0 Then
        'get configuration value
        ws_cfg.Range("B" & CStr(cfg_row)) = cfg_param_new_value
        out_val = True
    Else
        out_val = False
        AddLogEntry "Failed to set config parameter '" & cfg_param_name & "' because it was not found in the config file.", LogMsgType.Error
    End If
    SetConfigParameterValue = out_val
End Function

'parses caption of the Specimen Prep column to retrieve specimen prep name and timepoint from it
Public Function GetItemValueFromSpecimenHeader(item_name As String, Optional specimen_header_param As String = "") As String
    Dim out_val As String
    Dim specimen_header As String, timepoint As String
    Dim specimen_prep As String, specimen_prep_abbr As String, specimen_prep_name As String
    Dim arr() As String

    'get specimen column header
    If specimen_header_param = "" Then
        specimen_header = GetCellValuePerRow(1, "Specimen Column")
    Else
        specimen_header = specimen_header_param
    End If
    
    'get timepoint and specimen prep out of the header of the specimen column
    arr = Split(specimen_header, " ")
    
    Select Case item_name
        Case "specimen_prep_abbr"
            'get specimen prep
            If ArrLength(arr) > 0 Then
                specimen_prep = arr(0)
            Else
                specimen_prep = ""
            End If
            'get specimen prep values
            specimen_prep_abbr = GetSpecimenPrep(specimen_prep, True) 'get abbreviation for speciemn prep
            out_val = specimen_prep_abbr
        Case "specimen_prep_name"
            'get specimen prep
            If ArrLength(arr) > 0 Then
                specimen_prep = arr(0)
            Else
                specimen_prep = ""
            End If
            'get specimen prep values
            specimen_prep_name = GetSpecimenPrep(specimen_prep, False) 'get full name of the specimen prep
            out_val = specimen_prep_name
        Case "timepoint"
            'get timepoint
            out_val = GetConfigParameterValue("Subject Visit number")
            'CHARM1 approach
'            timepoint = arr(0)
'            'adjust timepoint, if needed
'            If Len(timepoint) = 2 Then
'                timepoint = Left(timepoint, 1) & "0" & Right(timepoint, 1)
'            End If
'            out_val = timepoint
    End Select
    
    GetItemValueFromSpecimenHeader = out_val
End Function

'prepares aliquot_id value based on the provided row number (on the "shipment" tab)
Public Function GetAliquotId(row_num As Integer) As String
    Dim specimen_prep_abbr As String
    Dim sample_id As String, aliquot_id As String
    Dim subject_id As String, date_part As String
    
    'get samlpe_id for the current row
    sample_id = GetSampleId(row_num)
    
    If Len(Trim(sample_id)) > 0 Then
        'get specimen prep abbreviation value out of the Specimen prep column's header
        specimen_prep_abbr = GetItemValueFromSpecimenHeader("specimen_prep_abbr")
        'get visit date value through the subject id assigned to the current row
        subject_id = GetCellValuePerRow(row_num, "Participant Id_shipment")
        date_part = convertDate(GetColumnValueBySubjectID("Visit Date", subject_id), "yymmdd")
        
        'concatenate sample id and specimen prep to get aliquot id
        aliquot_id = sample_id & "-" & date_part & "-" & Trim(specimen_prep_abbr)
        
'        If date_part = "" Then
'            'Log error that aliquot id was not properly created
'            AddLogEntry "Cannot obtain 'Visit Date' value for the subject '" & subject_id & "' while preparing an aliquot id for it - " & aliquot_id & ".", LogMsgType.Error
'        End If
    Else
        aliquot_id = ""
    End If
    GetAliquotId = aliquot_id
End Function

'prepares sample_id value based on the provided row number (on the "shipment" tab)
Public Function GetSampleId(row_num As Integer) As String
    Dim subjects_column As String, specimen_column As String
    Dim subject_id As String, timepoint As String, sample_prefix As String
    Dim specimen_prep_abbr As String
    Dim sample_id As String, timepoint_mark As String
    
    'get subject id
    subject_id = GetCellValuePerRow(row_num, "Participant Id_shipment")
    
    'get time point confirmation - check value of the specimen_column for the current row
    timepoint_mark = GetCellValuePerRow(row_num, "Specimen Column")
    
    'get timepoint value out of the Specimen prep column's header
    timepoint = GetItemValueFromSpecimenHeader("timepoint")
    
    'get a constant prefix part to be added to the sample id value
    sample_prefix = GetConfigParameterValue("Sample_ID_Prefix")
    
    'check that row is not empty (subject is present) and the timepoint field has value "Y"
    If Len(Trim(subject_id)) > 0 And UCase(timepoint_mark) = "Y" Then
        'combine aliquot id
        sample_id = sample_prefix & Trim(subject_id) & "-" & Trim(timepoint) '& "-" & specimen_prep_abbr
    Else
        sample_id = ""
    End If
    
    GetSampleId = sample_id
End Function

Public Function GetTimepoint() As String
    GetTimepoint = GetItemValueFromSpecimenHeader("timepoint")
End Function

Public Function GetSpecimenPrepAbbr() As String
    GetSpecimenPrepAbbr = GetItemValueFromSpecimenHeader("specimen_prep_abbr")
End Function

Public Function GetSpecimenPrepName() As String
    GetSpecimenPrepName = GetItemValueFromSpecimenHeader("specimen_prep_name")
End Function

'converts provided specimen_prep name to an expected name/abbreviation of it based on the harcoded mapping dictionary
Private Function GetSpecimenPrep(specimen_prep As String, Optional abbreviation As Boolean = False) As String
    Dim dictionary As Scripting.dictionary
    Dim out_val As String
    
    If abbreviation Then
        Set dictionary = GetSpecimenAbbreviationsDict()
    Else
        Set dictionary = GetSpecimenFullNameDict()
    End If
        
    If dictionary.exists(LCase(specimen_prep)) Then
        out_val = dictionary(LCase(specimen_prep))
    Else
        out_val = "N/D"
    End If
    GetSpecimenPrep = out_val
End Function

'searches for a given subject id value on the "shipment" tab (in the subject_id column defined in the config tab) and retrievs value of the requested column for the row
Public Function GetColumnValueBySubjectID(colName As String, subject_id As String) As String
    Dim cfg_row As Integer, row_num As Integer
    Dim parameter_column As String, parameter_val As String
    Dim parameter_tab_name As String
    Dim subjects_column As String, subject_config_parameter_name As String
    Dim ws As Worksheet
    
    'get column letter for the required parameter
    
    'get config value
    parameter_column = GetConfigParameterValue(colName)
    parameter_tab_name = GetConfigParameterValue_SheetAssignment(colName)
    
    If Len(Trim(parameter_column)) = 0 Then
        'if configuration returns no value, return blank
        GetColumnValueBySubjectID = ""
        Exit Function
    End If
    
    'set ws reference based on the parameter's sheet assignment
    If Len(Trim(parameter_tab_name)) > 0 Then
        Set ws = Worksheets(parameter_tab_name)
    Else
        Set ws = Worksheets(ShipmentWrkSheet)
    End If
    
    'get name of the "Participant Id" parameter for the given parameter's sheet assignment
    subject_config_parameter_name = "Participant Id" & "_" & parameter_tab_name
    'get column letter for the subject column
    subjects_column = GetConfigParameterValue(subject_config_parameter_name)
    
    If IsError(Application.Match(subject_id, ws.Range(subjects_column & ":" & subjects_column), 0)) Then
        GetColumnValueBySubjectID = ""
        Exit Function
    Else
        'row number where subject information is located on the Shipment sheet
        row_num = Application.Match(subject_id, ws.Range(subjects_column & ":" & subjects_column), 0)
        parameter_val = ws.Range(parameter_column & row_num)
        GetColumnValueBySubjectID = parameter_val
    End If
    
End Function

'searches for a given parameter name on the config page and returns the row number it was found on
Private Function FindRowNumberOfConfigParam(param_name As String, Optional wb As Workbook = Nothing) As Integer
    Dim ws_cfg As Worksheet
    
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(ConfigWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(ConfigWrkSheet)
    End If
    
    If IsError(Application.Match(param_name, ws_cfg.Range("A:A"), 0)) Then
        FindRowNumberOfConfigParam = 0
    Else
        FindRowNumberOfConfigParam = Application.Match(param_name, ws_cfg.Range("A:A"), 0)
    End If
    
End Function

'searches for a given parameter name on the config page and returns the row number it was found on
Private Function FindColNumberOfDictCategory(categ_name As String, Optional wb As Workbook = Nothing) As Integer
    Dim ws_cfg As Worksheet
    
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(DictionaryWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(DictionaryWrkSheet)
    End If
    
    If IsError(Application.Match(categ_name, ws_cfg.Rows(1), 0)) Then
        FindColNumberOfDictCategory = 0
    Else
        FindColNumberOfDictCategory = Application.Match(categ_name, ws_cfg.Rows(1), 0)
    End If
    
End Function

Private Function GetDictItemsPerColNum(col_num As Integer, Optional wb As Workbook = Nothing) As dictionary
    Dim ws_cfg As Worksheet
    Dim dict_range As Range
    Dim dict As New Scripting.dictionary
    Dim cell As Range
        
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(DictionaryWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(DictionaryWrkSheet)
    End If
    
    'get range based on the number of column provided as a parameter
    Set dict_range = ws_cfg.Columns(col_num)
    
    'loop through cells of the range
    For Each cell In dict_range.Cells
        If cell.Row > 1 Then 'proceed only if this is not a first (header) row
            If cell.Row > ws_cfg.UsedRange.Rows.Count Then
                Exit For
            End If
            'Debug.Print cell.Address & " - " & cell.Offset(0, 1).Address
            If Len(Trim(cell.Value2)) > 0 Then 'check the dictionary key value is not blank
                'Debug.Print cell.Value2 & " - " & cell.Offset(0, 1).Value2
                dict.Add cell.Value2, cell.Offset(0, 1).Value2
            End If
        End If
    Next
   
    Set GetDictItemsPerColNum = dict
End Function

'retrievs predefiend volume values based on the specimen prep name or abbreviation
Function GetSampleVolumeBySpecimenPrep(specimen_prep As String) As String
    Dim dictionary As Scripting.dictionary
    Dim out_val As String
    
    Set dictionary = GetVolumeBySpecPrepDict()
        
    If dictionary.exists(LCase(specimen_prep)) Then
        out_val = dictionary(LCase(specimen_prep))
    Else
        out_val = "N/D"
    End If
    GetSampleVolumeBySpecimenPrep = out_val
End Function

'retrievs predefiend comments based on the specimen prep name or abbreviation
Function GetCommentsBySpecimenPrep(specimen_prep As String) As String
    Dim dictionary As Scripting.dictionary
    Dim out_val As String
    
    Set dictionary = GetCommentBySpecPrepDict()
        
    If dictionary.exists(LCase(specimen_prep)) Then
        out_val = dictionary(LCase(specimen_prep))
    Else
        out_val = ""
    End If
    GetCommentsBySpecimenPrep = out_val
End Function

'retrievs predefiend tissue name based on the specimen prep name or abbreviation
Function GetSampleTissueBySpecimenPrep(specimen_prep As String) As String
    Dim dictionary As Scripting.dictionary
    Dim out_val As String
    
    Set dictionary = GetTissueBySpecPrepDict()
        
    If dictionary.exists(LCase(specimen_prep)) Then
        out_val = dictionary(LCase(specimen_prep))
    Else
        out_val = "N/D"
    End If
    GetSampleTissueBySpecimenPrep = out_val
End Function

Private Function GetFileNameToSave(file_type As String)
    'expected file_type values: "metadata", "manifest"
    
    Dim timepoint As String, spec_prep As String, ship_date As String, cfg_row As Integer
    Dim post_fix As String
    
    timepoint = GetTimepoint() 'get timepoint value
    spec_prep = GetSpecimenPrepAbbr() 'get specimen abbreviation value
    ship_date = GetConfigParameterValue("Shipment Date") 'get shipment date value from config tab
    post_fix = GetConfigParameterValue("File Name Post-fix") 'get post-fix value for created files
    
    If Not IsDate(ship_date) Then
        'verify the returned ship_date value
        ship_date = Date
    End If
    
    ship_date = format(ship_date, "yyyy_mm_dd")
    
    GetFileNameToSave = "CHARM2_" & file_type & "_" & ship_date & "_" & timepoint & "_" & spec_prep & "_prepared_internally" & Trim(post_fix) & ".xlsx"
End Function

Private Function CreateNewFile(file_type As String, file_name As String) As Workbook
    'expected file_type values: "metadata", "manifest"
    Dim path As String, cfg_row As Integer, new_file_name As String
    'Dim abort As Boolean: abort = False
    Dim wb As Workbook
    
    Set wb = Workbooks.Add(xlWBATWorksheet) 'create a new excel file with a single sheet
    wb.Sheets(1).Name = file_type 'rename the sheet of the workbook
    On Error GoTo Err_SaveAs
    wb.SaveAs (file_name) 'save new excel file to a specified folder with the given name
    On Error GoTo 0
    Set CreateNewFile = wb

    Exit Function
Err_SaveAs:
    wb.Close False
    Set CreateNewFile = Nothing
End Function

Function SavePreparedFiles(Optional show_confirmation As Boolean = True) As dictionary()
    Dim out_arr(1) As New dictionary
    Dim i As Integer
    Dim msg_str As String: msg_str = ""
    Dim msg_summary_str As String
    Dim total_status As Integer: total_status = vbInformation
    Dim validation_result As ValidationResults
    Dim specimen_header As String, specimen_letter As String
    Dim iResponse As Integer
    Dim cntOk As Integer: cntOk = 0
    Dim cntExist As Integer: cntExist = 0
    Dim cntError As Integer: cntError = 0
    Dim create_manifest As String, create_metadata As String
    
    On Error GoTo err_lab
    
    If show_confirmation Then
        'if the function runs not as a part of bigger process, confirm if user want to proceed.
        iResponse = MsgBox("The system is about to start processing manfiest data on the 'shipment' tab." & vbCrLf _
                    & "Be aware that some output files will be created as an otcome of the process!" _
                    & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'." & vbCrLf & vbCrLf _
                    & "Note: some screen flickering might occur during creating of output excel files.", _
                    vbOKCancel + vbInformation, "CHARM2 Manifest Processor")
    Else
        iResponse = vbOK
    End If
    
    If iResponse <> vbOK Then
        'exit function based on user's response
        Exit Function
    End If
    
    If show_confirmation Then
        'if show_confirmation is True, this is a single call to this function and previous logs should be cleared
        CleanLogsWorksheet
    End If
    
    'get specimen column header
    specimen_header = GetCellValuePerRow(1, "Specimen Column")
    'get config value corresponding to the column letter on the shipment tab
    specimen_letter = GetConfigParameterValue("Specimen Column")
    
    AddLogEntry "Start processing data for the '" & specimen_header & "' Specimen (column " & specimen_letter & ").", LogMsgType.info
    
    'recalculate whole workbook to make sure manifest and metadata sheets are filled properly
    RefreshWorkbookData

    AddLogEntry "Start validating tasks for the Specimen column '" & specimen_letter & "'.", LogMsgType.info
    
    'run actual validation process
    validation_result = ValidateCurrentSpecimenTimepointColumn(False)
                
    If validation_result = ValidationResults.OK Then
        AddLogEntry "Finish validating tasks for the Specimen column '" & specimen_letter & "'. No issues were identified.", LogMsgType.info
    Else
        AddLogEntry "Finish validating tasks for the Specimen column '" & specimen_letter & "'. Some issues or warnings were identified - see previous log entries.", LogMsgType.warning
    End If
    
    'retrieve config parameters to identify which files to be created
    create_manifest = GetConfigParameterValue("Create Manifest files")
    create_metadata = GetConfigParameterValue("Create Metadata files")
    
    Application.ScreenUpdating = False
    
    If validation_result = ValidationResults.OK Or validation_result = ValidationResults.warning Then
        AddLogEntry "Proceeding to create manifest and metadata files for the '" & specimen_header & "' Specimen combination.", LogMsgType.info
        
        'run the process of creating manifest files
        If UCase(create_manifest) = "TRUE" Then
            Set out_arr(0) = SavePreparedData("manifest")
        Else
            AddLogEntry "Creating Manifest file for the Specimen column '" & specimen_letter _
                & "' was not performed based on the 'Create Manifest files' configuration setting.", LogMsgType.warning
            out_arr(0).Add "status", "EXCLUDED"
            out_arr(0).Add "msg", "Manifest file creation was skipped based on the configuration settings."
        End If
        'add validation status to the results of file creation
        out_arr(0).Add "validation", validation_result
        
        'run the process of creating metadata files
        If UCase(create_metadata) = "TRUE" Then
            Set out_arr(1) = SavePreparedData("metadata")
            'add validation status to the results of file creation
            out_arr(1).Add "validation", validation_result
        Else
            AddLogEntry "Creating Metadata file for the Specimen column '" & specimen_letter _
                & "' was not performed based on the 'Create Metadata files' configuration setting.", LogMsgType.warning
            out_arr(1).Add "status", "EXCLUDED"
            out_arr(1).Add "msg", "Metadata file creation was skipped based on the configuration settings."
        End If

        AddLogEntry "Finish creating manifest and metadata files for the '" & specimen_header & "' Specimen combination.", LogMsgType.info
    Else
        out_arr(0).Add "status", "ERROR"
        out_arr(0).Add "msg", "File creation was skipped due to validation errors"
        out_arr(0).Add "validation", validation_result
        out_arr(1).Add "status", "ERROR"
        out_arr(1).Add "msg", "File creation was skipped due to validation errors"
        out_arr(1).Add "validation", validation_result
        AddLogEntry "Since validation failed for the '" & specimen_header & "' Specimen, no manifest or metadata files were created.", LogMsgType.Error
    End If
    
    Application.ScreenUpdating = True
    
    If show_confirmation Then
        Worksheets(LogsWrkSheet).Activate 'bring focus to the "logs" tab
        Worksheets(LogsWrkSheet).Cells(1, 1).Activate 'bring focus to the first cell on the sheet
        
        DisplayConfirmationMsg out_arr

    End If
    
    AddLogEntry "Finish processing data for the '" & specimen_header & "' Specimen.", LogMsgType.info
    
    SavePreparedFiles = out_arr
    
    Exit Function
    
err_lab:
    Application.ScreenUpdating = True
    AddLogEntry "The following error occurred: " & Err.Description, LogMsgType.Error
    If show_confirmation Then
        MsgBox "Unexpected error has occurred: " & Err.Description, vbCritical, "CHARM processing tool - ERROR"
    End If
End Function

Private Sub DisplayConfirmationMsg(out_arr() As dictionary) ', Optional valid_status As ValidationResults = ValidationResults.OK)
    Dim i As Integer
    Dim total_status As Integer: total_status = vbInformation
    Dim cntOk As Integer: cntOk = 0
    Dim cntExist As Integer: cntExist = 0
    Dim cntError As Integer: cntError = 0
    Dim cntEmpty As Integer: cntEmpty = 0
    Dim cntExcluded As Integer: cntExcluded = 0
    Dim msg_str As String, msg_summary_str As String
    Dim msg1_str As String: msg1_str = ""
    Dim valid_status As ValidationResults: valid_status = ValidationResults.OK
    
    
    If ArrLength(out_arr) > 1 Then 'if array is not empty
        For i = 0 To ArrLength(out_arr) - 1
            'loop through reported statuses and count each status kind
            If Not out_arr(i) Is Nothing Then
                'results of the process are present, collect details
                Select Case out_arr(i)("status")
                    Case "OK"
                        cntOk = cntOk + 1
                    Case "EXISTS"
                        cntExist = cntExist + 1
                        If total_status <> vbCritical Then total_status = vbExclamation
                    Case "EMPTY"
                        cntEmpty = cntEmpty + 1
                        If total_status <> vbCritical Then total_status = vbExclamation
                    Case "ERROR"
                        cntError = cntError + 1
                        total_status = vbCritical
                    Case "EXCLUDED"
                        cntExcluded = cntExcluded + 1
                        If total_status <> vbCritical Then total_status = vbExclamation
                End Select
                If IsNumeric(out_arr(i)("validation")) Then
                    'collect the worst validation status from the array of the provided results
                    If valid_status < out_arr(i)("validation") Then
                        valid_status = out_arr(i)("validation")
                    End If
                End If
            Else
                'if results of the process are not present, set status to Error
                'valid_status = ValidationResults.Error
                total_status = vbCritical
            End If
        Next
        
        If total_status = vbInformation And cntOk = 0 Then
            'if no file were prepared successfully, but total_status is still set OK (vbInformation), change it to vbExclamation
            total_status = vbExclamation
        End If
        
        If valid_status <> ValidationResults.OK And total_status = vbInformation Then
            'if validation status is not OK, but total_status is still set OK (vbInformation), change it to vbExclamation
            total_status = vbExclamation
        End If
        
        'prepare warning messages
        If cntExist > 0 Then
            msg1_str = msg1_str & vbCrLf & "- " & CStr(cntExist) & " of file(s) being created was(were) present and skipped based on the user's input, check the Warnings on the 'logs' tab for details."
        End If
        If cntEmpty > 0 Then
            msg1_str = msg1_str & vbCrLf & "- " & CStr(cntEmpty) & " of created file(s) was(were) empty and not saved, check the Warnings on the 'logs' tab for details."
        End If
        If cntError > 0 Then
            msg1_str = msg1_str & vbCrLf & "- " & "Some ERRORs were reported, check the warnings on the 'logs' tab for details!"
        End If
        If cntExcluded > 0 Then
            msg1_str = msg1_str & vbCrLf & "- " & CStr(cntExcluded) & " file(s) was(were) Excluded from the creation process based on the configuration settings, check the Warnings on the 'logs' tab for details."
        End If
        If valid_status = ValidationResults.Error Then
            msg1_str = msg1_str & vbCrLf & "- " & "Some Validation ERRORs were reported, check the Errors on the 'logs' tab for details!"
            total_status = vbCritical
        End If
        If valid_status = ValidationResults.warning Then
            msg1_str = msg1_str & vbCrLf & "- " & "Some Validation Warnings were reported, check the Warnings on the 'logs' tab for details!"
            If total_status = vbInformation Then
                total_status = vbExclamation
            End If
        End If
        
        If Len(Trim(msg1_str)) = 0 Then
            msg1_str = "Note:" & vbCrLf & "You can check details of the process on the 'logs' tab."
        Else
            msg1_str = "WARNINGS:" & msg1_str
        End If
        
        msg_summary_str = "Process has finished. Summary of prepared files: " & vbCrLf _
                & "Newly created files: " & CStr(cntOk) & vbCrLf _
                & "Preserved existing files: " & CStr(cntExist) & vbCrLf _
                & "Not created due to errors: " & CStr(cntError) & vbCrLf _
                & vbCrLf _
                & msg1_str

    Else
        msg_summary_str = "Process has finished, but no confirmation messages were retrieved. Check 'logs' tab for the available details!"
        total_status = vbCritical
    End If
    
    MsgBox msg_summary_str, total_status, "Create Manifest & Metadata files"
    
End Sub

Private Function SavePreparedData(file_type As String) As dictionary
    Dim abort As Boolean: abort = False
    Dim wb As Workbook, ws_source As Worksheet, ws_target As Worksheet
    Dim wb_source As Workbook
    Dim out_dict As New Scripting.dictionary
    Dim new_file_name As String, path As String
    Dim str1 As String
    Dim empty_file_flag As Boolean
        
    Set wb_source = Application.ActiveWorkbook
    
    Select Case file_type
        Case "metadata"
            Set ws_source = Worksheets(MetadataWrkSheet) 'set reference to the "metadata" sheet
            path = GetConfigParameterValue("Save Created Metadata Files Path")
        Case "manifest"
            Set ws_source = Worksheets(ManifestWrkSheet) 'set reference to the "manifest" sheet
            path = GetConfigParameterValue("Save Created Manifests Path")
        Case Else
            abort = True
    End Select
    
    new_file_name = path & "\" & GetFileNameToSave(file_type)  'get the file name for the new excel file
    
    If Not abort Then
        AddLogEntry "Start creating new '" & file_type & "' file in '" & path & "' folder.", LogMsgType.info
        
        Set wb = CreateNewFile(file_type, new_file_name) ' create a new excel file and get a reference to it
        If Not wb Is Nothing Then
            Set ws_target = wb.Sheets(1) 'get reference to the worksheet of the new file
            
            ws_source.Cells.Copy 'copy data from a source sheet to a memory
            ws_target.Cells.PasteSpecial Paste:=xlPasteValues 'paste data from memory to the target sheet as "values only"
            
            CleanCreatedFile ws_target, file_type, wb_source
            
            wb.Save 'save the new file
            
            If ws_target.UsedRange.Rows.Count > 1 Then
                'if the cleaned worksheet contains at least on data row (beside the header), proceed here
                empty_file_flag = False 'set flag to save changes on closing
                
                str1 = "New " & file_type & " file was created - " & wb.FullName '& ""
                out_dict.Add "msg", str1
                out_dict.Add "status", "OK" 'vbInformation
                
                AddLogEntry str1, LogMsgType.info, wb_source
            Else
                empty_file_flag = True 'set flag not to save changes on closing
            
                str1 = "Newly created " & file_type & " file - " & wb.FullName & " - appears to be empty and will be deleted."
                out_dict.Add "msg", str1
                out_dict.Add "status", "EMPTY" 'vbInformation
                
                AddLogEntry str1, LogMsgType.warning, wb_source
            End If
            
            wb.Close 'close the new file
            
            If empty_file_flag Then
                'delete just created empty file
                If DeleteFile(new_file_name) Then
                    str1 = "Newly created " & file_type & " file - " & new_file_name & " - was successfully deleted."
                    AddLogEntry str1, LogMsgType.warning, wb_source
                Else
                    str1 = "The application was not able to delete the newly created " & file_type & " file - " & wb.FullName & "."
                    AddLogEntry str1, LogMsgType.Error, wb_source
                End If
            End If
            
            Application.CutCopyMode = False 'clean clipboard
        Else
            str1 = "Existing file with the same name was present (" & new_file_name & "). Creation of the new " & file_type & " file was skipped based on user's input."
            out_dict.Add "msg", str1
            out_dict.Add "status", "EXISTS" 'vbInformation
            AddLogEntry str1, LogMsgType.warning
        End If
        AddLogEntry "Finish creating new '" & file_type & "' file.", LogMsgType.info
    Else
        str1 = "Failed to create a new " & file_type & " file!"
        out_dict.Add "msg", str1
        out_dict.Add "status", "ERROR" 'vbCritical
        AddLogEntry str1, LogMsgType.Error
    End If
    
    Set SavePreparedData = out_dict
End Function

Private Sub CleanCreatedFile(ws_target As Worksheet, file_type As String, wb_source As Workbook)
    Dim abort As Boolean: abort = False
    Dim deleteColumns As String: deleteColumns = ""
    Dim formatDateColumns As String: formatDateColumns = ""
    Dim col As Variant, cfg_row As Integer
    
    Select Case file_type
        Case "metadata"
            'get list of columns that have to be deleted in the created file
            deleteColumns = GetConfigParameterValue("Metadata Delete Columns In Target", wb_source)
        Case "manifest"
            'get list of columns that have to be deleted in the created file
            deleteColumns = GetConfigParameterValue("Manifest Delete Columns In Target", wb_source)
            'get list of columns that have to be formatted as date in the created file
            formatDateColumns = GetConfigParameterValue("Manifest Format Date Columns In Target", wb_source)
        Case Else
            abort = True
    End Select
    
    If Not abort Then
        If Len(Trim(deleteColumns)) > 0 Then
            'loop through columns and delete one by one
            For Each col In Split(deleteColumns, ",")
                col = Trim(col)
                If Len(col) > 0 Then
                    ws_target.Columns(col).Delete
                End If
            Next
        End If
        
        If Len(Trim(formatDateColumns)) > 0 Then
            'loop through columns and delete one by one
            For Each col In Split(formatDateColumns, ",")
                col = Trim(col)
                If Len(col) > 0 Then
                    ws_target.Columns(col).NumberFormat = "mm/dd/yyyy"
                End If
            Next
        End If
        
        DeleteBlankRows ws_target 'delete blank rows in the new file
    End If
End Sub

Private Sub DeleteBlankRows(ws_target As Worksheet)
    Dim SourceRange As Range
    Dim entireRow As Range
    Dim i As Long, non_blanks As Long, empty_strings As Long
 
    Set SourceRange = ws_target.UsedRange ' Cells.End(xlToLeft)
 
    If Not (SourceRange Is Nothing) Then
        'Application.ScreenUpdating = False
 
        For i = SourceRange.Rows.Count To 1 Step -1
            Set entireRow = SourceRange.Cells(i, 1).entireRow
            non_blanks = Application.WorksheetFunction.CountA(entireRow)
            empty_strings = Application.WorksheetFunction.CountIf(entireRow, "")
            If non_blanks = 0 Or entireRow.Cells.Count = empty_strings Then
                entireRow.Delete
            Else
                'Print ("Not blank row")
            End If
        Next
 
        'Application.ScreenUpdating = True
    End If
End Sub

'returns True if the file was deleted
Private Function DeleteFile(ByVal FileToDelete As String) As Boolean
    Dim fso As New FileSystemObject, aFile As File
    
    If (fso.FileExists(FileToDelete)) Then
        Set aFile = fso.GetFile(FileToDelete)
        aFile.Delete
    End If
    
    DeleteFile = Not (fso.FileExists(FileToDelete))
End Function

Private Function ArrLength(a As Variant) As Integer
   If IsEmpty(a) Then
      ArrLength = 0
   Else
      ArrLength = UBound(a) - LBound(a) + 1
   End If
End Function

Public Sub RefreshWorkbookData()
    Dim refresh_db As String
    
    'recalculate whole workbook to make sure manifest and metadata sheets are filled properly
    Application.CalculateFullRebuild
    
    refresh_db = GetConfigParameterValue("Run Database refresh link on a fly")
    If UCase(refresh_db) = "TRUE" Then
        RefreshDBConnections
    End If
    
    'Application.ScreenUpdating = False
    'ActiveWorkbook.ForceFullCalculation
'    Application.ScreenUpdating = True
End Sub

Public Sub RefreshDBConnections()
    'refresh the database linked data
    ActiveWorkbook.RefreshAll
End Sub

Public Sub ProcessAllSpecimens()
    Dim abort As Boolean: abort = False
    Dim specimen_columns As String: specimen_columns = ""
    Dim specimen_header As String
    Dim col As Variant, cfg_row As Integer
    Dim specimen_col_orig As String
    Dim results_all() As dictionary
    Dim result_arr() As dictionary
    Dim size As Integer, size_orig As Integer, i As Integer
    Dim msg_str As String: msg_str = ""
    Dim msg_summary_str As String
    Dim total_status As Integer: total_status = vbInformation
    Dim iResponse As Integer
    
    'confirm if user want to proceed.
    iResponse = MsgBox("The system is about to start processing manfiest data on the 'shipment' tab." & vbCrLf _
                & "Be aware that some output files will be created as an otcome of the process!" _
                & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'." & vbCrLf & vbCrLf _
                & "Note: some screen flickering might occur during creating of output excel files.", _
                vbOKCancel + vbInformation, "CHARM2 Manifest Processor")
    
    If iResponse <> vbOK Then
        'exit sub based on user's response
        Exit Sub
    End If
    
    'save the "specimen Column" config parameter original value
    specimen_col_orig = GetConfigParameterValue("Specimen Column")
    
    specimen_columns = GetConfigParameterValue("Automated Specimen Columns Processing")
    
    'clean the "logs" tab
    CleanLogsWorksheet
    
    AddLogEntry "Start automated processing for the following Specimen columns '" & specimen_columns & "'.", LogMsgType.info
    
    If Len(Trim(specimen_columns)) > 0 Then
        
        'loop through given specimen columns and process one by one
        For Each col In Split(specimen_columns, ",")
            AddLogEntry "---------------------------------------------------", LogMsgType.info
            AddLogEntry "Start processing of the Specimen column '" & col & "'.", LogMsgType.info
            
            Erase result_arr
            
            col = Trim(col)
            If Len(col) > 0 Then
                If SetConfigParameterValue("Specimen Column", CStr(col)) Then
                
                    'get specimen column header
                    specimen_header = GetCellValuePerRow(1, "Specimen Column")
                    
                    AddLogEntry "Start creating manfist and metadata files for the '" & specimen_header & "' Specimen.", LogMsgType.info
                    
                    'run the process to save manifest and metadata files for the selected specimen column
                    result_arr = SavePreparedFiles(False)
                    
                    AddLogEntry "Finish creating manfist and metadata files for the '" & specimen_header & "' Specimen.", LogMsgType.info
                    
                    If ArrLength(result_arr) > 0 Then
                        AddLogEntry "Below is the summary of the prepared manifest and metadata files.", LogMsgType.info
                        'if data returned from the function call proceed here
                        size_orig = ArrLength(results_all)
                        
                        'define new size of the results_all array
                        size = size_orig + ArrLength(result_arr)
                        'set results_arr to a new size
                        ReDim Preserve results_all(size - 1)
                        
                        For i = 0 To ArrLength(result_arr) - 1
                            Set results_all(size_orig + i) = result_arr(i)
                            If Not result_arr(i) Is Nothing Then
                                'add log entry for the prepared file
                                AddLogEntry "Details of the file #" & CStr(i + 1) & " prepared for the '" _
                                    & specimen_header & "' Specimen - Status: " & result_arr(i)("status") _
                                    & " => Details: " & result_arr(i)("msg"), LogMsgType.info
                            Else
                                'add log entry for the prepared file
                                AddLogEntry "No details were received for the file #" & CStr(i + 1) & " that should be prepared for the '" _
                                    & specimen_header & "' Specimen", LogMsgType.warning
                            End If
                        Next
                    Else
                        AddLogEntry "No details were received about prepared manifest and metadata files.", LogMsgType.warning
                    End If
                Else
                    AddLogEntry "Processing of the Specimen column '" & col & "' has failed.", LogMsgType.Error
                End If
            End If
        Next
        
    End If
    
    AddLogEntry "Finish automated processing for the following Specimen columns '" & specimen_columns & "'.", LogMsgType.info
    
    Worksheets(LogsWrkSheet).Activate 'bring focus to the "logs" tab
    Worksheets(LogsWrkSheet).Cells(1, 1).Activate 'bring focus to the first cell on the sheet
    
    DisplayConfirmationMsg results_all
    
    'set the "specimen Column" config parameter back to the original value
    SetConfigParameterValue "Specimen Column", specimen_col_orig
    
End Sub

Private Sub ImportFile(ws_target As Worksheet, file_type_to_open As String)
    Dim strFileToOpen As String
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim s As Worksheet
    
    Set s = ws_target 'Worksheets(ShipmentWrkSheet)
    
    'if this is the first time files are being loaded, use the current file location as default path
    If Not firstFileLoaded Then
        ChDir Application.ActiveWorkbook.path
        firstFileLoaded = True
    End If
    
    'select a file to be loaded
    strFileToOpen = Application.GetOpenFilename _
        (Title:="Please choose a " & file_type_to_open & " file to open", _
        FileFilter:="Excel Files *.xlsx* (*.xlsx*), Excel 2003 Files *.xls* (*.xls*),")
    
    If strFileToOpen = "False" Then
        GoTo ExitMark
    End If
    
    s.Cells.Clear 'delete everything on the target worksheet
    
    CopyDataFromFile s, strFileToOpen 'copy date of the main sheet from the source file to the given ws_target sheet
    
    Select Case file_type_to_open
        Case "SHIPMENT"
            'update the first column to text values
            CleanAndUpdateColumnValuesToText s, "A"
            
            'save the name of the loaded shipment file (only name of the file will be saved, no path)
            SetConfigParameterValueByColumn "Last Loaded Shipment File", "B", Dir(strFileToOpen)
    End Select
    
    Application.ScreenUpdating = True
    
    MsgBox "CHARM " & file_type_to_open & " file " & vbCrLf _
            & strFileToOpen & vbCrLf _
            & " was successfully loaded to the '" & ws_target.Name & "' tab." & vbCrLf & vbCrLf _
            & "Note: you might need to review settings of the 'config' tab to make sure that those are set correctly. " _
            & "Please pay special attention to the highlighed rows.", vbInformation, "CHARM2 Manifest Processor"
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical
ExitMark:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub ImportShipmentFile()
    ImportFile Worksheets(ShipmentWrkSheet), "SHIPMENT"
End Sub

Public Sub ImportDemographicFile()
    Dim run_demographic As String
    
    run_demographic = GetConfigParameterValue("Lock Demographic Loading")
    If UCase(run_demographic) = "TRUE" Then
        MsgBox "Importing of the demographic file is restricted based on the value of the 'Lock Demographic Loading' configuration setting.", _
                vbCritical, "Demographic File Importing"
    Else
        ImportFile Worksheets(DemographicWrkSheet), "STUDY/DEMOGRAPHIC"
    End If
End Sub

Public Sub ImportLabStatsFile()
    ImportFile Worksheets(LabStatsWrkSheet), "LABORATORY CORRECTED STATS OF SHIPMENT"
End Sub

'This sub opens specified file and loads it contents to a specified worksheet
Sub CopyDataFromFile(ws_target As Worksheet, _
                        src_file_path As String, _
                        Optional src_worksheet_name As String = "")
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    Dim path As String
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(src_file_path, True, True)
    
    If src_worksheet_name = "" Then
        src_worksheet_name = src.Worksheets(1).Name
    End If
    
    src.Worksheets(src_worksheet_name).Cells.Copy 'copy into a clipboard
    ws_target.Cells.PasteSpecial Paste:=xlPasteAll 'paste to the worksheet 'xlPasteValues
    Application.CutCopyMode = False 'clean clipboard
    
  
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
'    'update the first column to text values
'    CleanAndUpdateColumnValuesToText ws_target, "A"
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub CleanAndUpdateColumnValuesToText(ws_target As Worksheet, update_col As String, _
                                    Optional start_row As Integer = 2, _
                                    Optional deleteRowsOfBlankValues As Boolean = True)
    
    Dim used_rows As Integer
    Dim rng As Range, cell As Range
    Dim entireRow As Range
    
    used_rows = ws_target.UsedRange.Rows.Count
    
    Set rng = ws_target.Range(update_col & start_row & ":" & update_col & used_rows)
    
    For Each cell In rng
        If deleteRowsOfBlankValues Then
            If Trim(cell.Value) = "" Then
                cell.entireRow.Delete
            Else
                cell.Value = "'" & Trim(cell.Value)
            End If
        Else
            cell.Value = "'" & Trim(cell.Value)
        End If
    Next
End Sub

Public Sub OpenHelpLink()
    Dim url As String
    
    url = GetConfigParameterValue("Help document")
    ThisWorkbook.FollowHyperlink (url)
    
End Sub

Private Function CleanLogsWorksheet()
    'clear content
    Worksheets(LogsWrkSheet).Cells.Clear
    
    'assign default alignments to the columns being used
    Worksheets(LogsWrkSheet).Columns("A").Cells.HorizontalAlignment = xlHAlignCenter
    Worksheets(LogsWrkSheet).Columns("B").Cells.HorizontalAlignment = xlHAlignCenter
    Worksheets(LogsWrkSheet).Columns("C").Cells.HorizontalAlignment = xlHAlignLeft
End Function

Private Function GetLogMsgTypeColorsDict() As dictionary
    Dim dict As New Scripting.dictionary
        
    dict.Add LogMsgType.Debugging, LogMsgTypeColor.Debugging
    dict.Add LogMsgType.info, LogMsgTypeColor.info
    dict.Add LogMsgType.warning, LogMsgTypeColor.warning
    dict.Add LogMsgType.Error, LogMsgTypeColor.Error
    
    Set GetLogMsgTypeColorsDict = dict
End Function

Private Function GetLogMsgTypeNamesDict() As dictionary
    Dim dict As New Scripting.dictionary
        
    dict.Add LogMsgType.Debugging, "Debug"
    dict.Add LogMsgType.info, "Info"
    dict.Add LogMsgType.warning, "Warning"
    dict.Add LogMsgType.Error, "Error"
    
    Set GetLogMsgTypeNamesDict = dict
End Function

Function AddLogEntry(msg As String, msg_type As LogMsgType, Optional wb As Workbook = Nothing)
    Dim row_new As Integer
    Dim log_color_dict As dictionary, log_name_dict As dictionary
    Dim font_color As LogMsgTypeColor: font_color = LogMsgTypeColor.info
    Dim log_type_name As String
    Dim entire_row As Range, non_blanks As Integer
    Dim ws_log As Worksheet
    
    If wb Is Nothing Then
        Set ws_log = Worksheets(LogsWrkSheet)
    Else
        Set ws_log = wb.Worksheets(LogsWrkSheet)
    End If
    
    'identify color to be used for the log entry
    Set log_color_dict = GetLogMsgTypeColorsDict()
    font_color = log_color_dict(msg_type)
    
    'identify name of the log type
    Set log_color_dict = GetLogMsgTypeNamesDict()
    log_type_name = log_color_dict(msg_type)
    
    'get number of the next row to be filled
    row_new = ws_log.Cells(Rows.Count, 1).End(xlUp).Row ' + 1
    
    If row_new = 1 Then
        'if the last row is the first row on the page, check if it is blank - meaning that sheet is blank and the first value is being added
        Set entire_row = ws_log.Cells(row_new, 1).entireRow 'get the first row as a range
        non_blanks = Application.WorksheetFunction.CountA(entire_row) 'count number of non blank values
        
        'if some values are present, increment the new_row by one, otherwise use the first row to put new log info into it
        If non_blanks > 0 Then
            row_new = row_new + 1
        End If
    Else
        row_new = row_new + 1
    End If
    'check if font size should be set to UpperCase for the log type idetification
    Select Case msg_type
        Case LogMsgType.warning, LogMsgType.Error
            log_type_name = UCase(log_type_name)
    End Select
    
    'log date time info for a new entry
    ws_log.Cells(row_new, 1).Value = CStr(Now)
    ws_log.Cells(row_new, 2).Value = log_type_name
    ws_log.Cells(row_new, 2).Font.Color = font_color
    ws_log.Cells(row_new, 3).Value = msg
    ws_log.Cells(row_new, 3).Font.Color = font_color
    
End Function

Function ValidateAllSpecimenTimepointColumns(Optional show_completion_notification As Boolean = True)
    Dim specimen_header As String, specimen_columns As String
    Dim col As Variant, col_letter As String
    Dim valid_result As ValidationResults
    Dim valid_total_ok As ValidationResults: valid_total_ok = ValidationResults.OK
    Dim str1 As String
    Dim msgbox_type As Integer
    Dim specimen_col_orig As String
    
    specimen_columns = GetConfigParameterValue("Automated Specimen Columns Processing")
    
    'clean the "logs" tab
    CleanLogsWorksheet
    
    'save the "specimen Column" config parameter original value
    specimen_col_orig = GetConfigParameterValue("Specimen Column")
    
    If Len(Trim(specimen_columns)) > 0 Then
        AddLogEntry "Start validating of the following Specimen columns '" & specimen_columns & "'. No output files will be prepared.", LogMsgType.info
        
        'loop through given specimen columns and process one by one
        For Each col In Split(specimen_columns, ",")
            AddLogEntry "Start validating the Specimen column '" & col & "'.", LogMsgType.info
            
            If SetConfigParameterValue("Specimen Column", CStr(col)) Then
                
                RefreshWorkbookData
                
                col_letter = Trim(CStr(col))
                If Len(col_letter) > 0 Then
                    'run validation process for the currently selected column
                    valid_result = ValidateGivenSpecimenTimepointColumn(col_letter)
                    
                    If valid_result = ValidationResults.OK Then
                        AddLogEntry "Finish validating the Specimen column '" & col_letter & "'. No issues were identified.", LogMsgType.info
                    ElseIf valid_result = ValidationResults.warning Then
                        AddLogEntry "Finish validating the Specimen column '" & col_letter & "'. Some warning were identified - see previous log entries for details.", LogMsgType.warning
                    ElseIf valid_result = ValidationResults.Error Then
                        AddLogEntry "Finish validating the Specimen column '" & col_letter & "'. Some issues were identified - see previous log entries for details.", LogMsgType.warning
                    Else
                        AddLogEntry "Finish validating the Specimen column '" & col_letter & "'. Returned validation status '" & CStr(valid_result) _
                            & "' was not recognized by the system!", LogMsgType.Error
                        valid_result = ValidationResults.Error
                    End If
                    
                    If valid_total_ok < valid_result Then
                        'update total validation outcome of the process
                        valid_total_ok = valid_result
                    End If
                End If
            Else
                AddLogEntry "Failed to validate column '" & col & "' during executing 'ValidateAllSpecimenTimepointColumns' function.", LogMsgType.Error
            End If
        Next
        
        AddLogEntry "Finish validating of the following Specimen columns '" & specimen_columns & "'.", LogMsgType.info
    Else
        AddLogEntry "No Specimen columns were provided for validation. Verify that 'Automated Specimen Columns Processing' configuration property is set with some value.", LogMsgType.warning
    End If
    
    'revert setting for Specimen Column back to the original
    SetConfigParameterValue "Specimen Column", specimen_col_orig
    
    If show_completion_notification Then
        Worksheets(LogsWrkSheet).Activate 'bring focus to the "logs" tab
        Worksheets(LogsWrkSheet).Cells(1, 1).Activate
        
        If valid_total_ok = ValidationResults.OK Then
            str1 = "SUCCESSFULLY"
            msgbox_type = vbInformation
        ElseIf valid_total_ok = ValidationResults.warning Then
            str1 = "with Warnings"
            msgbox_type = vbExclamation
        Else
            str1 = "with ERRORS/Warnings"
            msgbox_type = vbCritical
        End If
        MsgBox "Validation procedure 'Validate ALL Specimen Preparation/Timepoint Columns' was completed " & str1 _
                & ". See results on the 'logs' tab.", _
                msgbox_type, "Validation Results"
    End If
End Function

Function ValidateCurrentSpecimenTimepointColumn(Optional show_completion_notification As Boolean = True) As ValidationResults
    Dim specimen_header As String, cur_specimen_letter As String
    Dim valid_result As ValidationResults
    Dim str1 As String
    Dim msgbox_type As Integer
    
    If show_completion_notification Then
        CleanLogsWorksheet
    End If
    
    RefreshWorkbookData
    
    cur_specimen_letter = GetConfigParameterValue("Specimen Column")
    valid_result = ValidateGivenSpecimenTimepointColumn(cur_specimen_letter) 'ValidateCurrentSpecimenTimepointColumn()
    
    If show_completion_notification Then
        'show summary notification message
        Worksheets(LogsWrkSheet).Activate 'bring focus to the "logs" tab
        Worksheets(LogsWrkSheet).Cells(1, 1).Activate ' bring focus to the first cell
        
        If valid_result = ValidationResults.OK Then
            str1 = "SUCCESSFULLY"
            msgbox_type = vbInformation
        ElseIf valid_result = ValidationResults.warning Then
            str1 = "with Warnings"
            msgbox_type = vbExclamation
        Else
            str1 = "with ERRORS/Warnings"
            msgbox_type = vbCritical
        End If
        MsgBox "Validation procedure of the currently selected Specimen Preparation/Timepoint column was completed " & str1 _
                & ". See results on the 'logs' tab.", _
                msgbox_type, "Validation Results"
    End If
    
    ValidateCurrentSpecimenTimepointColumn = valid_result
End Function

Private Function ValidateGivenSpecimenTimepointColumn(col As String) As ValidationResults
    Dim specimen_header As String, specimen_letter As String
    Dim valid_result1 As ValidationResults, valid_result2 As ValidationResults, valid_result As ValidationResults
    Dim valid_result3 As ValidationResults, valid_result4 As ValidationResults, valid_result5 As ValidationResults
    
    AddLogEntry "Start validating Visit Dates.", LogMsgType.info
    valid_result5 = ValidateVisitDates
    AddLogEntry "Finish validating Visit Dates.", LogMsgType.info
    
    'get specimen column header
    specimen_header = Worksheets(ShipmentWrkSheet).Range(col & "1") 'GetCellValuePerRow(1, "Specimen Column")
    'get config value corresponding to the column letter on the shipment tab
    specimen_letter = col 'GetConfigParameterValue("Specimen Column")
    
    AddLogEntry "Start validating of the following Specimen header '" & specimen_header & "'.", LogMsgType.info
    valid_result1 = ValidateSpecimenTimepointHeader(specimen_header)
    AddLogEntry "Finish validating of the following Specimen header '" & specimen_header & "'.", LogMsgType.info
    
    AddLogEntry "Start validating of the responses of the Specimen column '" & specimen_letter & "'.", LogMsgType.info
    valid_result2 = ValidateSpecimenTimepointColumnResponses(specimen_letter)
    AddLogEntry "Finish validating of the responses of the Specimen column '" & specimen_letter & "'.", LogMsgType.info
    
    AddLogEntry "Start validating if subject's information was provided in the demographic tab.", LogMsgType.info
    valid_result3 = ValidateSubjectPresentInDemographic
    AddLogEntry "Finish validating if subject's information was provided in the demographic tab.", LogMsgType.info
    
    AddLogEntry "Start validating if aliquot ids exist in Metadata DB.", LogMsgType.info
    valid_result4 = ValidateAliquotIdVsDB
    AddLogEntry "Finish validating if aliquot ids exist in Metadata DB.", LogMsgType.info
    
    'valid_result = valid_result1 And valid_result2
    'figure out the final status
    If valid_result1 = ValidationResults.Error Or valid_result2 = ValidationResults.Error Or valid_result3 = ValidationResults.Error Or _
        valid_result4 = ValidationResults.Error Or valid_result5 = ValidationResults.Error Then
        valid_result = ValidationResults.Error
    ElseIf valid_result1 = ValidationResults.warning Or valid_result2 = ValidationResults.warning Or valid_result3 = ValidationResults.warning Or _
        valid_result4 = ValidationResults.warning Or valid_result5 = ValidationResults.warning Then
        valid_result = ValidationResults.warning
    Else
        valid_result = ValidationResults.OK
    End If
    
    ValidateGivenSpecimenTimepointColumn = valid_result
End Function

Private Function ValidateVisitDates() As ValidationResults
    Dim visit_col_name As String, ws_name As String
    Dim visit_col As Range, subject_col As Range, cell As Range
    Dim used_rows As Integer
    Dim subject_col_name As String, subject_id As String, timepoint_mark As String
    Dim out As ValidationResults
    
    out = ValidationResults.OK
    
    subject_col_name = GetConfigParameterValue("Participant Id_shipment")
    visit_col_name = GetConfigParameterValue("Visit Date")
    ws_name = GetConfigParameterValue_SheetAssignment("Visit Date")
    
    Set subject_col = Worksheets(ws_name).Range(subject_col_name & ":" & subject_col_name)
    used_rows = Worksheets(ws_name).Cells(Rows.Count, subject_col.Column).End(xlUp).Row
    Set visit_col = Worksheets(ws_name).Range(visit_col_name & "2:" & visit_col_name & CStr(used_rows))
    
    For Each cell In visit_col
        'get value of this row for the currently processed Speciment column (value "Y" is expected)
        timepoint_mark = GetCellValuePerRow(cell.Row, "Specimen Column")
        'check if the row has a "Y" entry, otherwise skip this check.
        If timepoint_mark = "Y" And Not IsDate(cell.Value) Then
            out = ValidationResults.Error
            subject_id = GetCellValuePerRow(cell.Row, "Participant Id_shipment")
            AddLogEntry "Provided 'Visit Date' value ('" & cell.Value & "') for the subject '" & subject_id & _
                "' is not a date (see row #" & CStr(cell.Row) & " on the 'shipment' tab). Due to that, a properly structured aliquot id cannot be prepared.", LogMsgType.Error
        End If
            
    Next
    
    ValidateVisitDates = out
    
End Function

Private Function ValidateSpecimenTimepointHeader(specimen_header As String) As ValidationResults
    Dim out_val As String
    Dim timepoint As String
    Dim specimen_prep As String, specimen_prep_abbr As String, specimen_prep_name As String
    Dim volume_resp As String, tissue_resp As String
    Dim arr() As String
    Dim timepoint_validation_failed As Boolean: timepoint_validation_failed = False
    Dim specimen_validation_failed As Boolean: specimen_validation_failed = False
    Dim general_validation_failed As Boolean: general_validation_failed = False

    'get timepoint and specimen prep out of the header of the specimen column
    arr = Split(Trim(specimen_header), " ")
    
    'check if the general structure of the provided header is OK
    If ArrLength(arr) = 0 Then
        general_validation_failed = True
        AddLogEntry "Provided specimen prep value is blank.", LogMsgType.Error
    End If
    If ArrLength(arr) > 1 Then
        general_validation_failed = True
        AddLogEntry "Provided specimen prep value '" & specimen_header & "' has extra component(s).", LogMsgType.Error
    End If
    
    'proceed with additional validation checks only if there was no general validation rules were OK
    If Not general_validation_failed Then
        'check if timepoint value was provided
        If ArrLength(arr) >= 1 Then
            'extract timepoint and specimen prep values
            timepoint = GetItemValueFromSpecimenHeader("timepoint") 'arr(0)
        End If
        'check if specimen_prep value was provided
        If ArrLength(arr) = 1 Then
            specimen_prep = arr(0)
        End If
        
        'convert specimen prep value to abbreviation and full name
        specimen_prep_abbr = GetSpecimenPrep(specimen_prep, True)
        specimen_prep_name = GetSpecimenPrep(specimen_prep, False)
        
        'validate specimen prep value abbreviation
        If specimen_prep_abbr = "N/D" Then
            'log warning
            AddLogEntry "Provided specimen prep value '" & specimen_prep & "' was not recognized and no abbreviation was assigned to it.", LogMsgType.Error
            specimen_validation_failed = True
        Else
            'log info
            AddLogEntry "Provided specimen prep value '" & specimen_prep _
                    & "' was recognized and will be converted to '" & specimen_prep_abbr _
                    & "' abbreviation.", LogMsgType.info
        End If
        
        'validate specimen prep value full name
        If specimen_prep_name = "N/D" Then
            'log warning
            AddLogEntry "Provided Specimen Prep value '" & specimen_prep & "' was not recognized and no name was assigned to it.", LogMsgType.Error
            specimen_validation_failed = True
        Else
            'log info
            AddLogEntry "Provided Specimen Prep value '" & specimen_prep _
                    & "' was recognized and will be converted to '" & specimen_prep_name _
                    & "' name.", LogMsgType.info
        End If
                
        'validate time point
        If Len(timepoint) > 4 Then
            'log error
            AddLogEntry "Provided Timepoint value '" & timepoint & "' is too long (longer then 4 characters).", LogMsgType.Error
            timepoint_validation_failed = True
        End If
        'validate time point
        If Len(timepoint) < 2 Then
            'log error
            AddLogEntry "Provided Timepoint value '" & timepoint & "' is too short (less then 2 characters).", LogMsgType.Error
            timepoint_validation_failed = True
        End If
        'validate time point
        If UCase(Left(timepoint, 1)) <> "V" Then
            'log error
            AddLogEntry "The Timepoint value '" & timepoint & "' has to start with letter 'V'.", LogMsgType.Error
            timepoint_validation_failed = True
        End If
        'log message, if timepoint is valid
        If Not timepoint_validation_failed Then
            AddLogEntry "The Timepoint value '" & timepoint & "' was recognized as valid.", LogMsgType.info
        End If
        
        'validate volume mapping settings
        If specimen_prep_abbr <> "N/D" Then
            volume_resp = GetSampleVolumeBySpecimenPrep(specimen_prep_abbr)
            If volume_resp = "N/D" Then
                'log warning
                AddLogEntry "'Volume' mapping was not found for the '" & specimen_prep_abbr & "' Specimen Prep abbreviation on the dictionary tab.", LogMsgType.Error
                specimen_validation_failed = True
            Else
                'log info
                AddLogEntry "An assigned 'Volume' mapping '" & CStr(volume_resp) & "' was found for the '" & specimen_prep_abbr & "' Specimen Prep abbreviation on the dictionary tab.", LogMsgType.info
            End If
            volume_resp = ""
        End If
        
        'validate volume mapping settings
        If specimen_prep_name <> "N/D" Then
            volume_resp = GetSampleVolumeBySpecimenPrep(specimen_prep_abbr)
            If volume_resp = "N/D" Then
                'log warning
                AddLogEntry "'Volume' mapping was not found for the '" & specimen_prep_name & "' Specimen Prep name on the dictionary tab.", LogMsgType.Error
                specimen_validation_failed = True
            Else
                'log info
                AddLogEntry "An assigned 'Volume' mapping '" & CStr(volume_resp) & "' was found for the '" & specimen_prep_name & "' Specimen Prep name on the dictionary tab.", LogMsgType.info
            End If
            volume_resp = ""
        End If
        
        'validate tissue mapping settings
        If specimen_prep_abbr <> "N/D" Then
            tissue_resp = GetSampleTissueBySpecimenPrep(specimen_prep_abbr)
            If tissue_resp = "N/D" Then
                'log warning
                AddLogEntry "'Tissue' mapping was not found for the '" & specimen_prep_abbr & "' Specimen Prep abbreviation on the dictionary tab.", LogMsgType.Error
                specimen_validation_failed = True
            Else
                'log info
                AddLogEntry "An assigned 'Tissue' mapping '" & CStr(tissue_resp) & "' was found for the '" & specimen_prep_abbr & "' Specimen Prep abbreviation on the dictionary tab.", LogMsgType.info
            End If
            tissue_resp = ""
        End If
        
         'validate tissue mapping settings
        If specimen_prep_name <> "N/D" Then
            tissue_resp = GetSampleTissueBySpecimenPrep(specimen_prep_abbr)
            If tissue_resp = "N/D" Then
                'log warning
                AddLogEntry "'Tissue' mapping was not found for the '" & specimen_prep_name & "' Specimen Prep name on the dictionary tab.", LogMsgType.Error
                specimen_validation_failed = True
            Else
                'log info
                AddLogEntry "An assigned 'Tissue' mapping '" & CStr(tissue_resp) & "' was found for the '" & specimen_prep_name & "' Specimen Prep name on the dictionary tab.", LogMsgType.info
            End If
            tissue_resp = ""
        End If
    End If
    
    'returns True if no validation failed, otherwise will return False
    If specimen_validation_failed Or timepoint_validation_failed Or general_validation_failed Then
        ValidateSpecimenTimepointHeader = ValidationResults.Error
    Else
        ValidateSpecimenTimepointHeader = ValidationResults.OK
    End If
    
    'ValidateSpecimenTimepointHeader = Not (specimen_validation_failed Or timepoint_validation_failed Or general_validation_failed)
End Function

Private Function ValidateSpecimenTimepointColumnResponses(specimen_column_letter As String) As ValidationResults
    Dim col As Range
    Dim column_validation_failed As ValidationResults: column_validation_failed = OK
    Dim non_blanks As Integer, y_values As Integer
    
    Set col = Worksheets(ShipmentWrkSheet).Columns(specimen_column_letter).EntireColumn
    
    non_blanks = Application.WorksheetFunction.CountA(col)
    y_values = Application.WorksheetFunction.CountIf(col, "Y")
    
    If non_blanks - 1 <> y_values Then
        'log warning
        AddLogEntry "There is(are) " & CStr(non_blanks - 1 - y_values) & " response(s) in the column '" & specimen_column_letter & "' that is(are) neither equal 'Y' no blanks.", LogMsgType.Error
        column_validation_failed = ValidationResults.Error
    ElseIf y_values = 0 Then
        AddLogEntry "There are no 'Y' responses in the column '" & specimen_column_letter & "'", LogMsgType.warning
        column_validation_failed = ValidationResults.warning
    Else
        'Log info
        AddLogEntry "All responses in the column '" & specimen_column_letter & "' are equal 'Y' or blank.", LogMsgType.info
    End If
    
    ValidateSpecimenTimepointColumnResponses = column_validation_failed
End Function

Private Function ValidateSubjectPresentInDemographic() As ValidationResults
    Dim col As Range
    Dim col_letter As String, sheet_name As String
    Dim not_found_cnt As Integer
    Dim out_value As ValidationResults
    Dim cur_col_being_processed As String
    
    On Error GoTo err_lab
    
    col_letter = GetConfigParameterValue("Subject Present In Demographic")
    sheet_name = GetConfigParameterValue_SheetAssignment("Subject Present In Demographic")
    cur_col_being_processed = GetConfigParameterValue("Specimen Column")
    
    Set col = Worksheets(sheet_name).Columns(col_letter).EntireColumn
    
    not_found_cnt = Application.WorksheetFunction.CountIf(col, "NOT FOUND!!!")
    
    If not_found_cnt > 0 Then
        out_value = ValidationResults.warning
        AddLogEntry "There are " & CStr(not_found_cnt) & " subject(s) that was(were) not found in the 'demographics' tab." _
            & " To check details of this issue, run validation separately for the column '" & cur_col_being_processed _
            & "' only and check the '" & sheet_name & "' tab for details.", LogMsgType.warning
    Else
        out_value = ValidationResults.OK
        AddLogEntry "All subject(s) were found in the 'demographics' tab.", LogMsgType.info
    End If
    
    ValidateSubjectPresentInDemographic = out_value
    Exit Function
    
err_lab:
    out_value = ValidationResults.warning
    AddLogEntry "The following error occurred during validating for presence of Subject IDs on the Demographic tab ('ValidateSubjectPresentInDemographic' function). " & Err.Description, LogMsgType.Error
    AddLogEntry "Validation procedure to check for presence of Subject IDs on the Demographic tab was not performed.", LogMsgType.warning
    ValidateSubjectPresentInDemographic = out_value
End Function

Private Function ValidateAliquotIdVsDB() As ValidationResults
    Dim col As Range
    Dim col_letter As String, sheet_name As String
    Dim exist_cnt As Integer
    Dim out_value As ValidationResults
    Dim cur_col_being_processed As String
    
    On Error GoTo err_lab
    
    col_letter = GetConfigParameterValue("Validate vs MDB")
    sheet_name = GetConfigParameterValue_SheetAssignment("Validate vs MDB")
    cur_col_being_processed = GetConfigParameterValue("Specimen Column")
    
    'Set col = Worksheets(sheet_name).Range(col_letter & ":" & col_letter)
    Set col = Worksheets(sheet_name).Columns(col_letter).EntireColumn
    
    exist_cnt = Application.WorksheetFunction.CountIf(col, "EXISTS!!!")
    
    If exist_cnt > 0 Then
        out_value = ValidationResults.warning
        AddLogEntry "There are " & CStr(exist_cnt) & " aliquot(s) that already exists in Metdata DB." _
            & " To check details of this issue, run validation separately for the column '" & cur_col_being_processed _
            & "' only and check the '" & sheet_name & "' tab for details.", LogMsgType.warning
    Else
        out_value = ValidationResults.OK
        AddLogEntry "All provided aliquots are new and do not exist in the Metadata DB yet.", LogMsgType.info
    End If
    
    ValidateAliquotIdVsDB = out_value
    Exit Function
    
err_lab:
    out_value = ValidationResults.warning
    AddLogEntry "The following error occurred during validating for existing aliquots in Metadata DB ('ValidateAliquotIdVsDB' function). " & Err.Description, LogMsgType.Error
    AddLogEntry "Validation procedure to check for existing aliquots in Metadata DB was not performed.", LogMsgType.warning
    ValidateAliquotIdVsDB = out_value
End Function

Public Sub ShowVersionMsg()
    MsgBox "CHARM2 Manifest Processor - version #" & Version _
    & vbCrLf & vbCrLf _
    & "Please send any comment and questions about this tool to 'stas.rirak@mssm.edu'", _
    vbInformation, "CHARM2 Manifest Processor"
End Sub

Public Sub Convert_LabStats_To_Shipment_File()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim s As Worksheet, ls As Worksheet
    Dim specimen_cell As String, specimen_value As String
    Dim timepoint_row As String, timepoint_col_range As String
    Dim participant_rows_range As String
    Dim participant_id_col As String
    Dim first_shipment_timepoint_col As String
    Dim cur_col As Variant
    Dim tp_rng As Range, tp_cell As Range
    Dim last_filled_shipment_row As Integer 'last filled shipment row
    Dim last_filled_row_in_timepoint As Integer
    Dim last_filled_shipment_col As Range
    Dim row_cnt As Integer, col_cnt As Integer
    Dim iResponse As Integer
    
    'confirm if user want to proceed.
    iResponse = MsgBox("The system is about to start converting 'Lab Corrected Data' present on the 'lab_stats' tab to the shipment file format." & vbCrLf _
                & "Note that any data currently located on the 'shipment' tab will be erased and replaced with the created data!" _
                & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'.", _
                vbOKCancel + vbInformation, "CHARM Conversion Lab Corrected Data")
    
    If iResponse <> vbOK Then
        'exit sub based on user's response
        Exit Sub
    End If
    
    Set s = Worksheets(ShipmentWrkSheet)
    s.Cells.Clear 'delete everything on the target worksheet
    
    Set ls = Worksheets(LabStatsWrkSheet)
    
    specimen_cell = GetConfigParameterValue("Lab_stats_specimen_cell")
    timepoint_row = GetConfigParameterValue("Lab_stats_timepoints_row")
    timepoint_col_range = GetConfigParameterValue("Lab_stats_timepoints_columns")
    participant_rows_range = GetConfigParameterValue("Lab_stats_participants_row_range")
    participant_id_col = GetConfigParameterValue("Participant Id_shipment")
    first_shipment_timepoint_col = GetConfigParameterValue("Lab_stats_shipment_timepoint_start_column")
    
    specimen_value = ls.Range(specimen_cell).Value
    last_filled_shipment_row = 1
    last_filled_row_in_timepoint = 0
    
    Set last_filled_shipment_col = s.Range(first_shipment_timepoint_col & 1)
    
    'save names of constant columns on the Shipment tab
    's.Range(participant_id_col & "1").Value = "Participant_id"
    s.Range("A1").Value = "Participant_id"
    s.Range("B1").Value = "Known CoV2 +"
    s.Range("C1").Value = "Reported Symptoms (last 14 days)"
    
    col_cnt = 1
    
    For Each cur_col In Split(timepoint_col_range, ",")
        Set tp_rng = ls.Range(cur_col & Split(participant_rows_range, ":")(0) & ":" & cur_col & Split(participant_rows_range, ":")(1))
        If col_cnt > 1 Then
            Set last_filled_shipment_col = last_filled_shipment_col.Offset(0, 1)
        End If
        row_cnt = 1
        For Each tp_cell In tp_rng
            If row_cnt = 1 Then
                s.Cells("1", last_filled_shipment_col.Column).Value = ls.Range(cur_col & timepoint_row).Value & " " & specimen_value
            End If
            If LTrim(RTrim(tp_cell.Value)) <> "" Then
                s.Range(participant_id_col & last_filled_shipment_row + row_cnt).Value = tp_cell.Value
                s.Cells(last_filled_shipment_row + row_cnt, last_filled_shipment_col.Column).Value = "Y"
                last_filled_row_in_timepoint = row_cnt
            End If
            row_cnt = row_cnt + 1
        Next
        last_filled_shipment_row = last_filled_shipment_row + last_filled_row_in_timepoint
        col_cnt = col_cnt + 1
    Next
    'set tp_range = s.Range(
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical
ExitMark:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub DuplicateCPTColumnsAsPlasma(Optional show_confirmation As Boolean = True)
    DuplicateColumnsAsPlasma "Extend Specimens to Plasma List", show_confirmation
End Sub

Sub DuplicateHeparinColumnsAsPlasma(Optional show_confirmation As Boolean = True)
    DuplicateColumnsAsPlasma "Extend Specimens to Heparin Plasma List", show_confirmation
End Sub

Sub DuplicateColumnsAsPlasma(config_param As String, Optional show_confirmation As Boolean = True)
    Dim headers As Range, header As Range
    Dim used_cols As Integer
    Dim convertSpecimensList As String, specimenList As Variant
    Dim start_cell As String
    Dim cur_specimen_abbr As String, cur_specimen_timepoint As String
    Dim copy_col_source As Range, copy_col_destin As Range
    Dim iResponse As Integer
    Dim firstCopiedColumnNum As Integer
    Dim msg1 As String
    Dim newSpecPrep As String
    Dim created_columns As Integer
    
    If show_confirmation Then
        Select Case config_param
            Case "Extend Specimens to Plasma List"
                msg1 = "The system is about to duplicate 'CPT' columns as 'Plasma(PLA)' columns on the 'shipment' tab." & vbCrLf _
                    & "The new 'PLA' columns will be created to the right of the last currently filled column. " & vbCrLf & vbCrLf _
                    & "Note: the system will create new columns (if applicable) every time this command is executed despite the fact if 'PLA' columns already were created earlier. " _
                    & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'."
                    
                newSpecPrep = "PLA"
                
            Case "Extend Specimens to Heparin Plasma List"
                msg1 = "The system is about to duplicate 'Heparin' columns as 'Plasma(HPL)' columns on the 'shipment' tab." & vbCrLf _
                    & "The new 'HPL' columns will be created to the right of the last currently filled column. " & vbCrLf & vbCrLf _
                    & "Note: the system will create new columns (if applicable) every time this command is executed despite the fact if these columns were already created earlier. " _
                    & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'."
                
                newSpecPrep = "HPL"
        End Select
    
        'if the function runs not as a part of bigger process, confirm if user want to proceed.
        iResponse = MsgBox(msg1, vbOKCancel + vbInformation, "CHARM2 Manifest Processor")
    Else
        iResponse = vbOK
    End If
    
    If iResponse <> vbOK Then
        'exit function based on user's response
        Exit Sub
    End If
    
    convertSpecimensList = GetConfigParameterValue(config_param)
    specimenList = Split(convertSpecimensList, ",")
    start_cell = GetConfigParameterValue("Extend Specimens to Plasma Starting Column #")
    
    If Not IsNumeric(start_cell) Then start_cell = "1"
    
    
    used_cols = Worksheets(ShipmentWrkSheet).UsedRange.Columns.Count
    Set headers = subRange(Worksheets(ShipmentWrkSheet).Range("1:1"), CInt(start_cell), used_cols)
    created_columns = 0
    
    For Each header In headers
        'Debug.Print (header.Address & "->" & header.Value2)
        If Len(Trim(header.Value2)) > 0 Then
            cur_specimen_abbr = GetItemValueFromSpecimenHeader("specimen_prep_abbr", header.Value2)
        Else
            cur_specimen_abbr = ""
        End If
        If IsInArray(cur_specimen_abbr, specimenList) Then
            used_cols = Worksheets(ShipmentWrkSheet).UsedRange.Columns.Count
            
            'save number of the first newly created column to report in the final message
            If firstCopiedColumnNum = 0 Then
                firstCopiedColumnNum = used_cols + 1
            End If
            
            ' set source and destination column for copying
            Set copy_col_source = Worksheets(ShipmentWrkSheet).Columns(header.Column)
            Set copy_col_destin = Worksheets(ShipmentWrkSheet).Columns(used_cols + 1)
            'perform copying procedure
            copy_col_source.Copy copy_col_destin
            
            'update header of the destination column
            copy_col_destin.Cells(1).Value = newSpecPrep
            
            created_columns = created_columns + 1
            
        End If
    Next
    
    If show_confirmation Then
        If created_columns > 0 Then
            msg1 = "Creating 'Plasma' columns on the 'shipment' tab was completed. The first created column is " _
                    & Worksheets(ShipmentWrkSheet).Columns(firstCopiedColumnNum).Address
            MsgBox msg1, vbInformation, "CHARM Manifest Processing - Create Plasma columns"
        Else
            msg1 = "No 'Plasma' columns on the 'shipment' tab were created. If this is not expected, verify the set value " & _
                "of the 'Extend Specimens to Heparin Plasma List' configuration parameter and try again."
            MsgBox msg1, vbExclamation, "CHARM Manifest Processing - Create Plasma columns"
        End If
        
    End If
End Sub

Private Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = Not IsError(Application.Match(stringToBeFound, arr, 0))
End Function

Private Function subRange(r As Range, startPos As Integer, endPos As Integer) As Range
    Set subRange = r.Parent.Range(r.Cells(startPos), r.Cells(endPos))
End Function

Public Function get_string_val(pid_val As String) As String
    get_string_val = pid_val
End Function

Public Function convertDate(date_val As String, Optional format_str As String = "yymmdd") As String
    Dim str_year As String, str_month As String, str_day As String
    Dim out As String
    
    If IsDate(date_val) Then
        out = format(date_val, format_str)
    Else
        out = ""
    End If
    convertDate = out
    
    Exit Function
    
'    If IsDate(date_val) Then
'        'check for year
'        If InStr(format_str, "yyyy") Then
'            str_year = year(date_val)
'        ElseIf InStr(format_str, "yy") Then
'            str_year = Right(year(date_val), 2)
'        Else
'            str_year = ""
'        End If
'
'        If InStr(format_str, "mm") Then
'            str_month = Right("0" & CStr(month(date_val)), 2)
'        Else
'            str_month = ""
'        End If
'
'        If InStr(format_str, "dd") Then
'            str_day = Right("0" & CStr(day(date_val)), 2)
'        Else
'            str_day = ""
'        End If
'
'        out = Replace(format_str, "yyyy", str_year)
'        out = Replace(out, "yy", str_year)
'        out = Replace(out, "mm", str_month)
'        out = Replace(out, "dd", str_day)
'    Else
'        out = ""
'    End If
'    convertDate = out
End Function

Public Function GetManifestComments(specimentPrep As String)
    Dim comment As String
    
    comment = GetCommentsBySpecimenPrep(specimentPrep)
    
    comment = comment + IIf(Len(Trim(comment)) > 0, "| ", "") + "CHARM2 Shipment: " + GetConfigParameterValueByColumn("Last Loaded Shipment File", "B")
    
    GetManifestComments = comment
End Function

'--------------------------------------
'password for protected sheets: sealfon
'--------------------------------------

