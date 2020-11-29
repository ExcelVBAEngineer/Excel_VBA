Attribute VB_Name = "modEvaluateAndClean"
Option Explicit

Sub DataFindAndOpen()
    'This subroutine searches for datasets in acceptable formats
     '----------------------------
    '1. Loop through all files in the same folder as the workbook to find dataset files (.csv, .xls, .xlsx, etc.) to open.
    '2. Open dataset as its own workbook (may need some extra steps depending on file type).
    '3. Run evaluation tool on original dataset.
    '4. Selectively copy the dataset into a new workbook (while performing the cleaning operations).
    '5. Run the evaluation tool again on the cleaned dataset workbook (to show what changed).
    '6. Export the evaluation worksheets to their own workbook.
    '7. The clean dataset would include at least four worksheets: Train, Validation, Test, and Index (for any text string features that have been numerified).
    '----------------------------
    Dim fsoLibrary As Scripting.FileSystemObject
    Dim fsoFolder As Object
    Dim fsoFile As Object
    
    Dim dctFiles As Scripting.Dictionary
    Dim dctSelect As Scripting.Dictionary
    
    Dim wbkBook As Workbook
    Dim wbkData As Workbook
    Dim wbkClean As Workbook
    
    Dim wrkSheet As Worksheet
    
    Dim varKey As Variant
    Dim varSplit As Variant
    Dim varReview As Variant
    Dim varHyper As Variant
    
    Dim strTemp() As String
    Dim strPath As String
    Dim strFile As String
    Dim strName As String
    Dim strSelect As String
    Dim strExt As String
    
    Dim lngListCt As Long
    
    Dim intMaxPctUniq As Integer
    Dim intMaxPctBlank As Integer
    Dim intPctVal As Integer
    Dim intPctTest As Integer
    
    Dim blnSplit As Boolean
    Dim blnProceed As Boolean
    
    'Get filepath of current workbook
    '[Do not need to check if path exists, as ThisWorkbook is already there]
    strPath = ThisWorkbook.Path
    
    'Set all references to the FSO library
    Set fsoLibrary = New Scripting.FileSystemObject
    Set fsoFolder = fsoLibrary.GetFolder(strPath)
    Set fsoFile = fsoFolder.Files
    Set dctFiles = New Scripting.Dictionary
    Set dctSelect = New Scripting.Dictionary
    
    Application.ScreenUpdating = False
    
    'Loop through all open workbooks and close all except this one
    For Each wbkBook In Application.Workbooks
        If Not wbkBook.Name = ThisWorkbook.Name And Not wbkBook Is Nothing Then
            wbkBook.Close False
        End If
    Next
            
    'Loop through files in the same folder
    For Each fsoFile In fsoFile
        'Acceptable file extensions are .xls, .xlsx, .xlsm, .xlsb, .csv, and .txt - all others will be rejected
        'Filename may have have periods in it, but the text following the last period should be the file extension
        strFile = fsoFile.Name
        strTemp = Split(strFile, ".")
        strExt = strTemp(UBound(strTemp))
        strName = Replace(strFile, "." & strExt, "")
        Application.StatusBar = "Searching for datasets, " & strFile
        
        'Delete previous clean files
        If Left(strFile, 6) = "CLEAN-" Then
            'First check if workbook is open, if it is, close it so it can be deleted and regenerated.
            'If Not strFile = ThisWorkbook.Name And _
            '(strExt = "csv" Or strExt = "xls" Or strExt = "xlsx" Or strExt = "xlsm" Or strExt = "xlsb") And
            'On Error GoTo 0
            
            'Delete previous iteration
            Application.ScreenUpdating = False
            fsoFile.Delete
            Application.ScreenUpdating = True
        Else
            'Add files with acceptable extensions to the file list
            If Not strFile = ThisWorkbook.Name And Not Left(strFile, 1) = "~" And _
            (strExt = "csv" Or strExt = "xls" Or strExt = "xlsx" Or strExt = "xlsm" Or strExt = "xlsb") Then
                If Not dctFiles.Exists(strFile) Then
                    dctFiles.Add strFile, strExt
                End If
            End If
        End If
    Next fsoFile
    
    '---Ask user which files to open (listbox)---
    
    'Populate listbox with filenames
    frmFileList.lstFiles.List = dctFiles.Keys
    'Display userform
    frmFileList.Show
    
    'Loop through selected results and add to a new dictionary
    For lngListCt = 0 To frmFileList.lstFiles.ListCount - 1
        If frmFileList.lstFiles.Selected(lngListCt) Then
            If Not dctSelect.Exists(frmFileList.lstFiles.List(lngListCt)) Then
                dctSelect.Add frmFileList.lstFiles.List(lngListCt), _
                dctFiles(frmFileList.lstFiles.List(lngListCt))
            End If
            'If strSelect = vbNullString Then 'First
            '    strSelect = frmFileList.lstFiles.List(1)
            'Else
            '    strSelect = strSelect & frmFileList.lstFiles.List(1)
            'End If
        End If
    Next
    
    'Reset
    'frmFileList.Unload
    strFile = vbNullString
    strExt = vbNullString
    
    intMaxPctUniq = 50 'Must be 100 or less
    'If the number of unique samples exceed 50% of the total number of samples, the feature will be flagged for deletion
    intMaxPctBlank = 20
    'If the number of blanks exceed 20% of the total number of samples, the feature will be flagged for deletion
    
    varHyper = MsgBox("Do you wish to adjust the Evaluation HYPER-PARAMETERS for blanks and unique values?" & vbCrLf & vbCrLf & _
    "Choose YES to input HYPER-parameters." & vbCrLf & "Choose NO to proceed with DECAT default values.", vbYesNo, "DECAT")
    
    'Enter user input mode - initializing the values above will pre-populate the inputbox field
    If varHyper = vbYes Then
        intMaxPctBlank = InputBox("HYPER-parameter:" & vbCrLf & "Enter maximum allowable percentage of blanks for any feature (recommended value = " & _
        intMaxPctBlank & "). " & vbCrLf & vbCrLf & _
        "Features exceeding this percentage of blanks will be flagged for later deletion.", "DECAT", intMaxPctBlank)
    
        intMaxPctUniq = InputBox("HYPER-parameter:" & vbCrLf & "Enter maximum allowable percentage of unique samples for text-based features " & _
        "which can be numerified (recommended value = " & intMaxPctUniq & "). " & _
        vbCrLf & vbCrLf & "Features exceeding this percentage of unique values will be flagged for later deletion.", "DECAT", intMaxPctUniq)
    End If
    
    'Open each selected file in Excel
    For Each varKey In dctSelect.Keys
        strFile = varKey
        strExt = dctSelect(varKey)
        strName = Replace(strFile, "." & strExt, "")
        
        If Not strFile = strPath Then
            Application.StatusBar = "Evaluating selected files: " & strFile
            
            '-------EVALUATION-------
            
            Set wbkData = Application.Workbooks.Open(strPath & "\" & strFile) 'works for both Excel and CSV formats
        
            'Generate Data Evaluation for original file
            Call DataEval(wbkData.Worksheets(1), True, intMaxPctUniq, intMaxPctBlank)
            
            Application.ScreenUpdating = True
            ThisWorkbook.Worksheets("Control Panel").Visible = xlSheetHidden
            
            varReview = MsgBox("Would you like to review the preliminary Dataset Evaluation RESULTS for " & _
                strName & "?" & vbCrLf & vbCrLf & _
                "[Note: Evaluation will be performed again after data cleaning]" & vbCrLf & vbCrLf & _
                "Choose NO to proceed with the dataset cleaning operation." & vbCrLf & _
                "Choose YES to pause the program and review the evaluation results.", vbYesNo, "DECAT")
                
            ThisWorkbook.Worksheets("Control Panel").Visible = xlSheetVisible
            Application.ScreenUpdating = False
                 
            'Pause code execution while the user reviews the evaluation results
            Select Case varReview
                Case vbYes
                    Call revpause
                Case vbNo
                    'Do nothing
            End Select
                
            'Save a copy of the dataset to perform cleaning operations
            'wbkData.SaveAs strPath & "\" & "CLEAN-" & Left(strName, 30) & "." & strExt
            
            'Create New Workbook for clean data - especially important if original dataset is a .csv file
            Application.Workbooks.Add.SaveAs (strPath & "\" & "CLEAN-" & Left(strName, 30) & "." & "xlsx")
            Set wbkClean = ActiveWorkbook
            
            'Copy data worksheet into new workbook
            wbkData.Worksheets(1).Copy before:=wbkClean.Worksheets(1)
            
            'Remove default blank sheet
            If wbkClean.Worksheets("Sheet1").Visible = xlSheetVisible Then
                Application.DisplayAlerts = False
                wbkClean.Worksheets("Sheet1").Delete
                Application.DisplayAlerts = True
            End If
            
            wbkClean.Save
        
            'Close original dataset file WITHOUT saving changes
            wbkData.Close False
            
            '-------CLEANING-------
            
            Application.StatusBar = "Cleaning selected files: " & strFile
        
            'Run cleaning operations on new workbook
            Application.ScreenUpdating = False
            Call DataCleanup(wbkClean)
            Application.ScreenUpdating = True
            
            varSplit = MsgBox("Would you like to split the " & strName & " dataset into Training, Validation, and Test Data?" & vbCrLf & vbCrLf & _
                "[Note: Data divided into separate files can also be imported into DECAT as separate datasets]" & vbCrLf & vbCrLf & _
                "Choose NO if your dataset has already been split into Training, Validation, and Test Data." & vbCrLf & vbCrLf & _
                "Choose YES to assign percentages to Validation and Test data and proceed with the split.", vbYesNo, "DECAT")
                
            Select Case varSplit
                Case vbYes
                    'Initial values to prevent infinite loop
                    intPctVal = 20
                    intPctTest = 20
                    blnProceed = False
                    
                    'Ask user for percent validation data
                    While blnProceed = False
                        intPctVal = InputBox("Enter % of data to assign to Validation Data (Recommended Value = 20 to 30, Maximum = 50)" & vbCrLf & vbCrLf & _
                        "[Note: The next window will ask for percentage of Test Data]", "DECAT", 20)
                        
                        If intPctVal > 49 Then
                            MsgBox "Validation data percentages greater than 49 percent are not allowed. Please enter an integer from 0 to 49.", vbExclamation, "DECAT"
                        ElseIf intPctVal < 0 Then
                            MsgBox "Negative percentages are not allowed. Please enter an integer from 0 to 49.", vbExclamation, "DECAT"
                        Else
                            blnProceed = True
                        End If
                    Wend
                    
                    blnProceed = False 'Reset
                    
                    'Ask user for percent test data
                    While blnProceed = False
                        intPctTest = InputBox("Enter % of data to assign to Test (recommended value = 20 to 30)" & vbCrLf & vbCrLf & _
                        "[Reminder: Validation Data = " & intPctVal & "%]" & vbCrLf & _
                        "[Note: Remaining percentage after validation and test data splits will be the training data]", "DECAT", 20)
                    
                        If intPctTest > 49 Then
                            MsgBox "Test data percentages greater than 49 percent are not allowed. Please enter an integer from 0 to 49.", vbExclamation, "DECAT"
                         ElseIf intPctTest < 0 Then
                            MsgBox "Negative percentages are not allowed. Please enter an integer from 0 to 49.", vbExclamation, "DECAT"
                        Else
                            blnProceed = True
                        End If
                    Wend
                    
                    'Split data into training, validation, and test data
                    Call TrainValTestSplit(wbkClean, intPctVal, intPctTest)
                Case vbNo
                    'Do nothing
            End Select
            
            'Pass training data worksheet within wbkData through evaluation routine
            Call DataEval(wbkClean.Worksheets(1), False, intMaxPctUniq, intMaxPctBlank)
                        
            'Move Evaluation and Numerifying Index Sheets to wbkClean
            For Each wrkSheet In ThisWorkbook.Sheets
                If Not wrkSheet.Name = "Control Panel" And _
                Not wrkSheet.Name = "Password" Then
                    wrkSheet.Move after:=wbkClean.Sheets(wbkClean.Sheets.Count)
                End If
            Next wrkSheet
        
            'Close cleaned workbook AND save changes
            'wbkClean.Save
            wbkClean.Close True
            End If
    
    Next varKey

    MsgBox "Dataset Evaluation + Cleaning Complete!!!" & vbCrLf & vbCrLf & _
        "[See CLEAN file(s) in the same folder as this program]", vbInformation, "DECAT"
        
    Application.StatusBar = ""
    Application.ScreenUpdating = True
        
    'Reset all object variables
    Set fsoLibrary = Nothing
    Set fsoFolder = Nothing
    Set fsoFile = Nothing
    Set dctFiles = Nothing
    Set dctSelect = Nothing
    Set wbkData = Nothing
    Set wbkClean = Nothing
End Sub

Sub revpause()
    'Activates userform to pause code execution
    frmPause.Show
    
    Application.ScreenUpdating = True
    
    Do Until frmPause.Visible = False
        DoEvents
    Loop
    
    Application.ScreenUpdating = False
End Sub

Sub DataEval(wrkData As Worksheet, blnFirstRun As Boolean, intMaxPctUniq As Integer, intMaxPctBlank As Integer)
    'This subroutine runs the data evaluation subroutines for each feature in the dataset
    Dim dctUnique As Scripting.Dictionary
    Dim dctSummary As Scripting.Dictionary
    'Dim wrkData As Worksheet
    Dim wrkEval As Worksheet
    Dim wrkIndex As Worksheet
    Dim strDataType As String
    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngFeat As Long
    Dim lngHeadRow As Long
    Dim lngFirFeat As Long
    Dim lngLastRow As Long
    Dim lngLastFeat As Long
    Dim lngNumSamples As Long
    Dim lngNumBlanks As Long
    Dim blnBlankCol As Boolean
    Dim blnIndex As Boolean
    
    blnBlankCol = True
        
    'Set wrkData = wbkData.Worksheets(1)
    
    '--- Find Data and Limits ---
    
    With wrkData
        'Find data
        'Find row and column of left-most header
        lngHeadRow = .UsedRange.Cells(1).Row
        lngFirFeat = .UsedRange.Cells(1).Column
        
        If lngFirFeat = 1 And lngHeadRow = 1 Then
            'First used cell is 1,1 - likely the dataset is properly arranged
            'Still need to confirm, though, that the headers are on row 1
            lngLastFeat = .Cells(1, .Columns.Count).End(xlToLeft).Column
            
            'Assume every real dataset will have at least four features
            'If the total number of "features" is two or less, then
            'It is likely additional header information above the dataset.
            'This will be flagged as an error for the user to correct.
            If lngLastFeat <= 2 Then
                MsgBox "ERROR: Did not detect enough features in Row 1." & vbCrLf & vbCrLf & "Likely there is extra header information above the dataset. " & _
                "Remove any header information that is not a feature. Blank rows and columns around the perimeter of the dataset are okay.", vbExclamation
            End If
        End If
        
        'If not (1,1) delete the extra rows and columns so it becomes (1,1)
        '[We are working with a copy of the data, so deleting unimportant rows/columns is okay]
        If lngHeadRow > 1 Then
            .Range("A1:A" & (lngHeadRow - 1)).Rows.EntireRow.Delete
            lngHeadRow = 1 'First row is now Row 1
        End If
        
        'Find last column
        lngCol = lngFirFeat
        
        'Now that the headers are in row 1, loop to the right until you run out of features
        If lngFirFeat > 1 Then
            While blnBlankCol = True
                lngCol = lngCol - 1
                
                If lngCol > 0 Then
                    If .Cells(1, lngCol).Value = vbNullString Then
                        'We know headers have been moved up to row 1
                        'Anyhing that does not have header is a blank column
                        'Blank columns should be deleted
                        .Cells(1, lngCol).EntireColumn.Delete
                    End If
                Else
                    'Yay, we finally found the header!
                    'Stop deleting blank columns!
                    blnBlankCol = False
                    lngFirFeat = 1 'First column is now Column 1
                End If
            Wend
        End If
        
        'Will find last row in FeatEval
    End With
    
    Application.ScreenUpdating = False
    
    ThisWorkbook.Activate
    
    Call CreateEvalForm(wrkData, wrkEval, wrkIndex, lngFirFeat, lngLastFeat, blnFirstRun)
        
    'Fully evaluate, summarize, and report on one feature column at a time
    For lngFeat = lngFirFeat To lngLastFeat
        Application.StatusBar = "Evaluating feature column " & wrkData.Cells(1, lngFeat).Value & _
        " (feature " & lngFeat & " of " & lngLastFeat & ")"
        
        lngLastRow = 0
        lngNumSamples = 0
        lngNumBlanks = 0
        strDataType = vbNullString
        blnIndex = False
        Set dctUnique = Nothing
        Set dctSummary = Nothing
        
        'Find last row for each feature
        lngLastRow = wrkData.Cells(wrkData.Rows.Count, lngFeat).End(xlUp).Row
        'Debug.Print lngLastRow
                            
        Call FeatEval(wrkData, strDataType, lngFeat, lngHeadRow, lngLastRow, lngNumSamples, lngNumBlanks, blnIndex)
        
        'Reset dictionaries before each run
        Set dctUnique = New Scripting.Dictionary
        Set dctSummary = New Scripting.Dictionary
        
        'Pre-define known keys for the Summary dictionary, with default values
        dctSummary.Add "Maximum", "N/A"
        dctSummary.Add "Median", "N/A"
        dctSummary.Add "Minimum", "N/A"
        dctSummary.Add "ModeOrTop", "N/A"
        dctSummary.Add "Average", "N/A"
        dctSummary.Add "St. Dev.", "N/A"
        dctSummary.Add "COV", "N/A"
        dctSummary.Add "Can Numerify", "Research Needed"
        dctSummary.Add "Num. Unique", 0
        dctSummary.Add "Action", "None"
        
        Call FeatSummary(wrkData, strDataType, lngFeat, lngLastRow, lngNumSamples, lngNumBlanks, intMaxPctUniq, intMaxPctBlank, dctUnique, dctSummary)
        
        Call ReportEval(wrkEval, strDataType, lngFeat + 1, lngNumSamples, lngNumBlanks, blnIndex, dctUnique, dctSummary)
        '+1 offset accounts for addtion of evaluation categories column
        
        If blnFirstRun = True Then
            Call GenerateNumIndex(wrkEval, wrkIndex, lngLastFeat + 1)
            '+1 offset accounts for addtion of evaluation categories column
        End If
    Next lngFeat
    
    Application.ScreenUpdating = True
    
    'Clean dataset per instructions on evaluation sheet

End Sub

Sub CreateEvalForm(wrkData As Worksheet, wrkEval As Worksheet, wrkIndex As Worksheet, lngFirFeat As Long, lngLastFeat As Long, blnFirstRun As Boolean)
    'This subroutine creates the evaluation form. Inputs are wrkData, lngFirFeat, lnglastfeat. 'Returns wrkEval.
    'wrkEval is a newly created worksheet with the headers from the dataset, ready to be filled with evaluation results.
    Dim wrkSheet As Worksheet
    Dim lngCtr As Long
    
    For Each wrkSheet In ThisWorkbook.Worksheets
        'Remove previous evaluation of whole dataset before cleaning and train-val-test split
        'Later code moves evaluations to the clean file when done, so this is not an issue with the next dataset
        If Left(wrkSheet.Name, 7) = "EvalRpt" Then
            Application.DisplayAlerts = False
            wrkSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    '--- Create Evaluation Report ---
    'Create evaluation report sheet
    Set wrkEval = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wrkEval.Name = "EvalRpt_" & Left(wrkData.Name, 12)
    
    If blnFirstRun = True Then
        'Create numerifying index sheet
        Set wrkIndex = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wrkIndex.Name = "NumIndex_" & Left(wrkData.Name, 11)
    End If
    
    wrkEval.Activate

    'Copy in headers from data sheet
    For lngCtr = lngFirFeat To lngLastFeat
        'Offset report columns one from the actual columns to allow for labeling of
        'Evaluation Report categroies
        wrkEval.Cells(1, lngCtr + 1).Value = wrkData.Cells(1, lngCtr).Value
        wrkEval.Cells(1, lngCtr + 1).Font.Bold = True
    Next
    
    With wrkEval
        .Cells(2, 1).Value = "Data Type"
        .Cells(3, 1).Value = "Num. Samples"
        .Cells(4, 1).Value = "Num. Unique" 'Number of unique entries in a feature (dctUnique.Keys.Count)
        .Cells(5, 1).Value = "Num. Blanks"
        .Cells(6, 1).Value = "Is an Index?"
        .Cells(7, 1).Value = "Can Numerify?"
        .Cells(8, 1).Value = "----------"
        .Cells(9, 1).Value = "Numeric Summary"
        .Cells(10, 1).Value = "Maximum"
        .Cells(11, 1).Value = "Median"
        .Cells(12, 1).Value = "Minimum"
        .Cells(13, 1).Value = "Most Frequent (Mode)" 'Most Frequent
        .Cells(14, 1).Value = "Average"
        .Cells(15, 1).Value = "St. Dev."
        .Cells(16, 1).Value = "COV"
        .Cells(17, 1).Value = "----------"
        .Cells(18, 1).Value = "----------"
        .Cells(19, 1).Value = "Data Science ACTION"
        .Cells(20, 1).Value = "----------"
        .Cells(21, 1).Value = "----------"
        .Cells(22, 1).Value = "Unique Values"
        
        .Cells(19, 1).Font.Bold = True
    End With

    'Autofit column widths for all rows and columns
    wrkEval.Cells().EntireColumn.AutoFit
    wrkEval.Cells().HorizontalAlignment = xlCenter
    
    'Freeze Panes to make it easier to navigate later
    wrkEval.Cells(2, 2).Select
    ActiveWindow.FreezePanes = True
    
    Set wrkSheet = Nothing
End Sub

Sub FeatEval(wrkData As Worksheet, strDataType As String, lngFeat As Long, lngHeadRow As Long, lngLastRow As Long, lngNumSamples As Long, lngNumBlanks As Long, blnIndex As Boolean)
    'This subroutine performs an evaluation of the current feature column. Inputs are wkrData, lngFeat, and lngLastRow.
    'Returns strDataType, lngNumSamples, lngNumBlanks, and blnIndex.
    Dim lngSample As Long
    Dim lngStart As Long
    Dim lngCurr As Long
    Dim lngPrev As Long
    Dim lngNDiff As Long
    Dim lngNDPrev As Long
    Dim lngStrikes As Long
    Dim dblCurr As Double
    Dim dblPrev As Double
    Dim dblNDiff As Double
    Dim dblNDPrev As Double
    Dim strSample As String
    Dim strSmpType As String
    Dim blnBlank As Boolean
    
    blnIndex = False 'Assume the feature is not an index, unless proven otherwise
    
    'Since this subroutine is called from within a loop, we need to
    'reset the counters for each new feature
    lngNumSamples = 0
    lngNumBlanks = 0
    lngStart = lngHeadRow + 1   'Default value if while loop is not used
    blnBlank = True
    
    With wrkData
        'If first several rows are blank, loop until data is found
        If .Cells(lngHeadRow + 1, lngFeat).Value = vbNullString Then
            'Need to decrement this before entering the loop so it will start at the right place
            lngStart = lngStart - 1
            
            While blnBlank = True
                lngStart = lngStart + 1
                '***Need to deal with totally blank column***
                If Trim(.Cells(lngStart, lngFeat).Value) = vbNullString Then
                    lngNumBlanks = lngNumBlanks + 1
                Else
                    blnBlank = False
                End If
            Wend
        End If
        
        'Note that lngStart pick up where the first loop left off
        For lngSample = lngStart To lngLastRow
            'We don't know the datatype yet. Safe to assume it is string for evaluation purposes.
            strSample = CStr(.Cells(lngSample, lngFeat).Value)
            
            'If the current cell is blank, increment the blank counter.
            'Otherwise, increment the sample counter
            'Also account for stray spaces
            If Trim(strSample) = vbNullString Then
                lngNumBlanks = lngNumBlanks + 1
            Else
                lngNumSamples = lngNumSamples + 1
                'Find the data type
                
                'Watch for Dates
                If InStr(1, strSample, "/") = 2 And _
                    Len(strSample) <= 10 Then
                        strSmpType = "Date"
                End If
                
                'Watch for Times
                If InStr(1, strSample, ":") > 0 And _
                    InStr(1, strSample, ":") <= 3 And _
                    InStr(1, strSample, "/") = 0 Then
                        strSmpType = "Time"
                End If
                    
                'Watch for DateTimes
                If InStr(1, strSample, ":") > 0 And _
                    InStr(1, strSample, "/") > 0 Then
                        strSmpType = "DateTime"
                End If
                
                'First check IsNumeric
                If IsNumeric(strSample) = True Then
                    'If so, is it an integer or floating point?
                    If InStr(1, Trim(strSample), ".") > 0 Then
                        strSmpType = "Float. Pt."
                    Else
                        strSmpType = "Integer"
                    End If
                End If
                            
                'Second check IsAlphabetic
                If IsAlphabetic(Trim(strSample)) = True Then
                    strSmpType = "Text String"
                    '[will assign boolean data type later, if only entries are "true" and "false"]
                    
                    'Watch for hyperlinks
                    If (InStr(1, strSample, "http") > 0) Or _
                    (InStr(1, strSample, "www") > 0) Then
                        strSmpType = "Hyperlink"
                    End If
                End If
            
                'Flag alphanumeric features separately
                If (IsNumeric(strSample) = True And IsAlphabetic(strSample) = True) Or _
                (IsNumeric(strSample) = True And InStr(1, Trim(strSample), " ") > 0) Then
                    strSmpType = "AlphaNumeric"
                    
                    'Watch for hyperlinks
                    If (InStr(1, strSample, "http") > 0) Or _
                    (InStr(1, strSample, "www") > 0) Then
                        strSmpType = "Hyperlink"
                    End If
                End If
                
                'Assume sample type is consistent throughout
                'Check to see if this is not the case
                If lngSample = (lngHeadRow + 1) Then
                    'Set first sample data type as the main data type, then check
                    strDataType = strSmpType
                    
                    'Flag for user review if there is a sample there (i.e. cell is not blank),
                    'but it did not meet any of the above criteria
                    If strDataType = vbNullString Then
                        strDataType = "Unknown"
                    End If
                Else
                    'Compare datatype of current sample to assumed datatype of the feature
                    '(as obtained from the very first sample)
                    If (strDataType <> strSmpType) And (strSmpType <> vbNullString) Then
                        If (strDataType = "Float. Pt." And strSmpType = "Integer") Or _
                        (strDataType = "Integer" And strSmpType = "Float. Pt.") Then
                            'Floating point controls
                            strDataType = "Float. Pt."
                        Else
                            'Sample type not consistent for this feature!
                            'Append new data type to data type description in order encountered
                            'Before appending, confirm that type to be appended is not already in the list.
                            '(Don't add an "and" if the first row is blank)
                            If Not InStr(1, strDataType, strSmpType) > 0 Then
                                strDataType = strDataType & " AND " & strSmpType
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Function IsAlphabetic(strTest As String) As Boolean
    Dim lngChr As Long
    Dim strChar As String
    
    IsAlphabetic = False
    
    For lngChr = 1 To Len(strTest)
        strChar = Mid(strTest, lngChr, 1)
        'Flags as alphabetic if any alphabet characters are present
        'Ignore punctuation and special characters
        If Asc(strChar) > 64 And Asc(strChar) < 90 Or _
        Asc(strChar) > 96 And Asc(strChar) < 123 Then
            IsAlphabetic = True
        End If
    Next
End Function

Sub FeatSummary(wrkData As Worksheet, strDataType As String, lngFeat As Long, lngLastRow As Long, _
lngNumSamples As Long, lngNumBlanks As Long, _
intMaxPctUniq As Integer, intMaxPctBlank As Integer, _
ByRef dctUnique As Scripting.Dictionary, ByRef dctSummary As Scripting.Dictionary)
    'This subroutine summarizes primary evaluation results. Inputs are wrkData, strDataType, lngFeat, and lngLastRow.
    'Returns dctUnique and dctSummary.
    Dim rngFeat As Range
    Dim varKey As Variant
    Dim varSamp As Variant
    Dim strCol As String
    Dim strCurr As String
    Dim strTop As String
    Dim lngSamp As Long
    Dim lngCurr As Long
    Dim lngTop As Long
    Dim lngUniq As Long
    
    With wrkData
        'Get column letter for this feature, for later use with Range
        strCol = Split(.Cells(2, lngFeat).Address, "$")(1)
        Set rngFeat = .Range(strCol & "2:" & strCol & lngLastRow)
    
        'Populate one dictionary with the list of unique samples only, for the evaluation summary
        For lngSamp = 2 To lngLastRow
            'Read current sample, avoiding errors
            If IsError(.Cells(lngSamp, lngFeat).Value) Then
                varSamp = " " 'will be taken out by trim
            Else
                varSamp = UCase(CStr(.Cells(lngSamp, lngFeat).Value))
            End If
            
            'Only add value to dictionary if it is truly unique
            'No need to count unique values while populating the dictionary - will use .Count method later
            If Not Trim(varSamp) = vbNullString Then
                If Not dctUnique.Exists(varSamp) Then
                    'Entry is unique, add to dictionary
                    dctUnique.Add varSamp, 1
                ElseIf dctUnique.Exists(varSamp) Then
                    'Identical to entry found in unique list, increment to item count (used for ModeOrTop)
                    dctUnique(varSamp) = dctUnique(varSamp) + 1
                End If
            End If
        Next
        
        If dctUnique.Count = 0 Then
            lngUniq = 0
        Else
            lngUniq = dctUnique.Count
        End If
        
        'Starting values
        strTop = vbNullString
        lngTop = 0
            
        'Most Frequent (mode or top) = Greatest number of occurrences in feature
        For Each varKey In dctUnique.Keys
            strCurr = varKey 'Unique sample
            lngCurr = dctUnique(varKey) 'Number of occurrences
                
            'Assume that first value is the maximum
            'Tiebreaker is first value encountered with the maximum count
            If lngCurr > lngTop Then
                strTop = strCurr
                lngTop = lngCurr
            End If
        Next varKey
        
        'Add result to summary
        dctSummary("ModeOrTop") = strTop
        
        'Number of unique entries in a feature
        dctSummary("Num. Unique") = lngUniq
        
        'If data type begins with AND, it probably has some blanks - remove the extra AND before processing
        If Left(strDataType, 5) = " AND " Then strDataType = Right(strDataType, Len(strDataType) - 5)
        
        'Populate the other dictionary with a summary of evaluation results
        'Statistics/Numeric Summary
        'Max        'use Application.WorksheetFunction
        'Median     'use Application.WorksheetFunction
        'Min        'use Application.WorksheetFunction
        'Mode       'use Application.WorksheetFunction
        'Average    'use Application.WorksheetFunction
        'St. Dev.   'use Application.WorksheetFunction
        'COV        =St.Dev./Average
                
        If strDataType = "Integer" Or strDataType = "Float. Pt." Then
            '[Numeric data only]
            dctSummary("Maximum") = Application.WorksheetFunction.Max(rngFeat) 'For instance, .Range(A1:A512)
            dctSummary("Median") = Application.WorksheetFunction.Median(rngFeat)
            dctSummary("Minimum") = Application.WorksheetFunction.Min(rngFeat)
            dctSummary("Average") = Round(Application.WorksheetFunction.Average(rngFeat), 2)
            dctSummary("St. Dev.") = Round(Application.WorksheetFunction.StDev_P(rngFeat), 2)
            dctSummary("COV") = Round((CDbl(dctSummary("St. Dev.")) / CDbl(dctSummary("Average"))), 3)
            dctSummary("Can Numerify") = "N/A" 'Already a number
            
            '---The following line throws an error - it has been superseded by the second loop above---
            'dctSummary.Add "ModeOrTop", Application.WorksheetFunction.Mode_Sngl(rngFeat)

        ElseIf strDataType = "Text String" Then
            '[Alphabetic data only]
            '[If number of unique is the same as the number of samples, every entry is unique]
            '[If every entry is unique, further research is needed to make it actionable feature]
            If lngNumSamples = lngUniq Then
                dctSummary("Can Numerify") = "Research Needed"
            Else
                'Can numerify? [No if above condition is true, Yes if otherwise]
                '[To be more robust create HYPERparameter for maximum % unique to numerify]
                If ((lngUniq / lngNumSamples) * 100) > 0 And _
                ((lngUniq / lngNumSamples) * 100) <= intMaxPctUniq Then
                    dctSummary("Can Numerify") = "YES"
                Else
                    dctSummary("Can Numerify") = "NO"
                End If
            End If
            
            'Text string features will not have any of the following statistics
            '---The default values previously assigned at dictionary creation will be used---
            'dctSummary.Add "Maximum", "N/A"
            'dctSummary.Add "Median", "N/A"
            'dctSummary.Add "Minimum", "N/A"
            'dctSummary.Add "Average", "N/A"
            'dctSummary.Add "St. Dev.", "N/A"
            'dctSummary.Add "COV", "N/A"
            
        ElseIf strDataType = "AlphaNumeric" Then
            '[Alphabetic data only]

            '[If the above is the same as the number of samples, every entry is unique]
            '[If every entry is unique, further research is needed to make it actionable feature]
            If lngNumSamples = lngUniq Then
                dctSummary("Can Numerify") = "Research Needed"
            Else
                'Can numerify? [No if above condition is true, Yes if otherwise]
                '
                If ((lngUniq / lngNumSamples) * 100) > 0 And _
                ((lngUniq / lngNumSamples) * 100) <= intMaxPctUniq Then
                    dctSummary("Can Numerify") = "YES"
                Else
                    dctSummary("Can Numerify") = "NO"
                End If
            End If
            
            'AlphaNumeric features will not have any of the following statistics
            '---The default values previously assigned at dictionary creation will be used---
            'dctSummary.Add "Maximum", "N/A"
            'dctSummary.Add "Median", "N/A"
            'dctSummary.Add "Minimum", "N/A"
            'dctSummary.Add "Average", "N/A"
            'dctSummary.Add "St. Dev.", "N/A"
            'dctSummary.Add "COV", "N/A"
        
        
        ElseIf strDataType = "Date" Or strDataType = "Time" Or strDataType = "DateTime" Then
            Select Case strDataType
                Case "Date"
                    dctSummary("Can Numerify") = "BY PARSING"
                    dctSummary("Action") = "PARSE M/D/Y"
                Case "Time"
                    dctSummary("Can Numerify") = "BY PARSING"
                    dctSummary("Action") = "PARSE H:M:S"
                Case "DateTime"
                    dctSummary("Can Numerify") = "BY PARSING"
                    dctSummary("Action") = "PARSE DateTime"
            End Select
            
            dctSummary("Maximum") = "TBD"
            dctSummary("Median") = "TBD"
            dctSummary("Minimum") = "TBD"
            dctSummary("Average") = "TBD"
            dctSummary("St. Dev.") = "TBD"
            dctSummary("COV") = "TBD"
            
        Else    '=Unknown
            '---The default values previously assigned at dictionary creation will be used---
            'dctSummary.Add "Maximum", "N/A"
            'dctSummary.Add "Median", "N/A"
            'dctSummary.Add "Minimum", "N/A"
            'dctSummary.Add "Average", "N/A"
            'dctSummary.Add "St. Dev.", "N/A"
            'dctSummary.Add "COV", "N/A"
            'dctSummary.Add "Can Numerify", "Research Needed"
        End If
    End With

    'Assign boolean data type for Text String, if only unqiue entries are true/false, yes/no, on/off, or zero/one
    If dctUnique.Count = 2 Then
        If (dctUnique.Exists("TRUE") And dctUnique.Exists("FALSE")) Or _
        (dctUnique.Exists("YES") And dctUnique.Exists("NO")) Or _
        (dctUnique.Exists("ON") And dctUnique.Exists("OFF")) Or _
        (dctUnique.Exists("0") And dctUnique.Exists("1")) Then
            strDataType = "Boolean"
        End If
    End If
    
    'Determine action to be taken
    Select Case dctSummary("Can Numerify")
        Case "YES"
            dctSummary("Action") = "GO for DS Use"
        Case "NO"
            dctSummary("Action") = "DELETE (cannot numerify)"
        Case "Research Needed"
            If strDataType = "Hyperlink" Then
                dctSummary("Action") = "DELETE (hyperlink)"
            Else
                dctSummary("Action") = "DELETE (Needs MANUAL EDIT)"
            End If
        Case Else
            'Integers and floating points are N/A for numerifying, but they are still very much valid for Data Science
            If dctSummary("Action") = "None" And _
            (strDataType = "Integer" Or strDataType = "Float. Pt." Or strDataType = "Boolean") Then
                dctSummary("Action") = "GO for DS Use"
            End If 'Do nothing - especially for dates and times
            
            'Identify columns that have the same value repeated
            If (dctSummary("Num. Unique") = 1) Then
                dctSummary("Action") = "DELETE (no unique values)"
            End If
            
            'Need to add number of (non-blank) samples and number of blanks to get total rows
            If ((lngNumBlanks / (lngNumSamples + lngNumBlanks)) * 100) > 0 And _
            ((lngNumBlanks / (lngNumSamples + lngNumBlanks)) * 100) >= intMaxPctBlank Then
                dctSummary("Action") = "DELETE (too many blanks)"
            End If
    End Select
End Sub

Sub ReportEval(wrkEval As Worksheet, strDataType As String, lngFeat As Long, lngNumSamples As Long, lngNumBlanks As Long, blnIndex As Boolean, _
ByRef dctUnique As Scripting.Dictionary, ByRef dctSummary As Scripting.Dictionary)
     Dim varKey2 As Variant
     Dim lngOutRow As Long
     
     'Need HYPERparameters for gradation between zero and 100 to make robust
     'Compare unique count to total number of samples, to make a second pass on the index evaluation
     'By definition, an index column have a unique indentifier for every sample
     If blnIndex = False And _
     (lngNumSamples = dctSummary("Num. Unique")) Then
        blnIndex = True
     End If
    
     'Override any previous actions if it is an index column
     If blnIndex = True Then
        If Not dctSummary.Exists("Action") Then
            dctSummary.Add "Action", "DELETE (index)"
            dctSummary("Can Numerify") = "N/A"
        Else
            dctSummary("Action") = "DELETE (index)"
            dctSummary("Can Numerify") = "N/A"
        End If
     End If
     
     'For data science purposes, a feature where all samples have the same value is equally as useless as an index
     If dctSummary("Num. Unique") = 1 Then
        If Not dctSummary.Exists("Action") Then
            dctSummary.Add "Action", "DELETE (same value for all samples)"
            dctSummary("Can Numerify") = "N/A"
        Else
            dctSummary("Action") = "DELETE (same value for all samples)"
            dctSummary("Can Numerify") = "N/A"
        End If
     End If

     With wrkEval
        'Report evaulation results in current feature column in report sheet
        '1 is the Header Row
        .Cells(2, lngFeat).Value = strDataType    'Data Type (Integer, Floating Point, Text, Date, Time, Boolean, etc.)
        .Cells(3, lngFeat).Value = lngNumSamples  'Number of Actual Samples
        .Cells(4, lngFeat).Value = dctSummary("Num. Unique") 'Number of unique entries in a feature
        .Cells(5, lngFeat).Value = lngNumBlanks   'Number of Blanks
               
        Select Case blnIndex 'Is it a sequential? [Recommend deleting]
            Case True 'Column is indeed an index column
                .Cells(6, lngFeat).Value = "YES"
            Case False 'Column is not an index column
                .Cells(6, lngFeat).Value = "NO"
        End Select
        
        .Cells(7, lngFeat).Value = dctSummary("Can Numerify")
        
        .Cells(8, lngFeat).Value = "----------"
        
        'Statistics/Numeric Summary
        .Cells(10, lngFeat).Value = dctSummary("Maximum")
        .Cells(11, lngFeat).Value = dctSummary("Median")
        .Cells(12, lngFeat).Value = dctSummary("Minimum")
        .Cells(13, lngFeat).Value = dctSummary("ModeOrTop")
        .Cells(14, lngFeat).Value = dctSummary("Average")
        .Cells(15, lngFeat).Value = dctSummary("St. Dev.")
        .Cells(16, lngFeat).Value = dctSummary("COV")
        .Cells(17, lngFeat).Value = "----------"
        .Cells(18, lngFeat).Value = "----------"
        .Cells(19, lngFeat).Value = dctSummary("Action")
        .Cells(20, lngFeat).Value = "----------"
        .Cells(21, lngFeat).Value = "----------"
        
        .Cells(19, lngFeat).Font.Bold = True
        
        lngOutRow = 22 'where to begin the list of unique values
                
        'Summarize all unique entries
        For Each varKey2 In dctUnique.Keys
            .Cells(lngOutRow, lngFeat).Value = varKey2
            lngOutRow = lngOutRow + 1
        Next varKey2
        
        'Add a unique value to account for blanks, if present in the feature
        If lngNumBlanks > 0 Then
            .Cells(lngOutRow, lngFeat).Value = "---BLANK---"
        End If
                    
        'Autofit column widths for all rows and columns
        .Cells().EntireColumn.AutoFit
    End With
        'The report will serve as the "work order" that tells cleanup what to do
        '[For cleanup side, add HYPERparameter to determine what to do based on % blanks = (blanks/(samples+blanks))*100
End Sub

Sub GenerateNumIndex(wrkEval As Worksheet, wrkIndex As Worksheet, lngLastFeat As Long)
    'Create a separate sheet summarizing the index for each text string feature to be numerified
    'Only runs for evaluation of original dataset
    Dim lngFeat As Long
    Dim lngUniq As Long
    Dim lngCol As Long
    Dim lngRow As Long
    
    lngCol = 2
    
    With wrkEval
        For lngFeat = 2 To lngLastFeat
            'Reset row counters
            lngUniq = 21
            lngRow = 2
    
            'Look at the "Can Numerify?" row of evaluation to determine action to be taken
            If .Cells(7, lngFeat).Value = "YES" Then
                'Transfer header to first row
                wrkIndex.Cells(1, lngCol).Value = .Cells(1, lngFeat).Value
                
                'Copy list of unique values
                While .Cells(lngUniq, lngFeat).Value <> vbNullString
                    lngUniq = lngUniq + 1
                    wrkIndex.Cells(lngRow, lngCol - 1).Value = lngRow - 1 'Index
                    wrkIndex.Cells(lngRow, lngCol).Value = .Cells(lngUniq, lngFeat).Value 'Unique item
                    lngRow = lngRow + 1
                Wend
                
                'Remove last index number, which corresponds to the blank row that stopped the loop (sample is blank)
                wrkIndex.Cells(lngRow - 1, lngCol - 1).Value = vbNullString
                
                lngCol = lngCol + 3
            End If
        Next lngFeat
    End With
End Sub

Sub TrainValTestSplit(wbkClean As Workbook, intPctVal As Integer, intPctTest As Integer)
    'This subroutine splits clean data into train, test, and validation
    Dim wrkTrain As Worksheet
    Dim wrkVal As Worksheet
    Dim wrkTest As Worksheet
    Dim strName As String
    Dim strColLtr As String
    Dim lngFeat As Long
    Dim lngLastRow As Long
    Dim lngMaxLastRow As Long
    Dim lngLastFeat As Long
    Dim lngTestStart As Long
    Dim lngValStart As Long
    
    strName = wbkClean.Name
    
    lngMaxLastRow = 1
    
    With wbkClean.Worksheets(1)
        lngLastFeat = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        strColLtr = .Cells(1, lngLastFeat).Address
        strColLtr = Replace(Left(strColLtr, Len(strColLtr) - 1), "$", "")
    
        'Find maximum last row of dataset (should be the same for all features now that cleaning operation is done)
        For lngFeat = 1 To lngLastFeat
            lngLastRow = .Cells(.Rows.Count, lngFeat).End(xlUp).Row
            
            If lngLastRow > lngMaxLastRow Then
                'Promote current row to the maximum, as it is greater than the previous maximum
                lngMaxLastRow = lngLastRow
            End If
        Next
        
        'Now we know the maximum row, need to and scale proportionally to get starting point for split
        'Data was already randomized in the cleaning process
        'First chunk to removed from the bottom of the dataset will be the test data
        lngTestStart = lngLastRow - (lngLastRow * (intPctTest / 100))
        'Second chunk to removed, just above the test data, will be the validation data
        lngValStart = lngTestStart - (lngLastRow * (intPctVal / 100))
        'All remaining data not removed will be the training data
    End With
    
    'Rename clean full dataset as the training data, then move the validation and test data out
    Set wrkTrain = wbkClean.Worksheets(1)
    wrkTrain.Name = "TRAIN_" & Left(strName, 25)
    'Create worksheet for validation data
    Set wrkVal = wbkClean.Worksheets.Add(after:=wbkClean.Worksheets(1))
    wrkVal.Name = "VALIDATE_" & Left(strName, 22)
    'Create worksheet for test data
    Set wrkTest = wbkClean.Worksheets.Add(after:=wbkClean.Worksheets(2))
    wrkTest.Name = "TEST_" & Left(strName, 26)
    
    'Copy and paste feature headers
    wrkTrain.Range("A1:" & strColLtr & "1").Copy
    wrkVal.Range("A1:" & strColLtr & "1").PasteSpecial xlPasteValues
    wrkTest.Range("A1:" & strColLtr & "1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False 'Clear clipboard
    
    'Cut and paste data to Test Data
    wrkTrain.Range(lngTestStart & ":" & lngLastRow).Cut wrkTest.Range("2:" & (lngLastRow - lngTestStart))
    Application.CutCopyMode = False 'Clear clipboard
    
    'Cut and paste data to Validation Data
    wrkTrain.Range(lngValStart & ":" & lngTestStart - 1).Cut wrkVal.Range("2:" & (lngTestStart - lngValStart))
    Application.CutCopyMode = False 'Clear clipboard
End Sub


Sub DataCleanup(wbkClean As Workbook)
    'This subroutine performs data cleanup operations on a previously-copied version of the dataset (wbkClean)
    'Performing the cleaning operations:
    '--------------------------------------
    '1. Fill in blanks using average
    '2. Numerify all text columns
    '3. Create random number feature
    '4. Sort by random number feature
    '5. Delete random number column
    '6. Delete columns flagged for deletion
    '--------------------------------------
    Dim dctNumerify As Scripting.Dictionary
    Dim dctBlankFill As Scripting.Dictionary
    Dim dctDelete As Scripting.Dictionary
    Dim dctDateTime As Scripting.Dictionary
    Dim dctTemp As Scripting.Dictionary
    
    Dim wrkDataset As Worksheet
    Dim wrkEval As Worksheet
    Dim wrkIndex As Worksheet
    Dim wrkSheet As Worksheet
    
    Dim varKey As Variant
    Dim varSubKey As Variant
    Dim varBlankFill As Variant
    
    Dim strDataFeat As String
    Dim strFeat As String
    Dim strKey As String
    Dim strItem As String
    Dim strCurr As String
    Dim strOut As String
    Dim strColLtr As String
    Dim strDateTime As String
    Dim strTemp() As String
    Dim strTemp2() As String
    
    Dim lngFeatCtr As Long
    Dim lngRow As Long
    Dim lngLastRow As Long
    Dim lngMaxLastRow As Long
    Dim lngLastEvFeat As Long
    Dim lngDataFeat As Long
    Dim lngLastFeat As Long
    Dim lngNumBlanks As Long
    
    Dim blnDate As Boolean
    Dim blnTime As Boolean
    Dim blnDateTime As Boolean

    Randomize
    
    Application.StatusBar = "Performing dataset cleaning operations "
    
    'Initialize outer dictionary (will initialize the inner one within a loop)
    Set dctNumerify = New Scripting.Dictionary
    Set dctDelete = New Scripting.Dictionary
    Set dctDateTime = New Scripting.Dictionary
    
    'Find numerifying index sheet and evaluation report sheet, which are still in ThisWorkbook when the code gets here
    For Each wrkSheet In ThisWorkbook.Sheets
        If Left(wrkSheet.Name, 8) = "NumIndex" Then
            Set wrkIndex = wrkSheet
        ElseIf Left(wrkSheet.Name, 7) = "EvalRpt" Then
            Set wrkEval = wrkSheet
        End If
    Next wrkSheet
    
    'Assign dataset, currently the only worksheet in wbkClean, to an object variable
    Set wrkDataset = wbkClean.Worksheets(1)
    
    lngFeatCtr = -1 'Need to start on column 2
    
    '---Populate numerifying dictionary---
    With wrkIndex
        'Loop through numerifying index
        While .Cells(1, lngFeatCtr + 3) <> vbNullString
            lngFeatCtr = lngFeatCtr + 3 'Follows pattern by which the numerifying index was created
            
            ' Avoid adding any blanks to the dictionary, in case the while loop runs one more time after finding a blank
            If Not .Cells(1, lngFeatCtr).Value = vbNullString Then
                'Add feature and unique values to nested dictionary
                'Find last row of unique values
                lngLastRow = .Cells(.Rows.Count, lngFeatCtr).End(xlUp).Row
                
                'Populate dictionary by looping though unique entries
                For lngRow = 2 To lngLastRow
                    'No need to check if value exists, as all values should be unique (from previous dictionary operations)
                    'Creatively assign key to use only one dictionary (instead of nested dictionaries
                    strKey = .Cells(1, lngFeatCtr).Value & ";" & .Cells(lngRow, lngFeatCtr).Value 'Feature Name;Unique Sample Name
                    strItem = .Cells(lngRow, lngFeatCtr - 1).Value 'Number to replace unique sample name with
                        
                    'Add dictionary to larger dictionary all index sets
                    dctNumerify.Add strKey, strItem
                        
                    'Debug.Print .Cells(1, lngFeatCtr).Value & "_" & .Cells(lngRow, lngFeatCtr).Value & "_" & dctFeature(.Cells(lngRow, lngFeatCtr).Value)
                Next lngRow
            End If
        Wend
        
        ''Report on dictionary values
        'For Each varKey In dctNumerify.Keys
        '    Debug.Print varKey & "_" & dctNumerify(varKey)
        'Next varKey
    End With
    
    '---Populate dictionary with fillers for blanks (average or most frequent)---
    With wrkEval
        lngLastEvFeat = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        Set dctBlankFill = New Scripting.Dictionary
        
        'Loop through all features in the evaluation report
        For lngFeatCtr = 2 To lngLastEvFeat
            'If there are any blanks, populate the dictionary with what should take their place
            lngNumBlanks = .Cells(5, lngFeatCtr).Value
            
            'List columns for deletion
            If Left(.Cells(19, lngFeatCtr).Value, 6) = "DELETE" Then
                dctDelete.Add .Cells(1, lngFeatCtr).Value, lngFeatCtr - 1 'Feature name, worksheet column in dataset workbook
            End If
            
            'Make note of any date or time columns to be parsed
            If .Cells(2, lngFeatCtr).Value = "Date" Or _
            .Cells(2, lngFeatCtr).Value = "Time" Or _
            .Cells(2, lngFeatCtr).Value = "DateTime" Then
                dctDateTime.Add .Cells(1, lngFeatCtr).Value, .Cells(2, lngFeatCtr).Value
            End If
            
            'Avoid features that will be deleted when evaluating for blank fill
            If lngNumBlanks > 0 And Left(.Cells(19, lngFeatCtr).Value, 7) <> "DELETE" Then
            'And Not dctBlankFill.Exists(.Cells(1, lngFeatCtr).Value) Then
                strFeat = .Cells(1, lngFeatCtr).Value
        
                Select Case .Cells(2, lngFeatCtr).Value
                    Case "Integer"
                        dctBlankFill.Add strFeat, CLng(Round(.Cells(14, lngFeatCtr).Value, 0)) 'Feature name, Average rounded to integer
                    Case "Float. Pt."
                        dctBlankFill.Add strFeat, CDbl(.Cells(14, lngFeatCtr).Value) 'Feature name, Average
                    Case "Text String"
                        dctBlankFill.Add strFeat, "---BLANK---" 'Feature name, Unique value for blanks
                    Case "Boolean"
                        dctBlankFill.Add strFeat, "---BLANK---" 'Feature name, Unique value for blanks
                    Case "Date"
                        dctBlankFill.Add strFeat, CStr(.Cells(13, lngFeatCtr).Value) 'Feature name, Unique value for blanks
                    Case "Time"
                        dctBlankFill.Add strFeat, CStr(.Cells(13, lngFeatCtr).Value) 'Feature name, Unique value for blanks
                    Case "DateTime"
                        dctBlankFill.Add strFeat, CStr(.Cells(13, lngFeatCtr).Value) 'Feature name, Unique value for blanks
                End Select
        
                Debug.Print strFeat & "____" & dctBlankFill(strFeat)
            End If
        Next lngFeatCtr
    End With
    
    '---Loop through whole dataset---
    With wrkDataset
        'Initialize
        lngMaxLastRow = 0
        
        'Find last feature
        lngLastFeat = .Cells(1, .Columns.Count).End(xlToLeft).Column
        'Loop through all features (looping in reverse avoids reference issues due to deletion)
        For lngDataFeat = lngLastFeat To 1 Step -1
            ''Reset
            'varBlankFill = "---BLANK---" '
            
            'Identify feature name
            strDataFeat = .Cells(1, lngDataFeat).Value
            
            '6. Delete columns flagged for deletion
            'First, check if column needs to be deleted. No point in continuing this iteration if it is.
            If dctDelete.Exists(strDataFeat) Then
                'Feature is on the list for deletion - delete it and exit this iteration of the loop
                .Cells(1, lngDataFeat).EntireColumn.Delete
            Else
                'Find last row for current feature
                lngLastRow = .Cells(.Rows.Count, lngDataFeat).End(xlUp).Row
                
                'Find largest value of last row for later use with Sort
                If lngLastRow > lngMaxLastRow Then
                    lngMaxLastRow = lngLastRow
                End If
                
                'Reset
                blnDate = False
                blnTime = False
                blnDateTime = False
                
                'Identify date and time features
                If dctDateTime.Exists(strDataFeat) Then
                    'Insert two columns
                    .Columns(lngDataFeat + 1).Insert xlRight, xlFormatFromLeftOrAbove
                    .Columns(lngDataFeat + 1).Insert xlRight, xlFormatFromLeftOrAbove
                    
                    Select Case dctDateTime(strDataFeat)
                        Case "Date"
                            blnDate = True
                            'Populate headers
                            .Cells(1, lngDataFeat) = "Month"
                            .Cells(1, lngDataFeat + 1) = "Day"
                            .Cells(1, lngDataFeat + 2) = "Year"
                            
                        Case "Time"
                            blnTime = True
                            'Populate headers
                            .Cells(1, lngDataFeat) = "Hours"
                            .Cells(1, lngDataFeat + 1) = "Minutes"
                            .Cells(1, lngDataFeat + 2) = "Seconds"
                            
                        Case "DateTime"
                            blnDateTime = True
                            'Insert three more columns
                            .Columns(lngDataFeat + 1).Insert xlRight, xlFormatFromLeftOrAbove
                            .Columns(lngDataFeat + 1).Insert xlRight, xlFormatFromLeftOrAbove
                            .Columns(lngDataFeat + 1).Insert xlRight, xlFormatFromLeftOrAbove
                            'Populate headers
                            .Cells(1, lngDataFeat) = "Month"
                            .Cells(1, lngDataFeat + 1) = "Day"
                            .Cells(1, lngDataFeat + 2) = "Year"
                            .Cells(1, lngDataFeat + 3) = "Hours"
                            .Cells(1, lngDataFeat + 4) = "Minutes"
                            .Cells(1, lngDataFeat + 5) = "Seconds"
                    End Select
                End If
                
                Application.StatusBar = "Dataset Cleaning " & _
                "(Feature " & (lngLastFeat - lngDataFeat) & " of " & lngLastFeat & ")"
                    
                'Loop through all samples in the feature
                For lngRow = 2 To lngLastRow
                    
                    'Check for Excel errors in cells, replace with blanks
                    '(which will then be filled in the next If statement)
                    If IsError(.Cells(lngRow, lngDataFeat).Value) Then
                        .Cells(lngRow, lngDataFeat).Value = vbNullString
                    End If
                
                    '1. Fill in blanks
                    If .Cells(lngRow, lngDataFeat).Value = vbNullString Then
                        'Identify correct blankfill, if any
                        If dctBlankFill.Exists(strDataFeat) Then
                            varBlankFill = dctBlankFill(strDataFeat)
                        End If
                        
                        'Apply Blankfill
                        .Cells(lngRow, lngDataFeat).Value = varBlankFill
                        
                        'Also numerify blanks, referenced by the unique sample name of "---BLANK---"
                        strCurr = strDataFeat & ";" & UCase(varBlankFill)

                    ElseIf .Cells(lngRow, lngDataFeat).Value <> vbNullString Then
                        'Concatenate feature name to unique sample name to make the right key
                        'Unique sample name needs to make uppercase to match format of the dictionary
                        strCurr = strDataFeat & ";" & UCase(.Cells(lngRow, lngDataFeat).Value)
                        
                    End If
                    
                    'Identify numerifiable features
                    If dctNumerify.Exists(strCurr) = True Then
                        '2. Numerify all text columns
                        strOut = dctNumerify(strCurr)
                        .Cells(lngRow, lngDataFeat).Value = strOut
                    End If
                    
                    If blnDate = True Then
                        strDateTime = .Cells(lngRow, lngDataFeat).Value
                        strTemp = Split(strDateTime, "/")
                        .Cells(lngRow, lngDataFeat) = strTemp(0)
                        .Cells(lngRow, lngDataFeat + 1) = strTemp(1)
                        .Cells(lngRow, lngDataFeat + 2) = strTemp(2)
                        'Need to parse out before reformatting
                        'Reformat from date to number
                        .Cells(lngRow, lngDataFeat).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 1).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 2).NumberFormat = "0"
                    End If
                    
                    If blnTime = True Then
                        strDateTime = .Cells(lngRow, lngDataFeat).Value
                        strTemp = Split(strDateTime, ":")
                        .Cells(lngRow, lngDataFeat) = strTemp(0)
                        .Cells(lngRow, lngDataFeat + 1) = strTemp(1)
                        .Cells(lngRow, lngDataFeat + 2) = strTemp(2)
                        'Need to parse out before reformatting
                        'Reformat from date to number
                        .Cells(lngRow, lngDataFeat).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 1).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 2).NumberFormat = "0"
                    End If
                    
                    If blnDateTime = True Then
                    strDateTime = .Cells(lngRow, lngDataFeat).Value
                        strTemp = Split(strDateTime, "/")
                        strTemp2 = Split(UBound(strTemp), ":")
                        .Cells(lngRow, lngDataFeat) = strTemp(0)
                        .Cells(lngRow, lngDataFeat + 1) = strTemp(1)
                        .Cells(lngRow, lngDataFeat + 2) = strTemp(2)
                        .Cells(lngRow, lngDataFeat + 3) = strTemp2(0)
                        .Cells(lngRow, lngDataFeat + 4) = strTemp2(0)
                        .Cells(lngRow, lngDataFeat + 5) = strTemp2(0)
                        'Need to parse out before reformatting
                        'Reformat from date to number
                        .Cells(lngRow, lngDataFeat).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 1).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 2).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 3).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 4).NumberFormat = "0"
                        .Cells(lngRow, lngDataFeat + 5).NumberFormat = "0"
                    End If
                    
                    '3. Create random number feature (concurrently with operations on last column)
                    If lngDataFeat = lngLastFeat Then
                        .Cells(lngRow, lngLastFeat + 1).Value = 10 * Rnd()
                    End If
                    
                Next lngRow
            End If
        Next lngDataFeat
        
        'Get column letter of random number column for later sorting (to randomize)
        strColLtr = .Cells(1, lngLastFeat + 1 - dctDelete.Count).Address
        strColLtr = Replace(Left(strColLtr, Len(strColLtr) - 1), "$", "")
        
        '4. Sort by random number feature
        '5. Delete random number column
        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 Key:=Range(strColLtr & "2:" & strColLtr & lngMaxLastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SetRange Range("A1:" & strColLtr & lngMaxLastRow)
        .Sort.Header = xlYes
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.SortMethod = xlPinYin
        .Sort.Apply
        .Columns(strColLtr & ":" & strColLtr).Delete Shift:=xlToLeft

    End With
    
    '7. Ask user about train-val-test split
    
    wbkClean.Save
    
    'Free up memory
    Set dctNumerify = Nothing
    Set dctBlankFill = Nothing
    Set dctDelete = Nothing
    Set dctDateTime = Nothing
    Set dctTemp = Nothing
    Set wrkDataset = Nothing
    Set wrkEval = Nothing
    Set wrkIndex = Nothing
    Set wrkSheet = Nothing
End Sub
