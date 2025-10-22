Attribute VB_Name = "Module2"
' =================================================================================
' Module: Helper Functions for Rule Conditions
' =================================================================================

' Function to extract details for olConditionFrom (ToOrFromRuleCondition)
Function GetFromConditionDetails(ByRef objCondition As Outlook.ruleCondition) As String
    On Error GoTo ErrorHandler
    Dim objFromCond As Outlook.ToOrFromRuleCondition
    Dim objRecipient As Outlook.Recipient
    Dim strDetails As String
    
    If objCondition.conditionType = olConditionFrom Then
        Set objFromCond = objCondition
        If objFromCond.Recipients.Count > 0 Then
            For Each objRecipient In objFromCond.Recipients
                strDetails = strDetails & objRecipient.Address & "; "
            Next objRecipient
        Else
            strDetails = "(No sender specified)"
        End If
    End If
    
    ' Clean up trailing separator
    If Len(strDetails) > 2 Then
        GetFromConditionDetails = Left(strDetails, Len(strDetails) - 2)
    Else
        GetFromConditionDetails = strDetails
    End If
    Exit Function

ErrorHandler:
    GetFromConditionDetails = "Error extracting From: " & Err.Description
End Function

' Function to extract details for olConditionSenderAddress (AddressRuleCondition)
Function GetSenderAddressDetails(ByRef objCondition As Outlook.ruleCondition) As String
    On Error GoTo ErrorHandler
    Dim objAddressCond As Outlook.AddressRuleCondition
    Dim strDetails As String
    
    If objCondition.conditionType = olConditionSenderAddress Then
        Set objAddressCond = objCondition
        If Not IsEmpty(objAddressCond.Address) Then
            ' Address property returns an array of strings
            Dim vAddress As Variant
            For Each vAddress In objAddressCond.Address
                strDetails = strDetails & vAddress & "; "
            Next vAddress
        Else
            strDetails = "(No address specified)"
        End If
    End If
    
    ' Clean up trailing separator
    If Len(strDetails) > 2 Then
        GetSenderAddressDetails = Left(strDetails, Len(strDetails) - 2)
    Else
        GetSenderAddressDetails = strDetails
    End If
    Exit Function

ErrorHandler:
    GetSenderAddressDetails = "Error extracting SenderAddress: " & Err.Description
End Function

' Function to extract details for TextRuleConditions (Subject, BodyOrSubject, Body)
Function GetTextConditionDetails(ByRef objCondition As Outlook.ruleCondition) As String
    On Error GoTo ErrorHandler
    Dim objTextCond As Outlook.TextRuleCondition
    Dim strDetails As String
    
    ' Check for the three types that use TextRuleCondition
    If objCondition.conditionType = olConditionSubject Or _
       objCondition.conditionType = olConditionBodyOrSubject Or _
       objCondition.conditionType = olConditionBody Then
       
        Set objTextCond = objCondition
        
        If Not IsEmpty(objTextCond.Text) Then
            ' Text property returns an array of strings (the words/phrases)
            Dim vText As Variant
            For Each vText In objTextCond.Text
                strDetails = strDetails & "'" & vText & "'; "
            Next vText
        Else
            strDetails = "(No text specified)"
        End If
    End If
    
    ' Clean up trailing separator
    If Len(strDetails) > 2 Then
        GetTextConditionDetails = Left(strDetails, Len(strDetails) - 2)
    Else
        GetTextConditionDetails = strDetails
    End If
    Exit Function

ErrorHandler:
    GetTextConditionDetails = "Error extracting Text: " & Err.Description
End Function


' Function to extract details for olConditionSentTo (ToOrFromRuleCondition)
Function GetSentToConditionDetails(ByRef objCondition As Outlook.ruleCondition) As String
    On Error GoTo ErrorHandler
    Dim objSentToCond As Outlook.ToOrFromRuleCondition
    Dim objRecipient As Outlook.Recipient
    Dim strDetails As String
    
    If objCondition.conditionType = olConditionSentTo Then
        Set objSentToCond = objCondition
        If objSentToCond.Recipients.Count > 0 Then
            For Each objRecipient In objSentToCond.Recipients
                strDetails = strDetails & objRecipient.Address & "; "
            Next objRecipient
        Else
            strDetails = "(No recipient specified)"
        End If
    End If
    
    ' Clean up trailing separator
    If Len(strDetails) > 2 Then
        GetSentToConditionDetails = Left(strDetails, Len(strDetails) - 2)
    Else
        GetSentToConditionDetails = strDetails
    End If
    Exit Function

ErrorHandler:
    GetSentToConditionDetails = "Error extracting SentTo: " & Err.Description
End Function
' =================================================================================
' Function to map OlRuleConditionType/OlRuleActionType value to its string name
' =================================================================================
Function GetConditionTypeName(ByVal lngType As Long, ByVal isAction As Boolean) As String
    If isAction Then
        Select Case lngType
            Case 0: GetConditionTypeName = "olRuleActionUnknown"
            Case 1: GetConditionTypeName = "olRuleActionMoveToFolder"
            Case 2: GetConditionTypeName = "olRuleActionAssignToCategory"
            Case 3: GetConditionTypeName = "olRuleActionDelete"
            Case 4: GetConditionTypeName = "olRuleActionDeletePermanently"
            Case 5: GetConditionTypeName = "olRuleActionCopyToFolder" ' NEW ACTION
            Case 6: GetConditionTypeName = "olRuleActionForward"
            Case 7: GetConditionTypeName = "olRuleActionForwardAsAttachment"
            Case 8: GetConditionTypeName = "olRuleActionRedirect"
            Case 9: GetConditionTypeName = "olRuleActionServerReply"
            Case 10: GetConditionTypeName = "olRuleActionTemplate"
            Case 11: GetConditionTypeName = "olRuleActionFlagForActionInDays"
            Case 12: GetConditionTypeName = "olRuleActionFlagColor"
            Case 13: GetConditionTypeName = "olRuleActionFlagClear"
            Case 14: GetConditionTypeName = "olRuleActionImportance" ' NEW ACTION
            Case 15: GetConditionTypeName = "olRuleActionSensitivity"
            Case 16: GetConditionTypeName = "olRuleActionPrint"
            Case 17: GetConditionTypeName = "olRuleActionPlaySound"
            Case 18: GetConditionTypeName = "olRuleActionStartApplication"
            Case 19: GetConditionTypeName = "olRuleActionMarkRead"
            Case 20: GetConditionTypeName = "olRuleActionRunScript"
            Case 21: GetConditionTypeName = "olRuleActionStop"
            Case 22: GetConditionTypeName = "olRuleActionCustomAction"
            Case 23: GetConditionTypeName = "olRuleActionNewItemAlert"
            Case 24: GetConditionTypeName = "olRuleActionDesktopAlert" ' NEW ACTION
            Case 25: GetConditionTypeName = "olRuleActionNotifyRead"
            Case 26: GetConditionTypeName = "olRuleActionNotifyDelivery"
            Case 27: GetConditionTypeName = "olRuleActionCcMessage"
            Case 28: GetConditionTypeName = "olRuleActionDefer"
            Case 30: GetConditionTypeName = "olRuleActionClearCategories" ' NEW ACTION
            Case 41: GetConditionTypeName = "olRuleActionMarkAsTask"
            Case Else: GetConditionTypeName = "UnknownAction (" & lngType & ")"
        End Select
    Else
        Select Case lngType
            Case 0: GetConditionTypeName = "olConditionUnknown"
            Case 1: GetConditionTypeName = "olConditionFrom"
            Case 2: GetConditionTypeName = "olConditionSubject"
            Case 3: GetConditionTypeName = "olConditionAccount"
            Case 4: GetConditionTypeName = "olConditionOnlyToMe"
            Case 5: GetConditionTypeName = "olConditionTo"
            Case 6: GetConditionTypeName = "olConditionImportance"
            Case 7: GetConditionTypeName = "olConditionSensitivity"
            Case 8: GetConditionTypeName = "olConditionFlaggedForAction"
            Case 9: GetConditionTypeName = "olConditionCc"
            Case 10: GetConditionTypeName = "olConditionToOrCc"
            Case 11: GetConditionTypeName = "olConditionNotTo"
            Case 12: GetConditionTypeName = "olConditionSentTo"
            Case 13: GetConditionTypeName = "olConditionBody"
            Case 14: GetConditionTypeName = "olConditionBodyOrSubject"
            Case 15: GetConditionTypeName = "olConditionMessageHeader"
            Case 16: GetConditionTypeName = "olConditionRecipientAddress"
            Case 17: GetConditionTypeName = "olConditionSenderAddress"
            Case 18: GetConditionTypeName = "olConditionCategory"
            Case 19: GetConditionTypeName = "olConditionOOF"
            Case 20: GetConditionTypeName = "olConditionHasAttachment"
            Case 21: GetConditionTypeName = "olConditionSizeRange"
            Case 22: GetConditionTypeName = "olConditionDateRange"
            Case 23: GetConditionTypeName = "olConditionFormName"
            Case 24: GetConditionTypeName = "olConditionProperty"
            Case 25: GetConditionTypeName = "olConditionSenderInAddressBook"
            Case 26: GetConditionTypeName = "olConditionMeetingInviteOrUpdate"
            Case 27: GetConditionTypeName = "olConditionLocalMachineOnly"
            Case 28: GetConditionTypeName = "olConditionOtherMachine"
            Case 29: GetConditionTypeName = "olConditionAnyCategory"
            Case 30: GetConditionTypeName = "olConditionFromRssFeed"
            Case 31: GetConditionTypeName = "olConditionFromAnyRssFeed"
            Case Else: GetConditionTypeName = "UnknownCondition (" & lngType & ")"
        End Select
    End If
End Function
' =================================================================================
' Helper Function for Exporting to Excel
' =================================================================================

Function ExportToExcel(strData As String)
    On Error GoTo ErrorHandler
    
    Dim objExcel As Object
    Dim objWorkbook As Object
    Dim objWorksheet As Object
    Dim arrData() As String
    Dim arrRow() As String
    Dim i As Long, j As Long
    Dim lngLastRow As Long
    
    ' Create or get the Excel Application object
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    
    Set objWorkbook = objExcel.Workbooks.Add
    Set objWorksheet = objWorkbook.Sheets(1)
    
    ' Split the data into rows
    ' vbCrLf is the row separator
    arrData = Split(strData, vbCrLf)
    
    ' Loop through each row of data and write to cells
    For i = LBound(arrData) To UBound(arrData) - 1
        ' Split the row into columns (vbTab is the column separator)
        arrRow = Split(arrData(i), vbTab)
        
        ' Loop through each column of data
        For j = LBound(arrRow) To UBound(arrRow)
            ' Write the value to the corresponding cell (row i+1, column j+1)
            objWorksheet.Cells(i + 1, j + 1).Value = arrRow(j)
        Next j
    Next i
    
    ' Get the last row written
    lngLastRow = objWorksheet.UsedRange.Rows.Count
    
    ' === APPLY FORMATTING ===
  
    ' Column/Row Layout
    ' 1. Apply Wrap Text to all cells
    With objWorksheet.UsedRange
        .WrapText = True
    End With
        
    ' 2. Set Column Widths
    
    ' A. Set specific Condition Type columns to 80 (Columns F through N)
    ' This includes: Condition Type (Value), Condition Type (Name), and all detailed condition columns
    ' Column index mapping based on the main routine's output:
    ' F: Condition Type(s) (Value)
    ' G: Condition Type(s) (Name)
    ' H: From (Condition)
    ' I: Sender Address (Condition)
    ' J: Subject (Condition)
    ' K: Body/Subject (Condition)
    ' L: Body (Condition)
    ' M: Sent To (Condition)
    ' N: Any Category (Condition)
    ' O: Move to Folder (Action)
    
    Dim objRange As Object
    Set objRange = objWorksheet.Range("F:O")
    objRange.ColumnWidth = 80
    
    ' B. Set all other columns (A:E and O onwards) to AutoFit
    ' A:E are Rule Info
    objWorksheet.Range("A:E").Columns.AutoFit
    ' P onwards are Other Conditions, Actions, and Exceptions
    objWorksheet.Range("P:Z").Columns.AutoFit ' Use a range that covers the rest of the columns
    
    
    ' 3. After column adjustments, apply Auto Row Height to all used cells
    With objWorksheet.UsedRange
        .EntireRow.AutoFit
    End With
    
    ' Color Formatting
    ' 1. Format the Header Row
    With objWorksheet.Rows(1).Interior
        .Color = RGB(217, 217, 217) ' Light Grey
    End With
    objWorksheet.Rows(1).Font.Bold = True
    
    Exit Function

ErrorHandler:
    MsgBox "An error occurred during Excel export: " & Err.Description, vbCritical
    If Not objExcel Is Nothing Then
        If objExcel.Visible Then
            objExcel.Quit
        End If
    End If
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
End Function

' =================================================================================
' Main Subroutine
' =================================================================================

Sub ListRulesForSpecificMailbox()

    Dim objApp As Outlook.Application
    Dim objStore As Outlook.Store
    Dim objRules As Outlook.Rules
    Dim objRule As Outlook.Rule
    Dim objCondition As Outlook.ruleCondition
    Dim objAction As Outlook.ruleAction
    
    Dim strOutput As String
    Dim strMailboxName As String
    Dim fFound As Boolean
    Dim lngOutputChoice As Long

    ' Condition Variables
    Dim strRuleConditionTypes As String, strRuleConditionNames As String
    Dim strCondFrom As String, strCondSenderAddress As String, strCondSubject As String
    Dim strCondBodyOrSubject As String, strCondBody As String, strCondSentTo As String
    Dim strCondAnyCategory As String, strOtherConditions As String
    
    ' Action Variables
    Dim strMoveToFolder As String, strStopProcessing As String
    Dim strCopyToFolder As String ' NEW
    Dim strDesktopAlert As String ' NEW
    Dim strImportanceAction As String ' NEW
    Dim strClearCategories As String ' NEW
    Dim strOtherActions As String
    
    ' Exception Variables
    Dim strExceptionConditions As String
    Dim strExceptionValues As String ' NEW: Detailed exception values

    Set objApp = Outlook.Application

    ' === 1. Select/Specify Mailbox ===
    strMailboxName = InputBox("Enter the Display Name of the Mailbox:", "Select Mailbox")
    If strMailboxName = "" Then Exit Sub

    For Each objStore In objApp.Session.Stores
        If objStore.DisplayName = strMailboxName Then
            Set objRules = objStore.GetRules
            fFound = True
            Exit For
        End If
    Next objStore

    If Not fFound Then
        MsgBox "Mailbox '" & strMailboxName & "' not found.", vbCritical
        Exit Sub
    End If

    ' === 2. Build Header and Rules Data ===
    
    ' Initialize header with ALL NEW columns
    strOutput = "Mailbox" & vbTab & "Rule Index" & vbTab & "Rule Name" & vbTab & "Enabled" & vbTab & _
                "Condition Type(s) (Value)" & vbTab & "Condition Type(s) (Name)" & vbTab & _
                "From (Condition)" & vbTab & "Sender Address (Condition)" & vbTab & _
                "Subject (Condition)" & vbTab & "Body/Subject (Condition)" & vbTab & "Body (Condition)" & vbTab & _
                "Sent To (Condition)" & vbTab & "Any Category (Condition)" & vbTab & "Other Conditions" & vbTab & _
                "Action: Move To Folder" & vbTab & "Action: Copy To Folder" & vbTab & "Action: Stop Processing" & vbTab & _
                "Action: Desktop Alert" & vbTab & "Action: Set Importance" & vbTab & "Action: Clear Categories" & vbTab & _
                "Other Actions" & vbTab & _
                "Exception Types" & vbTab & "Exception Details" & vbCrLf

    If Not objRules Is Nothing Then
        For Each objRule In objRules
            ' Reset ALL variables
            strCondFrom = "": strCondSenderAddress = "": strCondSubject = ""
            strCondBodyOrSubject = "": strCondBody = "": strCondSentTo = ""
            strCondAnyCategory = "": strOtherConditions = "": strOtherActions = ""
            strRuleConditionTypes = "": strRuleConditionNames = ""
            strMoveToFolder = "No": strStopProcessing = "No"
            strCopyToFolder = "No": strDesktopAlert = "No": strImportanceAction = "No": strClearCategories = "No" ' NEW ACTIONS
            strExceptionConditions = "": strExceptionValues = "" ' NEW EXCEPTIONS

            ' === PROCESS CONDITIONS ===
            For Each objCondition In objRule.Conditions
                If objCondition.Enabled Then
                    strRuleConditionTypes = strRuleConditionTypes & objCondition.conditionType & "; "
                    strRuleConditionNames = strRuleConditionNames & GetConditionTypeName(objCondition.conditionType, False) & "; "

                    Select Case objCondition.conditionType
                        Case olConditionFrom: strCondFrom = GetFromConditionDetails(objCondition)
                        Case olConditionSenderAddress: strCondSenderAddress = GetSenderAddressDetails(objCondition)
                        Case olConditionSubject: strCondSubject = GetTextConditionDetails(objCondition)
                        Case olConditionBodyOrSubject: strCondBodyOrSubject = GetTextConditionDetails(objCondition)
                        Case olConditionBody: strCondBody = GetTextConditionDetails(objCondition)
                        Case olConditionSentTo: strCondSentTo = GetSentToConditionDetails(objCondition)
                        Case olConditionAnyCategory: strCondAnyCategory = "Yes"
                        Case Else: strOtherConditions = strOtherConditions & GetConditionTypeName(objCondition.conditionType, False) & "; "
                    End Select
                End If
            Next objCondition

            ' === PROCESS ACTIONS ===
            For Each objAction In objRule.Actions
                If objAction.Enabled Then
                    Select Case objAction.actionType
                        Case olRuleActionMoveToFolder
                            strMoveToFolder = "Yes (" & objAction.Folder.FolderPath & ")"
                        Case olRuleActionCopyToFolder
                            strCopyToFolder = "Yes (" & objAction.Folder.FolderPath & ")"
                        Case olRuleActionStop
                            strStopProcessing = "Yes"
                        Case olRuleActionDesktopAlert
                            strDesktopAlert = "Yes"
                        Case olRuleActionImportance
                            strImportanceAction = "Yes"
                        Case olRuleActionClearCategories
                            strClearCategories = "Yes"
                        Case Else
                            strOtherActions = strOtherActions & GetConditionTypeName(objAction.actionType, True) & "; "
                    End Select
                End If
            Next objAction
            
            ' === PROCESS EXCEPTIONS (WITH DETAILS) ===
            For Each objCondition In objRule.Exceptions
                If objCondition.Enabled Then
                    strExceptionConditions = strExceptionConditions & GetConditionTypeName(objCondition.conditionType, False) & "; "
                    
                    ' Extracting exception details (same logic as conditions)
                    Select Case objCondition.conditionType
                        Case olConditionFrom: strExceptionValues = strExceptionValues & "From: " & GetFromConditionDetails(objCondition) & " | "
                        Case olConditionSenderAddress: strExceptionValues = strExceptionValues & "Addr: " & GetSenderAddressDetails(objCondition) & " | "
                        Case olConditionSubject: strExceptionValues = strExceptionValues & "Subj: " & GetTextConditionDetails(objCondition) & " | "
                        Case olConditionBodyOrSubject: strExceptionValues = strExceptionValues & "B/S: " & GetTextConditionDetails(objCondition) & " | "
                        Case olConditionBody: strExceptionValues = strExceptionValues & "Body: " & GetTextConditionDetails(objCondition) & " | "
                        Case olConditionSentTo: strExceptionValues = strExceptionValues & "To: " & GetSentToConditionDetails(objCondition) & " | "
                        Case olConditionAnyCategory: strExceptionValues = strExceptionValues & "AnyCat: Yes | "
                        ' You would add more cases here for other conditions (e.g., olConditionSizeRange) if needed
                        Case Else: strExceptionValues = strExceptionValues & "Other: " & GetConditionTypeName(objCondition.conditionType, False) & " | "
                    End Select
                End If
            Next objCondition

            ' Clean up trailing separator
            If Len(strRuleConditionTypes) > 2 Then strRuleConditionTypes = Left(strRuleConditionTypes, Len(strRuleConditionTypes) - 2)
            If Len(strRuleConditionNames) > 2 Then strRuleConditionNames = Left(strRuleConditionNames, Len(strRuleConditionNames) - 2)
            If Len(strOtherConditions) > 2 Then strOtherConditions = Left(strOtherConditions, Len(strOtherConditions) - 2)
            If Len(strOtherActions) > 2 Then strOtherActions = Left(strOtherActions, Len(strOtherActions) - 2)
            If Len(strExceptionConditions) > 2 Then strExceptionConditions = Left(strExceptionConditions, Len(strExceptionConditions) - 2)
            If Len(strExceptionValues) > 2 Then strExceptionValues = Left(strExceptionValues, Len(strExceptionValues) - 3) ' Extra | space

            ' Append rule details to the output string
            strOutput = strOutput & objStore.DisplayName & vbTab & _
                        objRule.ExecutionOrder & vbTab & _
                        objRule.Name & vbTab & _
                        objRule.Enabled & vbTab & _
                        strRuleConditionTypes & vbTab & strRuleConditionNames & vbTab & _
                        strCondFrom & vbTab & strCondSenderAddress & vbTab & _
                        strCondSubject & vbTab & strCondBodyOrSubject & vbTab & strCondBody & vbTab & _
                        strCondSentTo & vbTab & strCondAnyCategory & vbTab & strOtherConditions & vbTab & _
                        strMoveToFolder & vbTab & strCopyToFolder & vbTab & strStopProcessing & vbTab & _
                        strDesktopAlert & vbTab & strImportanceAction & vbTab & strClearCategories & vbTab & _
                        strOtherActions & vbTab & _
                        strExceptionConditions & vbTab & strExceptionValues & vbCrLf
        Next objRule

        ' === 3. Determine Output Method and Print/Export ===
        lngOutputChoice = MsgBox("Do you want to export the rules to an Excel spreadsheet? Click 'No' to print to the Immediate Window.", vbYesNoCancel + vbQuestion, "Output Selection")

        If lngOutputChoice = vbYes Then
            Call ExportToExcel(strOutput)
            MsgBox "Rule details exported to Excel successfully.", vbInformation
        ElseIf lngOutputChoice = vbNo Then
            Debug.Print strOutput
            MsgBox "Rule details have been printed to the Immediate Window in the VBA Editor (Press Ctrl+G).", vbInformation
        ElseIf lngOutputChoice = vbCancel Then
            MsgBox "Operation cancelled by user.", vbExclamation
        End If

    Else
        MsgBox "No rules found for mailbox: " & strMailboxName
    End If

    ' Clean up
    Set objRules = Nothing: Set objStore = Nothing: Set objApp = Nothing
End Sub

