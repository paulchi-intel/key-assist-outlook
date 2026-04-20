Option Explicit

' ===== ExpertGPT Settings =====
Private Const APP_NAME As String = "ExpertGPTEmailAssistant"
Private Const SETTINGS_SECTION As String = "Settings"
Private Const OPENAI_BASE_URL As String = "https://expertgpt.intel.com/v1"
Private Const ANTHROPIC_BASE_URL As String = "https://expertgpt.intel.com/anthropic/v1"
Private Const GNAI_OPENAI_BASE_URL As String = "https://gnai.intel.com/api/providers/openai/v1"
Private Const GNAI_ANTHROPIC_BASE_URL As String = "https://gnai.intel.com/api/providers/anthropic"
Private Const ANTHROPIC_API_VERSION As String = "2023-06-01"
Private Const REQUEST_TIMEOUT_MS As Long = 20000
Private Const ANTHROPIC_REQUEST_TIMEOUT_MS As Long = 120000
Private Const DEFAULT_MODEL As String = "gpt-4.1-mini"
Private Const FAQ_FOLDER_PATH As String = "My Folders\FAQ"
Private Const FAQ_MAX_ITEMS As Long = 100
Private Const FAQ_CONTENT_CHARS As Long = 2000
Private Const DAILY_EMAIL_BODY_CHARS As Long = 500
Private Const DAILY_EMAIL_MAX_ITEMS As Long = 300

' ------------------------------------------------------------
' Quick Actions
' ------------------------------------------------------------
Sub ExpertGPT_AI_Summarize()
    ProcessEmailWithPromptExpertGPT "Summarize this email in bullet points using Traditional Chinese, focusing on key concepts.", "AI Processed (Summarize): "
End Sub

Sub ExpertGPT_AI_Translate()
    ProcessEmailWithPromptExpertGPT "Translate this email to Traditional Chinese. Keep it clear and professional.", "AI Processed (Translate): "
End Sub

Sub ExpertGPT_AI_ActionItems()
    ProcessEmailWithPromptExpertGPT "Extract all action items from this email. List them in bullet points with who should do what and by when (if specified). Output with Traditional Chinese.", "AI Processed (AR): "
End Sub

Sub ExpertGPT_AI_Reply()
    ProcessEmailWithPromptExpertGPT "Write a professional and concise reply to this email. Keep the tone friendly but businesslike.", "AI Processed (Reply): "
End Sub

Sub ExpertGPT_AI_FAQ()
    ProcessEmailWithPromptExpertGPT "Convert this email into a FAQ (Frequently Asked Questions) format. Extract key questions and provide clear answers. Use Traditional Chinese for output. Format: Q: [question] A: [answer]", "AI Processed (FAQ): "
End Sub

Sub ExpertGPT_AI_Custom()
    On Error GoTo ErrorHandler

    Dim userPrompt As String

    userPrompt = InputBox( _
        "Enter your instruction for AI:" & vbCrLf & vbCrLf & _
        "Examples:" & vbCrLf & _
        "- Summarize this email" & vbCrLf & _
        "- Translate to English" & vbCrLf & _
        "- Extract action items" & vbCrLf & _
        "- Write a reply", _
        "ExpertGPT Email Assistant", _
        "Summarize this email in bullet points")

    If Trim$(userPrompt) = "" Then
        Exit Sub
    End If

    ProcessEmailWithPromptExpertGPT userPrompt
    Exit Sub

ErrorHandler:
    MsgBox "Error processing email: " & Err.Description, vbCritical, "Error"
End Sub

' ------------------------------------------------------------
' FAQ / Daily summary features
' ------------------------------------------------------------
Sub ExpertGPT_FAQ_Ask()
    On Error GoTo ErrorHandler

    Dim userQuestion As String
    Dim apiKey As String
    Dim selectedModel As String
    Dim faqContent As String
    Dim faqCount As Long
    Dim aiPrompt As String
    Dim aiResponse As String
    Dim resultMail As Outlook.MailItem

    userQuestion = InputBox( _
        "Enter your question:" & vbCrLf & vbCrLf & _
        "I will search for answers from the FAQ database." & vbCrLf & vbCrLf & _
        "Example questions:" & vbCrLf & _
        "- Can graphics driver enable or disable LOBF feature?" & vbCrLf & _
        "- Where can I download DAR technical specifications?", _
        "ExpertGPT FAQ Assistant", _
        "")

    If Trim$(userQuestion) = "" Then Exit Sub

    apiKey = EnsureApiKeyReady()
    If Len(apiKey) = 0 Then Exit Sub

    selectedModel = EnsureModelReady(apiKey)
    If Len(selectedModel) = 0 Then Exit Sub

    faqContent = LoadFAQFromFolder(faqCount)
    If faqCount = 0 Then
        MsgBox "No emails found in FAQ folder." & vbCrLf & _
               "Please verify folder path: " & FAQ_FOLDER_PATH, vbExclamation, "FAQ Not Found"
        Exit Sub
    End If

    aiPrompt = "You are a helpful FAQ assistant. Based on the provided FAQ knowledge base, answer the user's question in Traditional Chinese. " & _
               "If the answer is found in the FAQs, provide a clear and detailed response. " & _
               "If the answer is not in the FAQs, politely say you don't have that information in the knowledge base. " & _
               "Always cite which FAQ email(s) you're referencing (by number or subject)." & vbCrLf & vbCrLf & _
               "User Question:" & vbCrLf & userQuestion

    aiResponse = ProcessWithExpertGPT(faqContent, aiPrompt, apiKey, selectedModel)

    Set resultMail = Application.CreateItem(olMailItem)
    With resultMail
        .To = GetCurrentUserSmtpAddress()
        .Subject = "FAQ Answer: " & Left$(userQuestion, 50) & IIf(Len(userQuestion) > 50, "...", "")
        .Body = "=== FAQ Smart Assistant ===" & vbCrLf & _
                String$(60, "=") & vbCrLf & vbCrLf & _
                "[Your Question]" & vbCrLf & _
                userQuestion & vbCrLf & vbCrLf & _
                String$(60, "-") & vbCrLf & vbCrLf & _
                "[AI Answer]" & vbCrLf & _
                aiResponse & vbCrLf & vbCrLf & _
                String$(60, "=") & vbCrLf & _
                "Source: " & faqCount & " FAQ emails" & vbCrLf & _
                "Folder: " & FAQ_FOLDER_PATH & vbCrLf & _
                "Generated by ExpertGPT"
        .Display
    End With

    Set resultMail = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error processing question: " & Err.Description, vbCritical, "Error"
End Sub

Sub ExpertGPT_SummarizeTodayEmails()
    SummarizeEmailsForDateExpertGPT Date, "Today"
End Sub

Sub ExpertGPT_SummarizeYesterdayEmails()
    SummarizeEmailsForDateExpertGPT Date - 1, "Yesterday"
End Sub

Sub ExpertGPT_SummarizeCustomDateEmails()
    On Error GoTo ErrorHandler

    Dim userInput As String
    Dim selectedDate As Date
    Dim dateLabel As String
    Dim daysAgo As Integer

    userInput = InputBox( _
        "Enter the date to summarize emails:" & vbCrLf & vbCrLf & _
        "Format: YYYY-MM-DD" & vbCrLf & _
        "Example: 2026-01-05" & vbCrLf & vbCrLf & _
        "Or use relative dates:" & vbCrLf & _
        "- Today" & vbCrLf & _
        "- Yesterday" & vbCrLf & _
        "- -2 (2 days ago)" & vbCrLf & _
        "- -7 (7 days ago)", _
        "Select Date for Email Summary", _
        Format$(Date, "yyyy-mm-dd"))

    If Trim$(userInput) = "" Then Exit Sub

    Select Case LCase$(Trim$(userInput))
        Case "today"
            selectedDate = Date
            dateLabel = "Today"
        Case "yesterday"
            selectedDate = Date - 1
            dateLabel = "Yesterday"
        Case Else
            If Left$(Trim$(userInput), 1) = "-" And IsNumeric(Mid$(Trim$(userInput), 2)) Then
                daysAgo = Abs(CInt(Trim$(userInput)))
                selectedDate = Date - daysAgo
                dateLabel = CStr(daysAgo) & " days ago"
            ElseIf IsDate(userInput) Then
                selectedDate = CDate(userInput)
                dateLabel = Format$(selectedDate, "yyyy-mm-dd")
            Else
                MsgBox "Invalid date format. Please use YYYY-MM-DD format.", vbExclamation, "Invalid Date"
                Exit Sub
            End If
    End Select

    SummarizeEmailsForDateExpertGPT selectedDate, dateLabel
    Exit Sub

ErrorHandler:
    MsgBox "Error processing date: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub SummarizeEmailsForDateExpertGPT(ByVal targetDate As Date, ByVal dateLabel As String)
    On Error GoTo ErrorHandler

    Dim apiKey As String
    Dim selectedModel As String
    Dim objNamespace As Outlook.NameSpace
    Dim objInbox As Outlook.MAPIFolder
    Dim objItem As Object
    Dim objMail As Outlook.MailItem
    Dim dateStart As Date
    Dim dateEnd As Date
    Dim emailList As String
    Dim emailCount As Long
    Dim aiPrompt As String
    Dim aiResponse As String
    Dim resultMail As Outlook.MailItem

    apiKey = EnsureApiKeyReady()
    If Len(apiKey) = 0 Then Exit Sub

    selectedModel = EnsureModelReady(apiKey)
    If Len(selectedModel) = 0 Then Exit Sub

    dateStart = targetDate
    dateEnd = targetDate + 1

    Set objNamespace = Application.GetNamespace("MAPI")
    Set objInbox = objNamespace.GetDefaultFolder(olFolderInbox)

    emailList = ""
    emailCount = 0

    For Each objItem In objInbox.Items
        If TypeName(objItem) = "MailItem" Then
            Set objMail = objItem

            If objMail.ReceivedTime >= dateStart And objMail.ReceivedTime < dateEnd Then
                emailCount = emailCount + 1
                If emailCount > DAILY_EMAIL_MAX_ITEMS Then Exit For

                emailList = emailList & _
                    "=== Email #" & emailCount & " ===" & vbCrLf & _
                    "From: " & objMail.SenderName & " <" & GetSenderEmail(objMail) & ">" & vbCrLf & _
                    "Subject: " & objMail.Subject & vbCrLf & _
                    "Time: " & Format$(objMail.ReceivedTime, "hh:nn AM/PM") & vbCrLf & _
                    "Content: " & Left$(NzStr(objMail.Body), DAILY_EMAIL_BODY_CHARS) & vbCrLf & vbCrLf
            End If
        End If
    Next objItem

    If emailCount = 0 Then
        MsgBox "No emails received on " & Format$(targetDate, "yyyy-mm-dd") & ".", vbInformation, "No Emails"
        GoTo CleanExit
    End If

    aiPrompt = "Please provide a comprehensive summary of " & dateLabel & "'s emails in Traditional Chinese. " & _
               "Follow these requirements:" & vbCrLf & _
               "1. Classify and organize emails by category" & vbCrLf & _
               "2. Put HP-related information at the TOP (highest priority)" & vbCrLf & _
               "3. Sort all content by urgency level (urgent/important/normal)" & vbCrLf & _
               "4. At the END, create a section called 'My Action Items' listing all action items that require my attention" & vbCrLf & _
               "5. Use bullet points for clarity" & vbCrLf & _
               "6. Highlight important deadlines" & vbCrLf & vbCrLf & _
               "Total emails: " & emailCount

    aiResponse = ProcessWithExpertGPT(emailList, aiPrompt, apiKey, selectedModel)

    Set resultMail = Application.CreateItem(olMailItem)
    With resultMail
        .To = GetCurrentUserSmtpAddress()
        .Subject = dateLabel & "'s Email Summary (" & Format$(targetDate, "yyyy-mm-dd") & ") - " & emailCount & " emails"
        .Body = "=== " & UCase$(dateLabel) & "'S EMAIL SUMMARY ===" & vbCrLf & _
                "Date: " & Format$(targetDate, "yyyy-mm-dd (dddd)") & vbCrLf & _
                "Total Emails: " & emailCount & vbCrLf & _
                String$(60, "=") & vbCrLf & vbCrLf & _
                aiResponse & vbCrLf & vbCrLf & _
                String$(60, "=") & vbCrLf & _
                "Generated by ExpertGPT"
        .Display
    End With

CleanExit:
    Set resultMail = Nothing
    Set objMail = Nothing
    Set objInbox = Nothing
    Set objNamespace = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error summarizing emails: " & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

Private Function LoadFAQFromFolder(ByRef faqCount As Long) As String
    On Error GoTo ErrorHandler

    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim objItem As Object
    Dim objMail As Outlook.MailItem
    Dim faqContent As String

    Set objNamespace = Application.GetNamespace("MAPI")
    Set objFolder = GetFolderByPath(objNamespace, FAQ_FOLDER_PATH)

    If objFolder Is Nothing Then
        LoadFAQFromFolder = ""
        faqCount = 0
        Exit Function
    End If

    faqContent = ""
    faqCount = 0

    For Each objItem In objFolder.Items
        If TypeName(objItem) = "MailItem" Then
            Set objMail = objItem
            faqCount = faqCount + 1
            If faqCount > FAQ_MAX_ITEMS Then Exit For

            faqContent = faqContent & _
                "=== FAQ #" & faqCount & " ===" & vbCrLf & _
                "Subject: " & objMail.Subject & vbCrLf & _
                "From: " & objMail.SenderName & vbCrLf & _
                "Date: " & Format$(objMail.ReceivedTime, "yyyy-mm-dd") & vbCrLf & _
                "Content:" & vbCrLf & _
                Left$(NzStr(objMail.Body), FAQ_CONTENT_CHARS) & vbCrLf & _
                String$(60, "-") & vbCrLf & vbCrLf
        End If
    Next objItem

    LoadFAQFromFolder = faqContent
    Set objMail = Nothing
    Set objFolder = Nothing
    Set objNamespace = Nothing
    Exit Function

ErrorHandler:
    LoadFAQFromFolder = ""
    faqCount = 0
End Function

Private Function GetFolderByPath(ByVal objNamespace As Outlook.NameSpace, ByVal folderPath As String) As Outlook.MAPIFolder
    On Error GoTo ErrorHandler

    Dim folders() As String
    Dim currentFolder As Outlook.MAPIFolder
    Dim i As Long

    folders = Split(folderPath, "\")
    If UBound(folders) < 0 Then
        Set GetFolderByPath = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set currentFolder = objNamespace.Folders(folders(0))
    If currentFolder Is Nothing Then
        Set currentFolder = objNamespace.GetDefaultFolder(olFolderInbox).Parent.Folders(folders(0))
    End If
    On Error GoTo ErrorHandler

    If currentFolder Is Nothing Then
        Set GetFolderByPath = Nothing
        Exit Function
    End If

    For i = 1 To UBound(folders)
        Set currentFolder = currentFolder.Folders(folders(i))
        If currentFolder Is Nothing Then
            Set GetFolderByPath = Nothing
            Exit Function
        End If
    Next i

    Set GetFolderByPath = currentFolder
    Exit Function

ErrorHandler:
    Set GetFolderByPath = Nothing
End Function

Private Function GetSenderEmail(ByVal mail As Outlook.MailItem) As String
    On Error Resume Next

    Dim sender As Outlook.AddressEntry
    Dim exchUser As Outlook.ExchangeUser

    Set sender = mail.Sender
    If Not sender Is Nothing Then
        If sender.AddressEntryUserType = olExchangeUserAddressEntry Or _
           sender.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
            Set exchUser = sender.GetExchangeUser
            If Not exchUser Is Nothing Then
                GetSenderEmail = Trim$(exchUser.PrimarySmtpAddress)
            End If
        Else
            GetSenderEmail = Trim$(sender.Address)
        End If
    End If

    If Len(GetSenderEmail) = 0 Then
        GetSenderEmail = Trim$(mail.SenderEmailAddress)
    End If
End Function

Private Function GetCurrentUserSmtpAddress() As String
    On Error Resume Next

    Dim addr As String
    Dim ae As Object
    Dim exUser As Object

    If Application.Session.Accounts.Count > 0 Then
        addr = Trim$(Application.Session.Accounts.Item(1).SmtpAddress)
        If Len(addr) > 0 Then
            GetCurrentUserSmtpAddress = addr
            Exit Function
        End If
    End If

    Set ae = Application.Session.CurrentUser.AddressEntry
    If Not ae Is Nothing Then
        Set exUser = ae.GetExchangeUser
        If Not exUser Is Nothing Then
            addr = Trim$(exUser.PrimarySmtpAddress)
            If Len(addr) > 0 Then
                GetCurrentUserSmtpAddress = addr
                Exit Function
            End If
        End If
    End If

    GetCurrentUserSmtpAddress = Trim$(Application.Session.CurrentUser.Address)
End Function

' ------------------------------------------------------------
' Settings / model management
' ------------------------------------------------------------
Sub ExpertGPT_Configure()
    On Error GoTo ErrorHandler

    Dim apiKey As String
    Dim selectedModel As String

    apiKey = GetStoredApiKey()
    selectedModel = GetStoredModel()
    If Len(selectedModel) = 0 Then
        selectedModel = DEFAULT_MODEL
    End If

    If ShowConfigurationForm(apiKey, selectedModel) Then
        MsgBox "Settings saved. Model selected: " & selectedModel, vbInformation, "Key Assist"
    Else
        MsgBox "Configuration was cancelled. Model selected: " & selectedModel, vbInformation, "Key Assist"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error configuring ExpertGPT: " & Err.Description, vbCritical, "Error"
End Sub

Private Function ShowConfigurationForm(ByRef apiKey As String, ByRef selectedModel As String) As Boolean
    On Error GoTo Fail

    Dim outputText As String
    Dim parts() As String
    Dim inputKey As String
    Dim inputModel As String

    outputText = RunConfigurationFormByPowerShell(apiKey, selectedModel)
    If Len(outputText) = 0 Then
        Exit Function
    End If

    parts = Split(outputText, vbTab)
    If UBound(parts) < 1 Then
        Exit Function
    End If

    inputKey = NormalizeApiKey(parts(0))
    inputModel = Trim$(parts(1))

    If Not IsValidApiKey(inputKey) Then
        MsgBox "Please enter a valid API key (ExpertGPT pak_ or GNAI).", vbExclamation, "ExpertGPT"
        Exit Function
    End If

    If Len(inputModel) = 0 Then
        MsgBox "Please select a model.", vbExclamation, "ExpertGPT"
        Exit Function
    End If

    SaveSetting APP_NAME, SETTINGS_SECTION, "ApiKey", inputKey
    SaveSelectedModel inputModel

    apiKey = inputKey
    selectedModel = inputModel
    ShowConfigurationForm = True
    Exit Function

Fail:
    ShowConfigurationForm = False
End Function

Private Function RunConfigurationFormByPowerShell(ByVal defaultApiKey As String, ByVal defaultModel As String) As String
    On Error GoTo CleanFail

    Dim fso As Object
    Dim shellObj As Object
    Dim tempDir As String
    Dim token As String
    Dim scriptPath As String
    Dim outputPath As String
    Dim commandText As String
    Dim scriptText As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shellObj = CreateObject("WScript.Shell")

    tempDir = Environ$("TEMP")
    token = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int((Rnd() * 90000) + 10000))
    scriptPath = tempDir & "\ExpertGPT_ConfigForm_" & token & ".ps1"
    outputPath = tempDir & "\ExpertGPT_ConfigForm_" & token & ".txt"

    scriptText = BuildConfigurationFormScript(defaultApiKey, defaultModel, outputPath)
    WriteAllTextFile scriptPath, scriptText

    commandText = "powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File """ & scriptPath & """"
    shellObj.Run commandText, 0, True

    If fso.FileExists(outputPath) Then
        RunConfigurationFormByPowerShell = Trim$(ReadAllTextFile(outputPath))
    End If

CleanExit:
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FileExists(scriptPath) Then fso.DeleteFile scriptPath, True
        If fso.FileExists(outputPath) Then fso.DeleteFile outputPath, True
    End If
    Exit Function

CleanFail:
    RunConfigurationFormByPowerShell = ""
    Resume CleanExit
End Function

Private Function BuildConfigurationFormScript(ByVal defaultApiKey As String, ByVal defaultModel As String, ByVal outputPath As String) As String
    Dim lines As String
    Dim defaultKeyEscaped As String
    Dim defaultModelEscaped As String
    Dim outputEscaped As String
    Dim openaiEscaped As String
    Dim anthropicEscaped As String
    Dim gnaiOpenaiEscaped As String
    Dim gnaiAnthropicEscaped As String
    Dim defaultModelIdEscaped As String

    defaultKeyEscaped = EscapeForPowerShellSingleQuoted(defaultApiKey)
    defaultModelEscaped = EscapeForPowerShellSingleQuoted(defaultModel)
    outputEscaped = EscapeForPowerShellSingleQuoted(outputPath)
    openaiEscaped = EscapeForPowerShellSingleQuoted(OPENAI_BASE_URL)
    anthropicEscaped = EscapeForPowerShellSingleQuoted(ANTHROPIC_BASE_URL)
    gnaiOpenaiEscaped = EscapeForPowerShellSingleQuoted(GNAI_OPENAI_BASE_URL)
    gnaiAnthropicEscaped = EscapeForPowerShellSingleQuoted(GNAI_ANTHROPIC_BASE_URL)
    defaultModelIdEscaped = EscapeForPowerShellSingleQuoted(DEFAULT_MODEL)

    lines = "$ErrorActionPreference = 'Stop'" & vbCrLf
    lines = lines & "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf
    lines = lines & "Add-Type -AssemblyName System.Drawing" & vbCrLf
    lines = lines & "$openaiBase = '" & openaiEscaped & "'" & vbCrLf
    lines = lines & "$anthropicBase = '" & anthropicEscaped & "'" & vbCrLf
    lines = lines & "$gnaiOpenaiBase = '" & gnaiOpenaiEscaped & "'" & vbCrLf
    lines = lines & "$gnaiAnthropicBase = '" & gnaiAnthropicEscaped & "'" & vbCrLf
    lines = lines & "$defaultModelId = '" & defaultModelIdEscaped & "'" & vbCrLf
    lines = lines & "$form = New-Object System.Windows.Forms.Form" & vbCrLf
    lines = lines & "$form.Text = 'ExpertGPT Configuration'" & vbCrLf
    lines = lines & "$form.StartPosition = 'CenterScreen'" & vbCrLf
    lines = lines & "$form.Size = New-Object System.Drawing.Size(760, 250)" & vbCrLf
    lines = lines & "$form.MinimizeBox = $false" & vbCrLf
    lines = lines & "$form.MaximizeBox = $false" & vbCrLf
    lines = lines & "$form.TopMost = $true" & vbCrLf
    lines = lines & "$labelKey = New-Object System.Windows.Forms.Label" & vbCrLf
    lines = lines & "$labelKey.Text = 'API Key:  (ExpertGPT pak_ or GNAI)'" & vbCrLf
    lines = lines & "$labelKey.Location = New-Object System.Drawing.Point(20, 20)" & vbCrLf
    lines = lines & "$labelKey.AutoSize = $true" & vbCrLf
    lines = lines & "$txtKey = New-Object System.Windows.Forms.TextBox" & vbCrLf
    lines = lines & "$txtKey.Location = New-Object System.Drawing.Point(20, 44)" & vbCrLf
    lines = lines & "$txtKey.Size = New-Object System.Drawing.Size(710, 24)" & vbCrLf
    lines = lines & "$txtKey.Text = '" & defaultKeyEscaped & "'" & vbCrLf
    lines = lines & "$labelModel = New-Object System.Windows.Forms.Label" & vbCrLf
    lines = lines & "$labelModel.Text = 'Model:'" & vbCrLf
    lines = lines & "$labelModel.Location = New-Object System.Drawing.Point(20, 82)" & vbCrLf
    lines = lines & "$labelModel.AutoSize = $true" & vbCrLf
    lines = lines & "$combo = New-Object System.Windows.Forms.ComboBox" & vbCrLf
    lines = lines & "$combo.Location = New-Object System.Drawing.Point(20, 106)" & vbCrLf
    lines = lines & "$combo.Size = New-Object System.Drawing.Size(560, 26)" & vbCrLf
    lines = lines & "$combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList" & vbCrLf
    lines = lines & "$combo.MaxDropDownItems = 20" & vbCrLf
    lines = lines & "$btnLoad = New-Object System.Windows.Forms.Button" & vbCrLf
    lines = lines & "$btnLoad.Text = 'Load Models'" & vbCrLf
    lines = lines & "$btnLoad.Location = New-Object System.Drawing.Point(590, 104)" & vbCrLf
    lines = lines & "$btnLoad.Size = New-Object System.Drawing.Size(140, 30)" & vbCrLf
    lines = lines & "$okBtn = New-Object System.Windows.Forms.Button" & vbCrLf
    lines = lines & "$okBtn.Text = 'OK'" & vbCrLf
    lines = lines & "$okBtn.Location = New-Object System.Drawing.Point(560, 156)" & vbCrLf
    lines = lines & "$okBtn.Size = New-Object System.Drawing.Size(80, 30)" & vbCrLf
    lines = lines & "$cancelBtn = New-Object System.Windows.Forms.Button" & vbCrLf
    lines = lines & "$cancelBtn.Text = 'Cancel'" & vbCrLf
    lines = lines & "$cancelBtn.Location = New-Object System.Drawing.Point(650, 156)" & vbCrLf
    lines = lines & "$cancelBtn.Size = New-Object System.Drawing.Size(80, 30)" & vbCrLf
    lines = lines & "$modelMap = @{}" & vbCrLf
    lines = lines & "$skipIndexChange = $false" & vbCrLf
    lines = lines & "$initialModel = '" & defaultModelEscaped & "'" & vbCrLf
    lines = lines & "if (-not [string]::IsNullOrWhiteSpace($initialModel)) { $initialModel = $initialModel.Trim() }" & vbCrLf
    lines = lines & "function Fill-ModelList([string]$apiKey) {" & vbCrLf
    lines = lines & "  $combo.Items.Clear(); $modelMap.Clear()" & vbCrLf
    lines = lines & "  if ([string]::IsNullOrWhiteSpace($apiKey)) {" & vbCrLf
    lines = lines & "    $modelIdToShow = $defaultModelId" & vbCrLf
    lines = lines & "    if (-not [string]::IsNullOrWhiteSpace($initialModel)) { $modelIdToShow = $initialModel }" & vbCrLf
    lines = lines & "    $display = ""$modelIdToShow (0/0)""" & vbCrLf
    lines = lines & "    $null = $combo.Items.Add($display); $modelMap[$display] = $modelIdToShow; $combo.SelectedIndex = 0; $labelModel.Text = ""Model:  ($($modelMap.Count) models loaded)""; return" & vbCrLf
    lines = lines & "  }" & vbCrLf
    lines = lines & "  $isGnai = -not $apiKey.StartsWith('pak_')" & vbCrLf
    lines = lines & "  $openaiUrl = if ($isGnai) { $gnaiOpenaiBase } else { $openaiBase }" & vbCrLf
    lines = lines & "  $anthropicModelsUrl = if ($isGnai) { $gnaiAnthropicBase + '/v1/models' } else { $anthropicBase + '/models' }" & vbCrLf
    lines = lines & "  $headers = @{ Authorization = ""Bearer $apiKey""; 'Content-Type' = 'application/json' }" & vbCrLf
    lines = lines & "  $quotaMap = @{}" & vbCrLf
    lines = lines & "  if (-not $isGnai) {" & vbCrLf
    lines = lines & "    try {" & vbCrLf
    lines = lines & "      $q = Invoke-RestMethod -Method Get -Uri ($openaiUrl + '/quota') -Headers $headers -TimeoutSec 20" & vbCrLf
    lines = lines & "      if ($q -and $q.model_quotas) { $q.model_quotas.PSObject.Properties | ForEach-Object { $quotaMap[$_.Name] = $_.Value } }" & vbCrLf
    lines = lines & "    } catch {}" & vbCrLf
    lines = lines & "  }" & vbCrLf
    lines = lines & "  $openaiModels = New-Object System.Collections.ArrayList" & vbCrLf
    lines = lines & "  $anthropicModels = New-Object System.Collections.ArrayList" & vbCrLf
    lines = lines & "  try {" & vbCrLf
    lines = lines & "    $m = Invoke-RestMethod -Method Get -Uri ($openaiUrl + '/models') -Headers $headers -TimeoutSec 20" & vbCrLf
    lines = lines & "    if ($m -and $m.data) { foreach ($x in $m.data) { if ($x.id) { $null = $openaiModels.Add([string]$x.id) } } }" & vbCrLf
    lines = lines & "  } catch {}" & vbCrLf
    lines = lines & "  try {" & vbCrLf
    lines = lines & "    $m = Invoke-RestMethod -Method Get -Uri $anthropicModelsUrl -Headers $headers -TimeoutSec 20" & vbCrLf
    lines = lines & "    if ($m -and $m.data) { foreach ($x in $m.data) { if ($x.id) { $null = $anthropicModels.Add([string]$x.id) } } }" & vbCrLf
    lines = lines & "  } catch {}" & vbCrLf
    lines = lines & "  if ($isGnai) {" & vbCrLf
    lines = lines & "    if ($openaiModels.Count -eq 0) {" & vbCrLf
    lines = lines & "      foreach ($id in @('gpt-4o','gpt-4.1','gpt-5-mini','gpt-5-nano','o3-mini')) { $null = $openaiModels.Add($id) }" & vbCrLf
    lines = lines & "    }" & vbCrLf
    lines = lines & "    if ($anthropicModels.Count -eq 0) {" & vbCrLf
    lines = lines & "      foreach ($id in @('claude-4-6-opus','claude-4-6-sonnet','claude-4-5-opus','claude-4-5-sonnet','claude-4-5-haiku')) { $null = $anthropicModels.Add($id) }" & vbCrLf
    lines = lines & "    }" & vbCrLf
    lines = lines & "  }" & vbCrLf
    lines = lines & "  $seen = @{}" & vbCrLf
    lines = lines & "  function Add-ModelGroup([string]$groupName, [System.Collections.ArrayList]$models) {" & vbCrLf
    lines = lines & "    if (-not $models -or $models.Count -eq 0) { return }" & vbCrLf
    lines = lines & "    $header = ""--- $groupName ---""" & vbCrLf
    lines = lines & "    $null = $combo.Items.Add($header)" & vbCrLf
    lines = lines & "    foreach ($id in $models) {" & vbCrLf
    lines = lines & "      if ($seen.ContainsKey($id)) { continue }" & vbCrLf
    lines = lines & "      $seen[$id] = $true" & vbCrLf
    lines = lines & "      $used = 0; $limit = 0" & vbCrLf
    lines = lines & "      if ($quotaMap.ContainsKey($id)) {" & vbCrLf
    lines = lines & "        try { $used = [int]$quotaMap[$id].used } catch {}" & vbCrLf
    lines = lines & "        try { $limit = [int]$quotaMap[$id].limit } catch {}" & vbCrLf
    lines = lines & "      }" & vbCrLf
    lines = lines & "      $display = if ($isGnai) { ""$id"" } else { ""$id ($used/$limit)"" }" & vbCrLf
    lines = lines & "      $null = $combo.Items.Add($display)" & vbCrLf
    lines = lines & "      $modelMap[$display] = $id" & vbCrLf
    lines = lines & "    }" & vbCrLf
    lines = lines & "  }" & vbCrLf
    lines = lines & "  Add-ModelGroup 'OpenAI' $openaiModels" & vbCrLf
    lines = lines & "  Add-ModelGroup 'Anthropic' $anthropicModels" & vbCrLf
    lines = lines & "  if ($modelMap.Count -eq 0) {" & vbCrLf
    lines = lines & "    $display = ""$defaultModelId (0/0)""" & vbCrLf
    lines = lines & "    $null = $combo.Items.Add($display)" & vbCrLf
    lines = lines & "    $modelMap[$display] = $defaultModelId" & vbCrLf
    lines = lines & "  }" & vbCrLf
    lines = lines & "  $selectedIndex = -1" & vbCrLf
    lines = lines & "  if (-not [string]::IsNullOrWhiteSpace($initialModel)) {" & vbCrLf
    lines = lines & "    for ($i = 0; $i -lt $combo.Items.Count; $i++) {" & vbCrLf
    lines = lines & "      $txt = [string]$combo.Items[$i]" & vbCrLf
    lines = lines & "      if ($modelMap.ContainsKey($txt)) {" & vbCrLf
    lines = lines & "        $candidateId = [string]$modelMap[$txt]" & vbCrLf
    lines = lines & "        if ([string]::Equals($candidateId, $initialModel, [System.StringComparison]::OrdinalIgnoreCase)) { $selectedIndex = $i; break }" & vbCrLf
    lines = lines & "      }" & vbCrLf
    lines = lines & "    }" & vbCrLf
    lines = lines & "  }" & vbCrLf
    lines = lines & "  if ($selectedIndex -lt 0) {" & vbCrLf
    lines = lines & "    for ($i = 0; $i -lt $combo.Items.Count; $i++) {" & vbCrLf
    lines = lines & "      $txt = [string]$combo.Items[$i]" & vbCrLf
    lines = lines & "      if ($modelMap.ContainsKey($txt)) { $selectedIndex = $i; break }" & vbCrLf
    lines = lines & "    }" & vbCrLf
    lines = lines & "  }" & vbCrLf
    lines = lines & "  if ($selectedIndex -ge 0) { $combo.SelectedIndex = $selectedIndex }" & vbCrLf
    lines = lines & "  $labelModel.Text = ""Model:  ($($modelMap.Count) models loaded)""" & vbCrLf
    lines = lines & "}" & vbCrLf
    lines = lines & "$combo.Add_SelectedIndexChanged({" & vbCrLf
    lines = lines & "  if ($script:skipIndexChange) { return }" & vbCrLf
    lines = lines & "  if ($combo.SelectedIndex -lt 0) { return }" & vbCrLf
    lines = lines & "  $txt = [string]$combo.SelectedItem" & vbCrLf
    lines = lines & "  if (-not $modelMap.ContainsKey($txt)) {" & vbCrLf
    lines = lines & "    $script:skipIndexChange = $true" & vbCrLf
    lines = lines & "    $found = $false" & vbCrLf
    lines = lines & "    for ($i = $combo.SelectedIndex + 1; $i -lt $combo.Items.Count; $i++) {" & vbCrLf
    lines = lines & "      if ($modelMap.ContainsKey([string]$combo.Items[$i])) { $combo.SelectedIndex = $i; $found = $true; break }" & vbCrLf
    lines = lines & "    }" & vbCrLf
    lines = lines & "    if (-not $found) {" & vbCrLf
    lines = lines & "      for ($i = $combo.SelectedIndex - 1; $i -ge 0; $i--) {" & vbCrLf
    lines = lines & "        if ($modelMap.ContainsKey([string]$combo.Items[$i])) { $combo.SelectedIndex = $i; break }" & vbCrLf
    lines = lines & "      }" & vbCrLf
    lines = lines & "    }" & vbCrLf
    lines = lines & "    $script:skipIndexChange = $false" & vbCrLf
    lines = lines & "  }" & vbCrLf
    lines = lines & "})" & vbCrLf
    lines = lines & "$btnLoad.Add_Click({ Fill-ModelList($txtKey.Text.Trim()) })" & vbCrLf
    lines = lines & "$txtKey.Add_Leave({ Fill-ModelList($txtKey.Text.Trim()) })" & vbCrLf
    lines = lines & "$okBtn.Add_Click({" & vbCrLf
    lines = lines & "  $k = $txtKey.Text.Trim()" & vbCrLf
    lines = lines & "  if ([string]::IsNullOrWhiteSpace($k)) { [System.Windows.Forms.MessageBox]::Show('Please enter a valid API key (ExpertGPT pak_ or GNAI).','ExpertGPT',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null; return }" & vbCrLf
    lines = lines & "  if ($combo.SelectedIndex -lt 0) { [System.Windows.Forms.MessageBox]::Show('Please select a model.','ExpertGPT',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null; return }" & vbCrLf
    lines = lines & "  $display = [string]$combo.SelectedItem" & vbCrLf
    lines = lines & "  if (-not $modelMap.ContainsKey($display)) { [System.Windows.Forms.MessageBox]::Show('Please select a model.','ExpertGPT',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null; return }" & vbCrLf
    lines = lines & "  $modelId = [string]$modelMap[$display]" & vbCrLf
    lines = lines & "  Set-Content -LiteralPath '" & outputEscaped & "' -Value ($k + [char]9 + $modelId) -Encoding ASCII" & vbCrLf
    lines = lines & "  $form.Tag = 'OK'" & vbCrLf
    lines = lines & "  $form.Close()" & vbCrLf
    lines = lines & "})" & vbCrLf
    lines = lines & "$cancelBtn.Add_Click({ $form.Tag = 'CANCEL'; $form.Close() })" & vbCrLf
    lines = lines & "$form.AcceptButton = $okBtn" & vbCrLf
    lines = lines & "$form.CancelButton = $cancelBtn" & vbCrLf
    lines = lines & "$form.Controls.Add($labelKey)" & vbCrLf
    lines = lines & "$form.Controls.Add($txtKey)" & vbCrLf
    lines = lines & "$form.Controls.Add($labelModel)" & vbCrLf
    lines = lines & "$form.Controls.Add($combo)" & vbCrLf
    lines = lines & "$form.Controls.Add($btnLoad)" & vbCrLf
    lines = lines & "$form.Controls.Add($okBtn)" & vbCrLf
    lines = lines & "$form.Controls.Add($cancelBtn)" & vbCrLf
    lines = lines & "$form.Add_Shown({ Fill-ModelList($txtKey.Text.Trim()); $txtKey.Focus() })" & vbCrLf
    lines = lines & "$null = $form.ShowDialog()" & vbCrLf

    BuildConfigurationFormScript = lines
End Function

Sub ExpertGPT_RefreshModelSelection()
    On Error GoTo ErrorHandler

    Dim apiKey As String
    Dim selectedModel As String

    apiKey = GetStoredApiKey()
    selectedModel = GetStoredModel()
    If Len(selectedModel) = 0 Then
        selectedModel = DEFAULT_MODEL
    End If

    If ShowConfigurationForm(apiKey, selectedModel) Then
        MsgBox "Model selected: " & selectedModel, vbInformation, "Key Assist"
    Else
        MsgBox "Configuration was cancelled. Model selected: " & selectedModel, vbInformation, "Key Assist"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error refreshing models: " & Err.Description, vbCritical, "Error"
End Sub

' ------------------------------------------------------------
' Common processing entry point
' ------------------------------------------------------------
Private Sub ProcessEmailWithPromptExpertGPT(ByVal userPrompt As String, Optional ByVal subjectPrefix As String = "AI Processed: ")
    On Error GoTo ErrorHandler

    Dim objMail As Outlook.MailItem
    Dim resultMail As Outlook.MailItem
    Dim emailContent As String
    Dim apiKey As String
    Dim selectedModel As String
    Dim aiResponse As String

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation, "No Email Selected"
        Exit Sub
    End If

    If TypeName(Application.ActiveExplorer.Selection.Item(1)) <> "MailItem" Then
        MsgBox "Please select an email message.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    apiKey = EnsureApiKeyReady()
    If apiKey = "" Then
        Exit Sub
    End If

    selectedModel = EnsureModelReady(apiKey)
    If selectedModel = "" Then
        Exit Sub
    End If

    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    emailContent = NzStr(objMail.Body)

    aiResponse = ProcessWithExpertGPT(emailContent, userPrompt, apiKey, selectedModel)

    Set resultMail = Application.CreateItem(olMailItem)
    With resultMail
        .To = GetCurrentUserSmtpAddress()
        .Subject = subjectPrefix & objMail.Subject
        .Body = aiResponse
        .Display
    End With

    Set resultMail = Nothing
    Set objMail = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error processing email: " & Err.Description, vbCritical, "Error"
End Sub

' ------------------------------------------------------------
' Settings persistence
' ------------------------------------------------------------
Private Function EnsureApiKeyReady() As String
    Dim apiKey As String
    Dim selectedModel As String

    apiKey = GetStoredApiKey()
    If IsValidApiKey(apiKey) Then
        EnsureApiKeyReady = apiKey
        Exit Function
    End If

    selectedModel = GetStoredModel()
    If Len(selectedModel) = 0 Then
        selectedModel = DEFAULT_MODEL
    End If

    If ShowConfigurationForm(apiKey, selectedModel) Then
        EnsureApiKeyReady = apiKey
    Else
        EnsureApiKeyReady = ""
    End If
End Function

Private Function EnsureModelReady(ByVal apiKey As String) As String
    Dim selectedModel As String
    Dim currentApiKey As String

    selectedModel = GetStoredModel()
    If Len(selectedModel) > 0 Then
        EnsureModelReady = selectedModel
        Exit Function
    End If

    currentApiKey = GetStoredApiKey()
    If Not IsValidApiKey(currentApiKey) Then
        currentApiKey = apiKey
    End If

    selectedModel = DEFAULT_MODEL
    If ShowConfigurationForm(currentApiKey, selectedModel) Then
        EnsureModelReady = selectedModel
    Else
        EnsureModelReady = ""
    End If
End Function

Private Function GetStoredApiKey() As String
    GetStoredApiKey = Trim$(GetSetting(APP_NAME, SETTINGS_SECTION, "ApiKey", ""))
End Function

Private Function GetStoredModel() As String
    GetStoredModel = Trim$(GetSetting(APP_NAME, SETTINGS_SECTION, "SelectedModel", ""))
End Function

Private Sub SaveSelectedModel(ByVal modelName As String)
    SaveSetting APP_NAME, SETTINGS_SECTION, "SelectedModel", modelName
End Sub

Private Function IsValidApiKey(ByVal apiKey As String) As Boolean
    Dim trimmed As String
    trimmed = Trim$(apiKey)
    If Len(trimmed) = 0 Then
        IsValidApiKey = False
    ElseIf Left$(trimmed, 4) = "pak_" Then
        IsValidApiKey = (Len(trimmed) > 8)
    Else
        ' GNAI key (non-pak_): accept any non-empty value
        IsValidApiKey = True
    End If
End Function

Private Function IsGnaiKey(ByVal apiKey As String) As Boolean
    Dim trimmed As String
    trimmed = Trim$(apiKey)
    IsGnaiKey = (Len(trimmed) > 0 And Left$(trimmed, 4) <> "pak_")
End Function

Private Function NormalizeApiKey(ByVal rawValue As String) As String
    Dim cleaned As String

    cleaned = CStr(rawValue)
    cleaned = Replace$(cleaned, ChrW$(&HFEFF), "") ' UTF-8/UTF-16 BOM
    cleaned = Replace$(cleaned, ChrW$(&H200B), "") ' zero-width space
    cleaned = Replace$(cleaned, vbCr, "")
    cleaned = Replace$(cleaned, vbLf, "")
    cleaned = Replace$(cleaned, vbTab, "")
    NormalizeApiKey = Trim$(cleaned)
End Function

Private Function NormalizeModelName(ByVal rawValue As String) As String
    Dim cleaned As String

    cleaned = CStr(rawValue)
    cleaned = Replace$(cleaned, ChrW$(&HFEFF), "") ' UTF-8/UTF-16 BOM
    cleaned = Replace$(cleaned, ChrW$(&H200B), "") ' zero-width space
    cleaned = Replace$(cleaned, vbCr, "")
    cleaned = Replace$(cleaned, vbLf, "")
    cleaned = Replace$(cleaned, vbTab, "")
    NormalizeModelName = Trim$(cleaned)
End Function

Private Sub WriteAllTextFile(ByVal filePath As String, ByVal content As String)
    Dim fso As Object
    Dim fileObj As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObj = fso.OpenTextFile(filePath, 2, True, 0)
    fileObj.Write content
    fileObj.Close
End Sub

Private Function ReadAllTextFile(ByVal filePath As String) As String
    Dim fso As Object
    Dim fileObj As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObj = fso.OpenTextFile(filePath, 1, False, 0)
    ReadAllTextFile = fileObj.ReadAll
    fileObj.Close
End Function

Private Function EscapeForPowerShellSingleQuoted(ByVal value As String) As String
    EscapeForPowerShellSingleQuoted = Replace$(value, "'", "''")
End Function

' ------------------------------------------------------------
' ExpertGPT API calls
' ------------------------------------------------------------
Private Function ProcessWithExpertGPT(ByVal emailContent As String, ByVal userPrompt As String, ByVal apiKey As String, ByVal modelName As String) As String
    On Error GoTo Fail

    Dim endpoint As String
    Dim payload As String
    Dim responseText As String
    Dim content As String
    Dim cleanedModelName As String

    cleanedModelName = NormalizeModelName(modelName)

    If IsAnthropicModel(cleanedModelName) Then
        If IsGnaiKey(apiKey) Then
            endpoint = GNAI_ANTHROPIC_BASE_URL & "/v1/messages"
            payload = BuildAnthropicPayload(emailContent, userPrompt, cleanedModelName)
            ' GNAI gateway does not require anthropic-version header
            responseText = SendHttpRequest("POST", endpoint, apiKey, payload, "", ANTHROPIC_REQUEST_TIMEOUT_MS)
        Else
            endpoint = ANTHROPIC_BASE_URL & "/messages"
            payload = BuildAnthropicPayload(emailContent, userPrompt, cleanedModelName)
            responseText = SendHttpRequest("POST", endpoint, apiKey, payload, ANTHROPIC_API_VERSION, ANTHROPIC_REQUEST_TIMEOUT_MS)
        End If
        content = ParseAnthropicContent(responseText)
    Else
        If IsGnaiKey(apiKey) Then
            endpoint = GNAI_OPENAI_BASE_URL & "/chat/completions"
        Else
            endpoint = OPENAI_BASE_URL & "/chat/completions"
        End If
        payload = BuildChatPayload(emailContent, userPrompt, cleanedModelName)
        responseText = SendHttpRequest("POST", endpoint, apiKey, payload)
        content = ParseChatContent(responseText)
    End If

    content = UnescapeJsonString(content)

    If Len(content) = 0 Then
        ProcessWithExpertGPT = "Error: Unable to process the email. Please try again."
    Else
        ProcessWithExpertGPT = content
    End If
    Exit Function

Fail:
    ProcessWithExpertGPT = "Error occurred: " & Err.Description
End Function

Private Function IsAnthropicModel(ByVal modelName As String) As Boolean
    Dim normalized As String

    normalized = LCase$(Trim$(modelName))
    IsAnthropicModel = (InStr(1, normalized, "claude", vbTextCompare) > 0 Or InStr(1, normalized, "anthropic", vbTextCompare) > 0)
End Function

Private Function BuildChatPayload(ByVal emailContent As String, ByVal userPrompt As String, ByVal modelName As String) As String
    Dim systemPrompt As String
    Dim userContent As String
    Dim tokenPart As String
    Dim lowered As String

    systemPrompt = "You are a helpful email assistant. Process the email content according to the user's instruction. Provide clear, concise, and well-formatted responses. Use plain text format with line breaks for readability."
    userContent = "Instruction: " & userPrompt & vbCrLf & vbCrLf & "Email Content:" & vbCrLf & emailContent

    ' Reasoning models (gpt-5*, o-series) require max_completion_tokens and reject temperature
    lowered = LCase$(Trim$(modelName))
    If Left$(lowered, 5) = "gpt-5" Or (Len(lowered) >= 2 And Left$(lowered, 1) = "o" And IsNumeric(Mid$(lowered, 2, 1))) Then
        tokenPart = """max_completion_tokens"":2000"
    Else
        tokenPart = """temperature"":0.7,""max_tokens"":2000"
    End If

    BuildChatPayload = _
        "{" & _
        """model"":""" & JsonEscape(modelName) & """," & _
        """messages"":[" & _
            "{""role"":""system"",""content"":""" & JsonEscape(systemPrompt) & """}," & _
            "{""role"":""user"",""content"":""" & JsonEscape(userContent) & """}" & _
        "]," & _
        """stream"":false," & _
        tokenPart & _
        "}"
End Function

Private Function BuildAnthropicPayload(ByVal emailContent As String, ByVal userPrompt As String, ByVal modelName As String) As String
    Dim systemPrompt As String
    Dim userContent As String

    systemPrompt = "You are a helpful email assistant. Process the email content according to the user's instruction. Provide clear, concise, and well-formatted responses. Use plain text format with line breaks for readability."
    userContent = "Instruction: " & userPrompt & vbCrLf & vbCrLf & "Email Content:" & vbCrLf & emailContent

    BuildAnthropicPayload = _
        "{" & _
        """model"":""" & JsonEscape(modelName) & """," & _
        """system"":""" & JsonEscape(systemPrompt) & """," & _
        """messages"": [" & _
            "{""role"":""user"",""content"": [{""type"":""text"",""text"":""" & JsonEscape(userContent) & """}]}" & _
        "]," & _
        """temperature"":0.7," & _
        """max_tokens"":1200" & _
        "}"
End Function

Private Function SendHttpRequest(ByVal method As String, ByVal endpoint As String, ByVal apiKey As String, ByVal payload As String, Optional ByVal anthropicVersion As String = "", Optional ByVal timeoutMs As Long = 0) As String
    Dim http As Object
    Dim statusCode As Long
    Dim responseText As String
    Dim effectiveTimeout As Long

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If timeoutMs > 0 Then
        effectiveTimeout = timeoutMs
    Else
        effectiveTimeout = REQUEST_TIMEOUT_MS
    End If
    http.setTimeouts effectiveTimeout, effectiveTimeout, effectiveTimeout, effectiveTimeout
    http.Open method, endpoint, False
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.setRequestHeader "Content-Type", "application/json"
    If Len(anthropicVersion) > 0 Then
        http.setRequestHeader "anthropic-version", anthropicVersion
    End If

    If UCase$(method) = "POST" Then
        http.Send payload
    Else
        http.Send
    End If

    statusCode = CLng(http.Status)
    responseText = NzStr(http.responseText)

    If statusCode < 200 Or statusCode >= 300 Then
        Err.Raise vbObjectError + 513, , endpoint & " failed (" & statusCode & "): " & responseText
    End If

    SendHttpRequest = responseText
End Function

' ------------------------------------------------------------
' Response parsing
' ------------------------------------------------------------

Private Function ParseAnthropicContent(ByVal jsonResponse As String) As String
    Dim re As Object
    Dim matches As Object
    Dim match As Object
    Dim combinedText As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = """text""\s*:\s*""((?:\\.|[^""\\])*)"""

    Set matches = re.Execute(jsonResponse)
    For Each match In matches
        If Len(combinedText) > 0 Then
            combinedText = combinedText & "\n\n"
        End If
        combinedText = combinedText & CStr(match.SubMatches(0))
    Next match

    ParseAnthropicContent = combinedText
End Function

Private Function ParseChatContent(ByVal jsonResponse As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim anchor As Long

    anchor = InStr(1, jsonResponse, """message""", vbTextCompare)
    If anchor = 0 Then anchor = 1

        startPos = InStr(anchor, jsonResponse, """content"":")
    If startPos = 0 Then
        ParseChatContent = ""
        Exit Function
    End If

        startPos = startPos + Len("""content"":")
    Do While startPos <= Len(jsonResponse) And (Mid$(jsonResponse, startPos, 1) = " " Or Mid$(jsonResponse, startPos, 1) = vbTab)
        startPos = startPos + 1
    Loop

    If Mid$(jsonResponse, startPos, 1) <> """" Then
        ParseChatContent = ""
        Exit Function
    End If

    startPos = startPos + 1
    endPos = FindJsonStringEnd(jsonResponse, startPos)

    If endPos > startPos Then
        ParseChatContent = Mid$(jsonResponse, startPos, endPos - startPos)
    Else
        ParseChatContent = ""
    End If
End Function

Private Function FindJsonStringEnd(ByVal text As String, ByVal startPos As Long) As Long
    Dim index As Long
    Dim escaped As Boolean
    Dim ch As String

    For index = startPos To Len(text)
        ch = Mid$(text, index, 1)

        If escaped Then
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            FindJsonStringEnd = index
            Exit Function
        End If
    Next index
End Function

' ------------------------------------------------------------
' String helpers
' ------------------------------------------------------------
Private Function JsonEscape(ByVal s As String) As String
    Dim t As String

    t = s
    t = Replace$(t, "\", "\\")
    t = Replace$(t, """", "\""")
    t = Replace$(t, vbCrLf, "\n")
    t = Replace$(t, vbCr, "\n")
    t = Replace$(t, vbLf, "\n")
    t = Replace$(t, vbTab, "\t")
    JsonEscape = t
End Function

Private Function UnescapeJsonString(ByVal s As String) As String
    Dim t As String
    Dim re As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long

    t = s
    t = Replace$(t, "\n", vbCrLf)
    t = Replace$(t, "\t", vbTab)

    ' Unescape \uXXXX Unicode sequences before handling \\
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "\\u([0-9A-Fa-f]{4})"
    Set matches = re.Execute(t)
    For i = matches.Count - 1 To 0 Step -1
        Set match = matches(i)
        t = Left$(t, match.FirstIndex) & ChrW(CLng("&H" & match.SubMatches(0))) & Mid$(t, match.FirstIndex + match.Length + 1)
    Next i

    t = Replace$(t, "\\", "\")
    t = Replace$(t, "\""", """")
    UnescapeJsonString = t
End Function

Private Function NzStr(ByVal s As Variant) As String
    If IsNull(s) Then
        NzStr = ""
    Else
        NzStr = CStr(s)
    End If
End Function