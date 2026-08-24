Attribute VB_Name = "Module7"
'-------------------------------------------------------------------------
' OutlookMailMigrator (OMMigrate)
'-------------------------------------------------------------------------
' Originator & Architect:    Kirk Shallcross - Shallcross Consulting
' Implementation Specialist: Gemini & Anthropic Claude AI
' Inception Date:            May 2026
' Version:                   1.5.2 (On-Behalf-Of MAPI Routing Engine)
'-------------------------------------------------------------------------
Option Explicit

' ==============================================================================
' MASTER ARCHIVE FOLDER CORRECTION -- CSV-DRIVEN, SINGLE-STORE ONLY
' ==============================================================================
Private Const WORKING_MISROUTES_FOLDER_NAME As String = "Working MisRoutes"

Private Function RULES_CSV_BASE_PATH() As String
    RULES_CSV_BASE_PATH = Environ("USERPROFILE") & "\Documents\OutlookMigration\Config\rules_inventory_"
End Function

' ==============================================================================
' MAIN ENTRY POINT
' ==============================================================================
Public Sub CorrectArchiveFolders()
    Dim outlookApp As Outlook.Application
    Dim ns As Outlook.NameSpace
    Dim archiveStore As Outlook.Store
    Dim archiveRoot As Outlook.folder
    Dim workingMisroutesFolder As Outlook.folder
    Dim senderMap As Object      ' Scripting.Dictionary: sender domain/address (lowercase) -> TargetFolderPath
    Dim movedCount As Long
    Dim notFoundCount As Long
    Dim pass2Recovered As Long
    Dim pass2Unresolved As Long
    Dim pass3Recovered As Long
    Dim archiveAccountName As String   ' resolved from the picker at runtime
    Dim rulesCsvPath As String         ' resolved from profile name at runtime

    Set outlookApp = New Outlook.Application
    Set ns = outlookApp.GetNamespace("MAPI")

    archiveAccountName = PromptForArchiveStore(ns)
    If archiveAccountName = "" Then
        MsgBox "No store selected -- nothing was scanned or modified.", vbInformation, "Cancelled"
        Exit Sub
    End If

    If MsgBox("This will scan and PERMANENTLY REARRANGE emails in:" & vbCrLf & vbCrLf & _
              "    " & archiveAccountName & vbCrLf & vbCrLf & _
              "True duplicates will be moved to a 'Duplicates' folder, and " & _
              "misrouted emails will be moved to match rules_inventory.csv." & vbCrLf & vbCrLf & _
              "Is this the correct store?", _
              vbYesNo Or vbQuestion Or vbDefaultButton2, "Confirm Store Selection") <> vbYes Then
        MsgBox "Cancelled -- nothing was scanned or modified.", vbInformation, "Cancelled"
        Exit Sub
    End If

    On Error Resume Next
    Load FrmProgress
    FrmProgress.Show vbModeless
    FrmProgress.UpdateProgress "Initializing...", 5
    On Error GoTo 0

    On Error Resume Next
    Set archiveStore = ns.Stores(archiveAccountName)
    On Error GoTo 0

    If archiveStore Is Nothing Then
        SafeUnloadUI
        MsgBox "Could not find the archive store named: '" & archiveAccountName & "'." & vbCrLf & _
               "Nothing was scanned or modified.", _
               vbCritical, "Execution Aborted"
        Exit Sub
    End If

    Set archiveRoot = archiveStore.GetRootFolder()
    rulesCsvPath = RULES_CSV_BASE_PATH & ns.currentProfileName & ".csv"

    UpdateUIProgress "Loading rules_inventory.csv...", 10
    Set senderMap = LoadSenderFolderMap(rulesCsvPath)

    If senderMap Is Nothing Or senderMap.Count = 0 Then
        SafeUnloadUI
        MsgBox "Could not load any sender/folder mappings from:" & vbCrLf & rulesCsvPath & vbCrLf & vbCrLf & _
               "Nothing was scanned or modified. Check RULES_CSV_BASE_PATH at the top of this module, " & _
               "and confirm this profile's rules_inventory CSV exists at that location.", _
               vbCritical, "Execution Aborted"
        Exit Sub
    End If

    UpdateUIProgress "Removing duplicates (Module1)...", 15
    Module1.DeleteDupsInFolder archiveRoot, False
    Dim dupesMoved As Long
    dupesMoved = Module1.gTotalDuplicatesMoved

    Set workingMisroutesFolder = GetOrCreateSubfolder(archiveRoot, WORKING_MISROUTES_FOLDER_NAME)
    If workingMisroutesFolder Is Nothing Then
        SafeUnloadUI
        MsgBox "Failed to create or access the '" & WORKING_MISROUTES_FOLDER_NAME & "' folder inside " & archiveAccountName & ".", _
               vbCritical, "Execution Aborted"
        Exit Sub
    End If

    ' -- PASS 1: sender-based correction over the whole archive ------------
    movedCount = 0
    notFoundCount = 0
    UpdateUIProgress "Pass 1: Scanning " & archiveAccountName & "...", 25
    ScanAndCorrectFolder archiveRoot, workingMisroutesFolder, senderMap, movedCount, notFoundCount, 25, 55

    ' -- PASS 2: recipient-based recovery over Working MisRoutes -----------
    pass2Recovered = 0
    pass2Unresolved = 0
    UpdateUIProgress "Pass 2: Checking recipients in " & WORKING_MISROUTES_FOLDER_NAME & "...", 60
    RunPass2RecipientRecovery workingMisroutesFolder, senderMap, pass2Recovered, pass2Unresolved, 60, 80

    ' -- PASS 3: re-checks Working MisRoutes against both ------------------
    pass3Recovered = 0
    UpdateUIProgress "Pass 3: Re-checking " & WORKING_MISROUTES_FOLDER_NAME & "...", 85
    RunPass3Recheck workingMisroutesFolder, senderMap, pass3Recovered, 85, 95

    UpdateUIProgress "Execution Complete!", 100
    SafeUnloadUI

    MsgBox "Process complete." & vbCrLf & vbCrLf & _
           "Store scanned                 : " & archiveAccountName & vbCrLf & _
           "Duplicates removed (Module1)  : " & dupesMoved & vbCrLf & vbCrLf & _
           "Pass 1 -- Emails relocated by sender    : " & movedCount & vbCrLf & _
           "Pass 1 -- Sender not in CSV             : " & notFoundCount & vbCrLf & vbCrLf & _
           "Pass 2 -- Recovered by recipient        : " & pass2Recovered & vbCrLf & _
           "Pass 2 -- Still unresolved              : " & pass2Unresolved & vbCrLf & vbCrLf & _
           "Pass 3 -- Recovered (rule added since)  : " & pass3Recovered & vbCrLf & vbCrLf & _
           "No other store was opened or modified.", _
           vbInformation, "Archive Correction Finished"
End Sub

' ==============================================================================
' CORE SCAN (PASS 1) -- Recurses through folders, checking sender mappings
' ==============================================================================
Private Sub ScanAndCorrectFolder(ByVal currentFolder As Outlook.folder, ByVal workingMisroutesFolder As Outlook.folder, _
                                  ByRef senderMap As Object, ByRef movedCount As Long, ByRef notFoundCount As Long, _
                                  ByVal startPct As Single, ByVal endPct As Single)
    Dim itemIndex As Long
    Dim currentItem As Object
    Dim mailItem As Outlook.mailItem
    Dim subFolder As Outlook.folder
    Dim senderKey As String
    Dim targetPath As String
    Dim destFolder As Outlook.folder

    If currentFolder.entryID = workingMisroutesFolder.entryID Then Exit Sub

    If StrComp(currentFolder.Name, "Duplicates", vbTextCompare) = 0 Then
        Dim parentEntryID As String
        Dim rootEntryIDCheck As String
        parentEntryID = ""
        On Error Resume Next
        parentEntryID = currentFolder.Parent.entryID
        On Error GoTo 0
        rootEntryIDCheck = currentFolder.Store.GetRootFolder().entryID
        If parentEntryID <> "" And parentEntryID = rootEntryIDCheck Then Exit Sub
    End If

    UpdateUIProgress "Scanning: " & currentFolder.Name, startPct

    If currentFolder.DefaultItemType = olMailItem Then
        For itemIndex = currentFolder.Items.Count To 1 Step -1
            Set currentItem = currentFolder.Items(itemIndex)

            If TypeOf currentItem Is Outlook.mailItem Then
                Set mailItem = currentItem

                senderKey = GetSenderKey(mailItem, senderMap)

                If senderKey <> "" Then
                    targetPath = senderMap(senderKey)

                    If Not FolderPathMatches(currentFolder, targetPath) Then
                        Set destFolder = GetOrCreateFolderByPath(currentFolder.Store.GetRootFolder(), targetPath)

                        If Not destFolder Is Nothing Then
                            mailItem.Move destFolder
                            movedCount = movedCount + 1
                        End If
                    End If
                Else
                    ' Bypassed: unmatched emails are safely left exactly where they are
                    notFoundCount = notFoundCount + 1
                End If
            End If
        Next itemIndex
    End If

    DoEvents

    For Each subFolder In currentFolder.Folders
        ScanAndCorrectFolder subFolder, workingMisroutesFolder, senderMap, movedCount, notFoundCount, startPct, endPct
    Next subFolder
End Sub

' ==============================================================================
' PASS 2 -- Recipient-based recovery over items in Working MisRoutes
' ==============================================================================
Private Sub RunPass2RecipientRecovery(ByVal workingMisroutesFolder As Outlook.folder, ByRef senderMap As Object, _
                                       ByRef recoveredCount As Long, ByRef unresolvedCount As Long, _
                                       ByVal startPct As Single, ByVal endPct As Single)
    Dim itemIndex As Long
    Dim currentItem As Object
    Dim mailItem As Outlook.mailItem
    Dim recipientKey As String
    Dim targetPath As String
    Dim destFolder As Outlook.folder
    Dim totalItems As Long

    totalItems = workingMisroutesFolder.Items.Count

    For itemIndex = totalItems To 1 Step -1
        Set currentItem = workingMisroutesFolder.Items(itemIndex)

        If TypeOf currentItem Is Outlook.mailItem Then
            Set mailItem = currentItem

            If (totalItems - itemIndex) Mod 25 = 0 Then
                UpdateUIProgress "Pass 2: Checking recipients (" & (totalItems - itemIndex + 1) & " / " & totalItems & ")...", _
                                 startPct + (((totalItems - itemIndex) / totalItems) * (endPct - startPct))
            End If

            recipientKey = GetRecipientKey(mailItem, senderMap)

            If recipientKey <> "" Then
                targetPath = senderMap(recipientKey)
                Set destFolder = GetOrCreateFolderByPath(workingMisroutesFolder.Store.GetRootFolder(), targetPath)

                If Not destFolder Is Nothing Then
                    mailItem.Move destFolder
                    recoveredCount = recoveredCount + 1
                Else
                    unresolvedCount = unresolvedCount + 1
                End If
            Else
                unresolvedCount = unresolvedCount + 1
            End If
        End If
    Next itemIndex

    DoEvents
End Sub

' ==============================================================================
' PASS 3 -- Re-checks Working MisRoutes against both sender and recipient
' ==============================================================================
Private Sub RunPass3Recheck(ByVal workingMisroutesFolder As Outlook.folder, ByRef senderMap As Object, _
                             ByRef recoveredCount As Long, ByVal startPct As Single, ByVal endPct As Single)
    Dim itemIndex As Long
    Dim currentItem As Object
    Dim mailItem As Outlook.mailItem
    Dim matchKey As String
    Dim targetPath As String
    Dim destFolder As Outlook.folder
    Dim totalItems As Long

    totalItems = workingMisroutesFolder.Items.Count

    For itemIndex = totalItems To 1 Step -1
        Set currentItem = workingMisroutesFolder.Items(itemIndex)

        If TypeOf currentItem Is Outlook.mailItem Then
            Set mailItem = currentItem

            If (totalItems - itemIndex) Mod 25 = 0 Then
                UpdateUIProgress "Pass 3: Re-checking (" & (totalItems - itemIndex + 1) & " / " & totalItems & ")...", _
                                 startPct + (((totalItems - itemIndex) / totalItems) * (endPct - startPct))
            End If

            matchKey = GetSenderKey(mailItem, senderMap)
            If matchKey = "" Then
                matchKey = GetRecipientKey(mailItem, senderMap)
            End If

            If matchKey <> "" Then
                targetPath = senderMap(matchKey)
                Set destFolder = GetOrCreateFolderByPath(workingMisroutesFolder.Store.GetRootFolder(), targetPath)

                If Not destFolder Is Nothing Then
                    mailItem.Move destFolder
                    recoveredCount = recoveredCount + 1
                End If
            End If
        End If
    Next itemIndex

    DoEvents
End Sub

' ==============================================================================
' CSV LOADER -- Builds lowercase map of SendersDomain -> TargetFolderPath.
' UPGRADED (2026-07-16): Reads as strict UTF-8 using ADODB.Stream instead of
' fso.OpenTextFile to ensure full parity with Module 3 database integrity.
' ==============================================================================
Private Function LoadSenderFolderMap(ByVal csvPath As String) As Object
    Dim map As Object
    Dim fso As Object
    Dim headerLine As String
    Dim dataLine As String
    Dim headers() As String
    Dim fields() As String
    Dim senderColIdx As Long
    Dim targetColIdx As Long
    Dim i As Long
    Dim senderVal As String
    Dim targetVal As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(csvPath) Then
        Set LoadSenderFolderMap = map
        Exit Function
    End If

    On Error GoTo LoadError

    ' Use ADODB.Stream for accurate UTF-8 reading and safe BOM stripping
    Dim csvStream As Object
    Set csvStream = CreateObject("ADODB.Stream")
    csvStream.Type = 2  ' adTypeText
    csvStream.Charset = "utf-8"
    csvStream.Open
    csvStream.LoadFromFile csvPath
    
    Dim csvRawText As String
    csvRawText = csvStream.ReadText(-1)  ' ReadAll
    csvStream.Close

    ' Safe BOM strip
    If Len(csvRawText) > 0 And AscW(Left(csvRawText, 1)) = 65279 Then
        csvRawText = Mid(csvRawText, 2)
    End If

    Dim allLines() As String
    allLines = Split(csvRawText, vbCrLf)

    If UBound(allLines) < 0 Then
        Set LoadSenderFolderMap = map
        Exit Function
    End If

    headerLine = allLines(0)
    headers = SplitCsvLine(headerLine)

    senderColIdx = -1
    targetColIdx = -1
    For i = 0 To UBound(headers)
        Select Case Trim(headers(i))
            Case "SendersDomain"
                senderColIdx = i
            Case "TargetFolderPath"
                targetColIdx = i
        End Select
    Next i

    If senderColIdx = -1 Or targetColIdx = -1 Then
        Set LoadSenderFolderMap = map
        Exit Function
    End If

    Dim lineIdx As Long
    For lineIdx = 1 To UBound(allLines)
        dataLine = allLines(lineIdx)
        If Trim(dataLine) <> "" Then
            fields = SplitCsvLine(dataLine)
            If UBound(fields) >= senderColIdx And UBound(fields) >= targetColIdx Then
                senderVal = Trim(fields(senderColIdx))
                targetVal = Trim(fields(targetColIdx))
                If senderVal <> "" And targetVal <> "" Then
                    Dim termParts() As String
                    Dim termIdx As Long
                    Dim term As String
                    termParts = Split(senderVal, " ")
                    For termIdx = 0 To UBound(termParts)
                        term = Trim(termParts(termIdx))
                        If term <> "" And Not map.Exists(LCase(term)) Then
                            map.Add LCase(term), targetVal
                        End If
                    Next termIdx
                End If
            End If
        End If
    Next lineIdx

    Set LoadSenderFolderMap = map
    Exit Function

LoadError:
    Set LoadSenderFolderMap = map
End Function

' ==============================================================================
' PARSER: Quote-Aware CSV Line Parser (Ported from Module 3 parity upgrade)
' ==============================================================================
Private Function SplitCsvLine(ByVal line As String) As String()
    Dim result() As String
    ReDim result(0)
    Dim idx As Long: idx = 0
    Dim pos As Long: pos = 1
    Dim length As Long: length = Len(line)
    Dim inQuotes As Boolean: inQuotes = False
    Dim char As String
    Dim value As String: value = ""
    
    Do While pos <= length
        char = Mid(line, pos, 1)
        If char = """" Then
            inQuotes = Not inQuotes
        ElseIf char = "," And Not inQuotes Then
            ReDim Preserve result(idx)
            result(idx) = value
            value = ""
            idx = idx + 1
        Else
            value = value & char
        End If
        pos = pos + 1
    Loop
    ReDim Preserve result(idx)
    result(idx) = value
    SplitCsvLine = result
End Function

' ==============================================================================
' STRICT ADDRESS & RECIPIENT MATCHERS (SUBSTRING LOOPS REMOVED)
' ==============================================================================

' Matches a mail item's SENDER (not recipient) against the CSV-derived map.
' Tries full sender email address first, then domain. NO substring fallback.
' UPGRADED (1.6.0): Explicitly intercepts MAPI 'PR_SENT_REPRESENTING_EMAIL_ADDRESS' 
' to gracefully handle "On Behalf Of" relay conditions (e.g. Salesforce CRM relays).
Private Function GetSenderKey(ByVal mail As Outlook.mailItem, ByRef senderMap As Object) As String
    Dim senderAddr As String
    Dim senderDomain As String
    Dim representingAddr As String
    Dim representingDomain As String
    Dim atPos As Long
    Dim propAccessor As Outlook.propertyAccessor
    Const PR_SENT_REPRESENTING_EMAIL_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"

    GetSenderKey = ""
    
    ' 1. Check for True Author ("On Behalf Of") email address via MAPI Property Tag
    On Error Resume Next
    Set propAccessor = mail.propertyAccessor
    representingAddr = LCase(Trim(propAccessor.GetProperty(PR_SENT_REPRESENTING_EMAIL_ADDRESS)))
    On Error GoTo 0

    If representingAddr <> "" Then
        atPos = InStr(representingAddr, "@")
        If atPos > 0 Then
            representingDomain = Mid(representingAddr, atPos + 1)
        Else
            representingDomain = ""
        End If

        ' Exact match on "On Behalf Of" address
        If senderMap.Exists(representingAddr) Then
            GetSenderKey = representingAddr
            Exit Function
        End If

        ' Domain match on "On Behalf Of" domain
        If representingDomain <> "" And senderMap.Exists(representingDomain) Then
            GetSenderKey = representingDomain
            Exit Function
        End If
    End If

    ' 2. Fallback to standard explicit Envelope Sender if no "On Behalf Of" match occurs
    On Error Resume Next
    senderAddr = LCase(Trim(mail.SenderEmailAddress))
    On Error GoTo 0

    If senderAddr = "" Then Exit Function

    atPos = InStr(senderAddr, "@")
    If atPos > 0 Then
        senderDomain = Mid(senderAddr, atPos + 1)
    Else
        senderDomain = ""
    End If

    ' Check exact envelope address match
    If senderMap.Exists(senderAddr) Then
        GetSenderKey = senderAddr
        Exit Function
    End If

    ' Check exact envelope domain match
    If senderDomain <> "" And senderMap.Exists(senderDomain) Then
        GetSenderKey = senderDomain
        Exit Function
    End If
End Function

' Matches a mail item's RECIPIENT (first recipient's address) against the
' CSV-derived map. Tries full email address first, then domain. NO substring fallback.
Private Function GetRecipientKey(ByVal mail As Outlook.mailItem, ByRef senderMap As Object) As String
    Dim recipAddr As String
    Dim recipDomain As String
    Dim atPos As Long

    GetRecipientKey = ""
    On Error Resume Next
    If mail.Recipients.Count > 0 Then
        recipAddr = LCase(Trim(mail.Recipients(1).Address))
    End If
    On Error GoTo 0

    If recipAddr = "" Then Exit Function

    atPos = InStr(recipAddr, "@")
    If atPos > 0 Then
        recipDomain = Mid(recipAddr, atPos + 1)
    Else
        recipDomain = ""
    End If

    ' 1. Check exact email address match
    If senderMap.Exists(recipAddr) Then
        GetRecipientKey = recipAddr
        Exit Function
    End If

    ' 2. Check exact domain match
    If recipDomain <> "" And senderMap.Exists(recipDomain) Then
        GetRecipientKey = recipDomain
        Exit Function
    End If
End Function

' ==============================================================================
' STRUCTURAL PATH HELPERS
' ==============================================================================
Private Function FolderPathMatches(ByVal folder As Outlook.folder, ByVal targetPath As String) As Boolean
    Dim currentPath As String
    currentPath = GetFolderRelativePath(folder)
    FolderPathMatches = (StrComp(currentPath, targetPath, vbTextCompare) = 0)
End Function

Private Function GetFolderRelativePath(ByVal folder As Outlook.folder) As String
    Dim pathParts As String
    Dim current As Outlook.folder
    Dim rootEntryID As String

    On Error Resume Next
    rootEntryID = folder.Store.GetRootFolder().entryID
    On Error GoTo 0

    pathParts = ""
    Set current = folder

    Do While Not current Is Nothing
        If current.entryID = rootEntryID Then Exit Do
        If pathParts = "" Then
            pathParts = current.Name
        Else
            pathParts = current.Name & "\" & pathParts
        End If
        On Error Resume Next
        Set current = current.Parent
        On Error GoTo 0
        If TypeName(current) <> "Folder" And TypeName(current) <> "MAPIFolder" Then Exit Do
    Loop

    GetFolderRelativePath = pathParts
End Function

Private Function GetOrCreateFolderByPath(ByVal rootFolder As Outlook.folder, ByVal folderPath As String) As Outlook.folder
    Dim parts() As String
    Dim current As Outlook.folder
    Dim nextFolder As Outlook.folder
    Dim i As Long

    On Error GoTo FailExit

    parts = Split(folderPath, "\")
    Set current = rootFolder

    For i = 0 To UBound(parts)
        If Trim(parts(i)) <> "" Then
            Set nextFolder = Nothing
            On Error Resume Next
            Set nextFolder = current.Folders(parts(i))
            On Error GoTo 0

            If nextFolder Is Nothing Then
                On Error Resume Next
                Set nextFolder = current.Folders.Add(parts(i), olFolderInbox)
                On Error GoTo 0
            End If

            If nextFolder Is Nothing Then GoTo FailExit
            Set current = nextFolder
        End If
    Next i

    Set GetOrCreateFolderByPath = current
    Exit Function

FailExit:
    Set GetOrCreateFolderByPath = Nothing
End Function

Private Function GetOrCreateSubfolder(ByVal parentFolder As Outlook.folder, ByVal subfolderName As String) As Outlook.folder
    Dim targetFolder As Outlook.folder

    On Error Resume Next
    Set targetFolder = parentFolder.Folders(subfolderName)
    On Error GoTo 0

    If targetFolder Is Nothing Then
        On Error Resume Next
        Set targetFolder = parentFolder.Folders.Add(subfolderName, olFolderInbox)
        On Error GoTo 0
    End If

    Set GetOrCreateSubfolder = targetFolder
End Function

Private Sub UpdateUIProgress(ByVal statusText As String, ByVal pct As Single)
    On Error Resume Next
    FrmProgress.UpdateProgress statusText, pct
    DoEvents
    On Error GoTo 0
End Sub

Private Sub SafeUnloadUI()
    On Error Resume Next
    Unload FrmProgress
    On Error GoTo 0
End Sub

Private Function PromptForArchiveStore(ByVal ns As Outlook.NameSpace) As String
    Dim storeNames As New Collection
    Dim i As Long
    Dim currentStore As Outlook.Store
    Dim filePath As String

    For i = 1 To ns.Stores.Count
        Set currentStore = ns.Stores(i)
        filePath = ""
        On Error Resume Next
        filePath = currentStore.filePath
        On Error GoTo 0

        If LCase(Right(filePath, 4)) = ".pst" Then
            storeNames.Add currentStore.DisplayName
        End If
    Next i

    If storeNames.Count = 0 Then
        MsgBox "No PST-backed stores are currently attached in this profile.", vbExclamation, "Nothing to Pick"
        PromptForArchiveStore = ""
        Exit Function
    End If

    Load FrmArchivePicker
    FrmArchivePicker.PopulateStores storeNames
    FrmArchivePicker.Show vbModal

    If FrmArchivePicker.WasCancelled Then
        PromptForArchiveStore = ""
    Else
        PromptForArchiveStore = FrmArchivePicker.SelectedStoreName
    End If

    Unload FrmArchivePicker
End Function