Attribute VB_Name = "Module3"
'-------------------------------------------------------------------------
' OutlookMailMigrator (OMMigrate)
'-------------------------------------------------------------------------
' Originator & Architect:    Kirk Shallcross - Shallcross Consulting
' Implementation Specialist: Anthropic Claude AI
' Inception Date:            May 2026
' Version:                   1.5.2 (Fixed CSV Serialization Engine)
'-------------------------------------------------------------------------
' ==============================================================================================
' PROJECT: OMMigrate Outlook Automation Pipeline
' FILE: Module3.bas — Rule Deployment, Consolidation, and Sorting Engine
'
' ORIGINAL ARCHITECTS: Kirk Shallcross - Shallcross Consulting
' CO-PILOT / ENHANCEMENTS: Gemini & Anthropic Claude AI
' LAST MODIFIED: 2026-07-17
'
' REVISION HISTORY:
'   * 2026-07-17 (Gemini): IMPLEMENTED SMART SENDER HARVESTING. Re-introduced rule condition 
'     harvesting with an SLD-matching regex engine. Automatically drops legacy bare-word remnants 
'     matching incoming CSV domains while strictly preserving genuine manual rule modifications.
'   * 2026-07-16 (Gemini): RESOLVED CRITICAL MULTI-ACCOUNT RULE DUPLICATION BUG. Removed loose
'     InStr() account-matching check inside ExecuteBatchBuild. Restored strict 1-to-1 account
'     binding equality to prevent single-account manual rules from propagating across all profile
'     mailboxes during execution.
'   * 2026-07-09 (Gemini): Implemented UI Rule Manager Condition Preservation Engine.
'     Extracts and retains manual conditions (e.g., Subject text, categories, body phrases)
'     while explicitly ignoring Account and From properties to protect execution flow.
' ==============================================================================================

Public LocalPathLookupTable As Object

' Holds one archive folder's live Outlook.folder object, keyed by the
' TargetStoreName string that names it
Public ResolvedArchiveRoots As Object

Public Type ArchiveMappingEntry
    targetStoreName As String
    RuleStoreNames() As String
    RuleStoreCount As Long
End Type

Public Function CONFIG_FOLDER_PATH() As String
    CONFIG_FOLDER_PATH = Environ("USERPROFILE") & "\Documents\OutlookMigration\Config\"
End Function

Sub NormalizeSenderDomains(ByVal rawValue As String, ByRef outArray() As String, ByRef outCount As Long)
    outCount = 0
    ReDim outArray(0)

    If Trim(rawValue) = "" Then Exit Sub

    Dim rawParts() As String: rawParts = Split(Trim(rawValue), " ")

    Dim emailRegex As Object: Set emailRegex = CreateObject("VBScript.RegExp")
    emailRegex.Pattern = "^[a-zA-Z0-9_+.-]+(\.[a-zA-Z0-9_+.-]+)*@([a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$"
    emailRegex.IgnoreCase = True

    Dim wordRegex As Object: Set wordRegex = CreateObject("VBScript.RegExp")
    wordRegex.Pattern = "^[a-zA-Z0-9]([a-zA-Z0-9._-]*[a-zA-Z0-9])?$"
    wordRegex.IgnoreCase = True

    Dim p As Long
    For p = LBound(rawParts) To UBound(rawParts)
        Dim trimmedPart As String: trimmedPart = LCase(Trim(rawParts(p)))
        If trimmedPart <> "" Then
            If emailRegex.Test(trimmedPart) Or wordRegex.Test(trimmedPart) Then
                ReDim Preserve outArray(outCount)
                outArray(outCount) = trimmedPart
                outCount = outCount + 1
            Else
                Debug.Print "NormalizeSenderDomains: dropped invalid SendersDomain token '" & trimmedPart & "' (matched neither email nor bare-word pattern)."
            End If
        End If
    Next p
End Sub

Private Function ParseCSVLine(ByVal textLine As String) As String()
    Dim result() As String
    ReDim result(0)
    Dim idx As Long: idx = 0
    Dim pos As Long: pos = 1
    Dim length As Long: length = Len(textLine)
    Dim inQuotes As Boolean: inQuotes = False
    Dim char As String
    Dim value As String: value = ""
    
    Do While pos <= length
        char = Mid(textLine, pos, 1)
        If char = """" Then
            If pos < length And Mid(textLine, pos + 1, 1) = """" Then
                value = value & """"
                pos = pos + 1
            Else
                inQuotes = Not inQuotes
            End If
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
    ParseCSVLine = result
End Function

Function GetActiveProfileName() As String
    Dim ns As Outlook.NameSpace: Set ns = Application.GetNamespace("MAPI")
    Dim rawName As String: rawName = ""
    On Error Resume Next
    rawName = ns.currentProfileName
    On Error GoTo 0

    If Trim(rawName) = "" Then
        GetActiveProfileName = ""
        Exit Function
    End If

    Dim sanitizeRegex As Object: Set sanitizeRegex = CreateObject("VBScript.RegExp")
    sanitizeRegex.Global = True
    sanitizeRegex.Pattern = "[\\/:\*\?""<>\|]"
    GetActiveProfileName = Trim(sanitizeRegex.Replace(rawName, "_"))
End Function

Private Function FindJSONStringValue(ByVal jsonText As String, ByVal keyName As String) As String
    Dim keyRegex As Object: Set keyRegex = CreateObject("VBScript.RegExp")
    keyRegex.IgnoreCase = False
    keyRegex.Global = False
    keyRegex.Pattern = """" & keyName & """\s*:\s*""((?:[^""\\]|\\.)*)"""

    Dim m As Object
    If keyRegex.Test(jsonText) Then
        Set m = keyRegex.Execute(jsonText)
        FindJSONStringValue = UnescapeJSONString(m(0).SubMatches(0))
    Else
        FindJSONStringValue = ""
    End If
End Function

Private Function UnescapeJSONString(ByVal rawText As String) As String
    Dim result As String: result = rawText
    result = Replace(result, "\r\n", vbCrLf)
    result = Replace(result, "\n", vbLf)
    result = Replace(result, "\r", vbCr)
    result = Replace(result, "\t", vbTab)
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")
    UnescapeJSONString = result
End Function

Private Function ExtractJSONStringArray(ByVal arrayText As String, ByRef outCount As Long) As String()
    Dim result() As String
    ReDim result(0)
    outCount = 0

    Dim itemRegex As Object: Set itemRegex = CreateObject("VBScript.RegExp")
    itemRegex.Global = True
    itemRegex.IgnoreCase = False
    itemRegex.Pattern = """((?:[^""\\]|\\.)*)"""

    Dim matches As Object: Set matches = itemRegex.Execute(arrayText)
    Dim m As Object
    For Each m In matches
        ReDim Preserve result(outCount)
        result(outCount) = UnescapeJSONString(m.SubMatches(0))
        outCount = outCount + 1
    Next m

    ExtractJSONStringArray = result
End Function

Private Function FindJSONArrayText(ByVal jsonText As String, ByVal keyName As String) As String
    Dim keyRegex As Object: Set keyRegex = CreateObject("VBScript.RegExp")
    keyRegex.IgnoreCase = False
    keyRegex.Pattern = """" & keyName & """\s*:\s*\["

    If Not keyRegex.Test(jsonText) Then
        FindJSONArrayText = ""
        Exit Function
    End If

    Dim m As Object: Set m = keyRegex.Execute(jsonText)(0)
    Dim openBracketPos As Long: openBracketPos = m.FirstIndex + Len(m.value)

    Dim depth As Long: depth = 1
    Dim pos As Long: pos = openBracketPos
    Dim textLen As Long: textLen = Len(jsonText)
    Dim inQuotes As Boolean: inQuotes = False
    Dim ch As String

    Do While pos < textLen And depth > 0
        ch = Mid(jsonText, pos + 1, 1)
        If ch = """" Then
            Dim precedingBackslashes As Long: precedingBackslashes = 0
            Dim scanPos As Long: scanPos = pos
            Do While scanPos > 0 And Mid(jsonText, scanPos, 1) = "\"
                precedingBackslashes = precedingBackslashes + 1
                scanPos = scanPos - 1
            Loop
            If (precedingBackslashes Mod 2) = 0 Then inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = "[" Then depth = depth + 1
            If ch = "]" Then depth = depth - 1
        End If
        pos = pos + 1
    Loop

    FindJSONArrayText = Mid(jsonText, openBracketPos + 1, (pos - 1) - openBracketPos)
End Function

Private Function SplitJSONObjectArray(ByVal arrayText As String, ByRef outCount As Long) As String()
    Dim result() As String
    ReDim result(0)
    outCount = 0

    Dim depth As Long: depth = 0
    Dim inQuotes As Boolean: inQuotes = False
    Dim itemStart As Long: itemStart = 0
    Dim pos As Long: pos = 1
    Dim textLen As Long: textLen = Len(arrayText)
    Dim ch As String

    Do While pos <= textLen
        ch = Mid(arrayText, pos, 1)
        If ch = """" Then
            Dim precedingBackslashes As Long: precedingBackslashes = 0
            Dim scanPos As Long: scanPos = pos - 1
            Do While scanPos >= 1 And Mid(arrayText, scanPos, 1) = "\"
                precedingBackslashes = precedingBackslashes + 1
                scanPos = scanPos - 1
            Loop
            If (precedingBackslashes Mod 2) = 0 Then inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = "{" Then
                If depth = 0 Then itemStart = pos
                depth = depth + 1
            ElseIf ch = "}" Then
                depth = depth - 1
                If depth = 0 Then
                    ReDim Preserve result(outCount)
                    result(outCount) = Mid(arrayText, itemStart, pos - itemStart + 1)
                    outCount = outCount + 1
                End If
            End If
        End If
        pos = pos + 1
    Loop

    SplitJSONObjectArray = result
End Function

Function ReadOMMigrateSettingsJSON(ByVal profileName As String, _
                                    ByRef outBasePath As String, _
                                    ByRef outMappings() As ArchiveMappingEntry, _
                                    ByRef outMappingCount As Long, _
                                    ByRef outMasterArchiveName As String, _
                                    ByRef outError As String) As Boolean
    ReadOMMigrateSettingsJSON = False
    outError = ""
    ReDim outMappings(0)
    outMappingCount = 0

    If Trim(profileName) = "" Then
        outError = "Could not determine the active Outlook profile name."
        Exit Function
    End If

    Dim settingsPath As String
    settingsPath = CONFIG_FOLDER_PATH & "OMMigrate_Settings_" & profileName & ".json"

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(settingsPath) Then
        outError = "Settings file not found: " & settingsPath
        Exit Function
    End If

    Dim jsonText As String
    Dim adoStream As Object: Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 2
    adoStream.Charset = "utf-8"
    adoStream.Open
    adoStream.LoadFromFile settingsPath
    jsonText = adoStream.ReadText(-1)
    adoStream.Close

    If Len(jsonText) > 0 And AscW(Left(jsonText, 1)) = 65279 Then
        jsonText = Mid(jsonText, 2)
    End If

    outBasePath = FindJSONStringValue(jsonText, "BasePath")
    If Trim(outBasePath) = "" Then
        outError = "Could not find DefaultPaths.BasePath in " & settingsPath
        Exit Function
    End If

    Dim masterArrayText As String
    masterArrayText = FindJSONArrayText(jsonText, "MasterArchiveNames")
    Dim masterCount As Long
    Dim masterNames() As String: masterNames = ExtractJSONStringArray(masterArrayText, masterCount)
    If masterCount = 0 Then
        outError = "Could not find RulesEngine.MasterArchiveNames in " & settingsPath
        Exit Function
    End If
    outMasterArchiveName = masterNames(0)

    Dim mappingsArrayText As String
    mappingsArrayText = FindJSONArrayText(jsonText, "ArchiveStoreMappings")
    Dim mappingObjCount As Long
    Dim mappingObjTexts() As String: mappingObjTexts = SplitJSONObjectArray(mappingsArrayText, mappingObjCount)

    If mappingObjCount = 0 Then
        ReadOMMigrateSettingsJSON = True
        Exit Function
    End If

    ReDim outMappings(mappingObjCount - 1)
    Dim moi As Long
    For moi = 0 To mappingObjCount - 1
        outMappings(moi).targetStoreName = FindJSONStringValue(mappingObjTexts(moi), "TargetStoreName")

        Dim ruleStoreArrayText As String
        ruleStoreArrayText = FindJSONArrayText(mappingObjTexts(moi), "RuleStoreNames")
        Dim rsCount As Long
        Dim rsNames() As String: rsNames = ExtractJSONStringArray(ruleStoreArrayText, rsCount)

        If rsCount > 0 Then
            ReDim outMappings(moi).RuleStoreNames(rsCount - 1)
            Dim rsi As Long
            For rsi = 0 To rsCount - 1
                outMappings(moi).RuleStoreNames(rsi) = rsNames(rsi)
            Next rsi
        End If
        outMappings(moi).RuleStoreCount = rsCount
    Next moi
    outMappingCount = mappingObjCount

    ReadOMMigrateSettingsJSON = True
End Function

Function ResolveArchiveStoreForAccount(ByVal ruleStoreName As String, _
                                        ByRef mappings() As ArchiveMappingEntry, _
                                        ByVal mappingCount As Long, _
                                        ByVal masterArchiveName As String) As String
    Dim needle As String: needle = LCase(Trim(ruleStoreName))

    Dim mi As Long, ri As Long
    For mi = 0 To mappingCount - 1
        For ri = 0 To mappings(mi).RuleStoreCount - 1
            If LCase(Trim(mappings(mi).RuleStoreNames(ri))) = needle Then
                ResolveArchiveStoreForAccount = mappings(mi).targetStoreName
                Exit Function
            End If
        Next ri
    Next mi

    ResolveArchiveStoreForAccount = masterArchiveName
End Function

Function GetResolvedArchiveRoot(ByVal ns As Outlook.NameSpace, ByVal targetStoreName As String) As Outlook.folder
    If ResolvedArchiveRoots Is Nothing Then
        Set ResolvedArchiveRoots = CreateObject("Scripting.Dictionary")
    End If

    If ResolvedArchiveRoots.Exists(targetStoreName) Then
        Set GetResolvedArchiveRoot = ResolvedArchiveRoots(targetStoreName)
        Exit Function
    End If

    Dim resolvedFolder As Outlook.folder: Set resolvedFolder = Nothing
    On Error Resume Next
    Set resolvedFolder = ns.Folders.item(targetStoreName)
    On Error GoTo 0

    If Not resolvedFolder Is Nothing Then
        Set ResolvedArchiveRoots(targetStoreName) = resolvedFolder
    End If

    Set GetResolvedArchiveRoot = resolvedFolder
End Function

Sub SortKeysArrayByLabel(ByRef keysArray() As Variant, ByVal itemCount As Long)
    If itemCount <= 1 Then Exit Sub

    Dim i As Long, j As Long
    Dim labelI As String, labelJ As String
    Dim temp As Variant

    For i = 0 To itemCount - 2
        For j = 0 To itemCount - 2 - i
            labelI = ExtractSortLabel(CStr(keysArray(j)))
            labelJ = ExtractSortLabel(CStr(keysArray(j + 1)))
            If labelI > labelJ Then
                temp = keysArray(j)
                keysArray(j) = keysArray(j + 1)
                keysArray(j + 1) = temp
            End If
        Next j
    Next i
End Sub

Private Function ExtractSortLabel(ByVal mapKeyText As String) As String
    Dim pipePos As Long: pipePos = InStr(mapKeyText, "|")
    Dim pathPart As String
    If pipePos > 0 Then
        pathPart = Mid(mapKeyText, pipePos + 1)
    Else
        pathPart = mapKeyText
    End If

    Dim segments() As String: segments = Split(pathPart, "\")
    ExtractSortLabel = LCase(Trim(segments(UBound(segments))))
End Function

Private Function ExtractSortLabelFromRuleName(ByVal ruleNameText As String) As String
    Dim result As String: result = ruleNameText

    If Left(ruleNameText, 7) = "Rule: [" Then
        Dim closeBracketPos As Long: closeBracketPos = InStr(8, ruleNameText, "] ")
        If closeBracketPos > 0 Then
            Dim afterAccount As String: afterAccount = Mid(ruleNameText, closeBracketPos + 2)
            Dim partPos As Long: partPos = InStrRev(afterAccount, " (Part ")
            If partPos > 0 And Right(afterAccount, 1) = ")" Then
                result = Left(afterAccount, partPos - 1)
            Else
                result = ruleNameText
            End If
        End If
    End If

    ExtractSortLabelFromRuleName = LCase(Trim(result))
End Function

Sub ResortRulesByLabel(ByVal targetRules As Outlook.Rules)
    Dim ruleCount As Long: ruleCount = targetRules.Count
    If ruleCount <= 1 Then Exit Sub

    Dim ruleSnapshot() As Outlook.Rule: ReDim ruleSnapshot(ruleCount - 1)
    Dim snapIdx As Long
    For snapIdx = 1 To ruleCount
        Set ruleSnapshot(snapIdx - 1) = targetRules.item(snapIdx)
    Next snapIdx

    Dim si As Long, sj As Long
    Dim labelI As String, labelJ As String
    Dim tempRule As Outlook.Rule
    For si = 0 To ruleCount - 2
        For sj = 0 To ruleCount - 2 - si
            labelI = ExtractSortLabelFromRuleName(ruleSnapshot(sj).Name)
            labelJ = ExtractSortLabelFromRuleName(ruleSnapshot(sj + 1).Name)
            If labelI > labelJ Then
                Set tempRule = ruleSnapshot(sj)
                Set ruleSnapshot(sj) = ruleSnapshot(sj + 1)
                Set ruleSnapshot(sj + 1) = tempRule
            End If
        Next sj
    Next si

    On Error Resume Next
    Dim writeIdx As Long
    For writeIdx = 0 To ruleCount - 1
        ruleSnapshot(writeIdx).ExecutionOrder = writeIdx + 1
    Next writeIdx
    On Error GoTo 0

    VBA.Interaction.CallByName targetRules, "Save", VbMethod
End Sub

Sub DeployConsolidatedRules()
    Dim csvPath As String, line As String
    Dim fileNum As Integer: fileNum = FreeFile
    Set LocalPathLookupTable = CreateObject("Scripting.Dictionary")
    Set ResolvedArchiveRoots = Nothing

    Dim ns As Outlook.NameSpace: Set ns = Application.GetNamespace("MAPI")
    Dim currentProfileName As String: currentProfileName = GetActiveProfileName()
    If currentProfileName = "" Then
        MsgBox "Could not determine the active Outlook profile.", vbCritical, "Profile Not Detected"
        Exit Sub
    End If

    Dim promptResult As VbMsgBoxResult
    promptResult = MsgBox("You are running this macro inside profile: '" & currentProfileName & "'." & vbCrLf & vbCrLf & _
                          "Do you want to apply/resume deployment using the LastDeployedRun timestamps?", _
                          vbYesNo + vbQuestion + vbDefaultButton1, "Profile Verification Reminder")
    If promptResult = vbNo Then Exit Sub

    Dim settingsBasePath As String
    Dim archiveMappings() As ArchiveMappingEntry
    Dim archiveMappingCount As Long
    Dim masterArchiveName As String
    Dim settingsError As String

    If Not ReadOMMigrateSettingsJSON(currentProfileName, settingsBasePath, archiveMappings, _
                                      archiveMappingCount, masterArchiveName, settingsError) Then
        MsgBox "Could not load OMMigrate settings for profile '" & currentProfileName & "':" & _
               vbCrLf & vbCrLf & settingsError, vbCritical, "Settings Load Failed"
        Exit Sub
    End If

    Dim masterArchiveRoot As Outlook.folder
    Set masterArchiveRoot = GetResolvedArchiveRoot(ns, masterArchiveName)
    If masterArchiveRoot Is Nothing Then
        MsgBox "Cannot find master/default archive store: '" & masterArchiveName & "'", vbCritical, "Archive Missing"
        Exit Sub
    End If

    csvPath = settingsBasePath & "\Config\rules_inventory_" & currentProfileName & ".csv"

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(csvPath) Then
        MsgBox "CSV File not found at: " & csvPath, vbCritical
        Exit Sub
    End If

    Dim csvAdoStream As Object: Set csvAdoStream = CreateObject("ADODB.Stream")
    csvAdoStream.Type = 2
    csvAdoStream.Charset = "utf-8"
    csvAdoStream.Open
    csvAdoStream.LoadFromFile csvPath
    Dim csvRawText As String: csvRawText = csvAdoStream.ReadText(-1)
    csvAdoStream.Close

    If Len(csvRawText) > 0 And AscW(Left(csvRawText, 1)) = 65279 Then
        csvRawText = Mid(csvRawText, 2)
    End If

    Dim allLines As Variant: allLines = Split(csvRawText, vbCrLf)

    Dim i As Long
    For i = LBound(allLines) + 1 To UBound(allLines)
        line = allLines(i)
        If Trim(line) <> "" Then
            Dim rawFields() As String: rawFields = ParseCSVLine(line)

            If UBound(rawFields) >= 6 Then
                Dim timestamp As String: timestamp = Trim(rawFields(3))
                Dim relativePath As String: relativePath = rawFields(5)
                Dim rawDomain As String: rawDomain = Trim(Replace(rawFields(6), vbTab, ""))

                Dim normDomains() As String, normCount As Long
                Call NormalizeSenderDomains(rawDomain, normDomains, normCount)

                If timestamp = "" And relativePath <> "" And normCount > 0 Then
                    Dim compositeGroupKey As String: compositeGroupKey = LCase(Trim(rawFields(0))) & "|" & relativePath
                    If Not LocalPathLookupTable.Exists(compositeGroupKey) Then
                        Set LocalPathLookupTable(compositeGroupKey) = CreateObject("Scripting.Dictionary")
                    End If
                    Dim nd As Long
                    For nd = 0 To normCount - 1
                        LocalPathLookupTable(compositeGroupKey)(normDomains(nd)) = True
                    Next nd
                End If
            End If
        End If
    Next i

    Call ExecuteBatchBuild(allLines, csvPath, LocalPathLookupTable, archiveMappings, archiveMappingCount, masterArchiveName)
End Sub

Sub ExecuteBatchBuild(ByRef allLines As Variant, ByVal csvPath As String, ByVal pathLookup As Object, _
                       ByRef archiveMappings() As ArchiveMappingEntry, ByVal archiveMappingCount As Long, _
                       ByVal masterArchiveName As String)
    Dim harvestedDuplicateCount As Long: harvestedDuplicateCount = 0
    Dim i As Long, processedCount As Long: processedCount = 0
    Dim ruleActionCounter As Long: ruleActionCounter = 0
    Dim olAcc As Outlook.Account, timestampString As String
    timestampString = Format(Now, "yyyy-mm-dd\THh:Nn:Ss") & ".000000-05:00"

    Dim ns As Outlook.NameSpace: Set ns = Application.GetNamespace("MAPI")
    Dim domainsInThisBatch As Object: Set domainsInThisBatch = CreateObject("Scripting.Dictionary")
    Dim keysProcessedThisBatch As Object: Set keysProcessedThisBatch = CreateObject("Scripting.Dictionary")

    Dim acctCount As Long: acctCount = 0
    On Error Resume Next: acctCount = Application.Session.Accounts.Count: On Error GoTo 0
    Dim acctIdx As Long
    For acctIdx = 1 To acctCount
        Set olAcc = Nothing
        On Error Resume Next: Set olAcc = Application.Session.Accounts.item(acctIdx): On Error GoTo 0
        If olAcc Is Nothing Then GoTo NextAccountBuild

        Dim accountRules As Outlook.Rules: Set accountRules = Nothing
        On Error Resume Next
        If Not olAcc.DeliveryStore Is Nothing Then
            If Not olAcc.DeliveryStore.GetRootFolder() Is Nothing Then
                Set accountRules = olAcc.DeliveryStore.GetRules()
            End If
        End If
        On Error GoTo 0

        If accountRules Is Nothing Then
            Dim thisAcctDisplay As String: thisAcctDisplay = ""
            On Error Resume Next: thisAcctDisplay = LCase(olAcc.DisplayName): On Error GoTo 0
            Dim thisAcctSmtpForStores As String: thisAcctSmtpForStores = ""
            On Error Resume Next: thisAcctSmtpForStores = LCase(olAcc.SmtpAddress): On Error GoTo 0

            Dim candidateStore As Outlook.Store
            Dim candidateDisplay As String
            Dim isMatch As Boolean
            For Each candidateStore In Application.Session.Stores
                candidateDisplay = ""
                On Error Resume Next: candidateDisplay = LCase(candidateStore.DisplayName): On Error GoTo 0
                If Len(candidateDisplay) > 0 Then
                    isMatch = False
                    If Len(thisAcctDisplay) > 0 And candidateDisplay = thisAcctDisplay Then isMatch = True
                    If Not isMatch And Len(thisAcctSmtpForStores) > 0 And candidateDisplay = thisAcctSmtpForStores Then isMatch = True

                    If isMatch Then
                        On Error Resume Next: Set accountRules = candidateStore.GetRules(): On Error GoTo 0
                        If Not accountRules Is Nothing Then Exit For
                    End If
                End If
            Next candidateStore
        End If

        If Not accountRules Is Nothing Then
            Dim targetAccAddress As String: targetAccAddress = LCase(Trim(olAcc.SmtpAddress))
            Dim targetAccDisplayName As String: targetAccDisplayName = LCase(Trim(olAcc.DisplayName))

            Dim keysArray() As Variant
            Dim kCount As Long: kCount = 0
            Dim k As Variant

            For Each k In pathLookup.Keys()
                ReDim Preserve keysArray(kCount)
                keysArray(kCount) = k
                kCount = kCount + 1
            Next k

            Call SortKeysArrayByLabel(keysArray, kCount)

            Dim mapIdx As Long
            For mapIdx = kCount - 1 To 0 Step -1
                Dim mapKey As Variant: mapKey = keysArray(mapIdx)
                Dim keySegments() As String: keySegments = Split(CStr(mapKey), "|")
                Dim ruleAccountBinding As String: ruleAccountBinding = LCase(Trim(keySegments(0)))
                Dim targetPath As String: targetPath = keySegments(1)

                Dim accountMatches As Boolean: accountMatches = False
                If ruleAccountBinding = targetAccAddress Then
                    accountMatches = True
                ElseIf ruleAccountBinding = targetAccDisplayName Then
                    accountMatches = True
                End If

                If accountMatches Then
                    Dim resolvedStoreName As String
                    resolvedStoreName = ResolveArchiveStoreForAccount(ruleAccountBinding, archiveMappings, archiveMappingCount, masterArchiveName)
                    Call BuildRulesFromMap(accountRules, pathLookup, mapKey, targetPath, ruleAccountBinding, olAcc, domainsInThisBatch, ruleActionCounter, processedCount, ns, resolvedStoreName, harvestedDuplicateCount)
                    keysProcessedThisBatch(CStr(mapKey)) = True
                End If
                If ruleActionCounter >= 1000 Then Exit For
            Next mapIdx
        End If
        If ruleActionCounter >= 1000 Then Exit For
NextAccountBuild:
    Next acctIdx

    Dim defaultStoreRules As Outlook.Rules: Set defaultStoreRules = Nothing
    On Error Resume Next
    Set defaultStoreRules = Application.Session.DefaultStore.GetRules()
    On Error GoTo 0

    If Not defaultStoreRules Is Nothing And ruleActionCounter < 1000 Then
        Dim remKeysArray() As Variant
        Dim remCount As Long: remCount = 0
        Dim rk As Variant
        For Each rk In pathLookup.Keys(): ReDim Preserve remKeysArray(remCount): remKeysArray(remCount) = rk: remCount = remCount + 1: Next rk

        Call SortKeysArrayByLabel(remKeysArray, remCount)

        Dim remIdx As Long
        For remIdx = remCount - 1 To 0 Step -1
            Dim remainingKey As Variant: remainingKey = remKeysArray(remIdx)

            If Not keysProcessedThisBatch.Exists(CStr(remainingKey)) Then
                Dim remSegments() As String: remSegments = Split(CStr(remainingKey), "|")
                Dim remAccountBinding As String: remAccountBinding = LCase(Trim(remSegments(0)))
                Dim remTargetPath As String: remTargetPath = remSegments(1)

                Dim subDomainDict As Object: Set subDomainDict = pathLookup(remainingKey)
                Dim dKey As Variant, remAlreadyHandled As Boolean: remAlreadyHandled = True

                For Each dKey In subDomainDict.Keys()
                    If Not domainsInThisBatch.Exists(CStr(dKey)) Then
                        remAlreadyHandled = False
                        Exit For
                    End If
                Next dKey

                If Not remAlreadyHandled Then
                    Dim remResolvedStoreName As String
                    remResolvedStoreName = ResolveArchiveStoreForAccount(remAccountBinding, archiveMappings, archiveMappingCount, masterArchiveName)
                    Call BuildRulesFromMap(defaultStoreRules, pathLookup, remainingKey, remTargetPath, remAccountBinding, Nothing, domainsInThisBatch, ruleActionCounter, processedCount, ns, remResolvedStoreName, harvestedDuplicateCount)
                    keysProcessedThisBatch(CStr(remainingKey)) = True
                End If
            End If
            If ruleActionCounter >= 1000 Then Exit For
        Next remIdx
    End If

    Dim finalBatch As Object: Set finalBatch = CreateObject("Scripting.Dictionary")
    For Each k In domainsInThisBatch.Keys()
        finalBatch(CStr(k)) = True
    Next k

    Dim resortAcc As Outlook.Account
    Dim resortAcctCount As Long: resortAcctCount = 0
    On Error Resume Next: resortAcctCount = Application.Session.Accounts.Count: On Error GoTo 0
    Dim resortAcctIdx As Long
    For resortAcctIdx = 1 To resortAcctCount
        Set resortAcc = Nothing
        On Error Resume Next: Set resortAcc = Application.Session.Accounts.item(resortAcctIdx): On Error GoTo 0
        If resortAcc Is Nothing Then GoTo NextAccountResort

        Dim resortAccRules As Outlook.Rules: Set resortAccRules = Nothing
        On Error Resume Next
        If Not resortAcc.DeliveryStore Is Nothing Then
            If Not resortAcc.DeliveryStore.GetRootFolder() Is Nothing Then
                Set resortAccRules = resortAcc.DeliveryStore.GetRules()
            End If
        End If
        On Error GoTo 0

        If resortAccRules Is Nothing Then
            Dim thisSortAcctDisplay As String: thisSortAcctDisplay = ""
            On Error Resume Next: thisSortAcctDisplay = LCase(resortAcc.DisplayName): On Error GoTo 0
            Dim thisSortAcctSmtpForStores As String: thisSortAcctSmtpForStores = ""
            On Error Resume Next: thisSortAcctSmtpForStores = LCase(resortAcc.SmtpAddress): On Error GoTo 0

            Dim sortCandidateStore As Outlook.Store, sortCandidateDisplay As String, sortIsMatch As Boolean
            For Each sortCandidateStore In Application.Session.Stores
                sortCandidateDisplay = ""
                On Error Resume Next: sortCandidateDisplay = LCase(sortCandidateStore.DisplayName): On Error GoTo 0
                If Len(sortCandidateDisplay) > 0 Then
                    sortIsMatch = False
                    If Len(thisSortAcctDisplay) > 0 And sortCandidateDisplay = thisSortAcctDisplay Then sortIsMatch = True
                    If Not sortIsMatch And Len(thisSortAcctSmtpForStores) > 0 And sortCandidateDisplay = thisSortAcctSmtpForStores Then sortIsMatch = True

                    If sortIsMatch Then
                        On Error Resume Next: Set resortAccRules = sortCandidateStore.GetRules(): On Error GoTo 0
                        If Not resortAccRules Is Nothing Then Exit For
                    End If
                End If
            Next sortCandidateStore
        End If

        If Not resortAccRules Is Nothing Then
            Call ResortRulesByLabel(resortAccRules)
        End If
NextAccountResort:
    Next resortAcctIdx

    Dim resortDefaultRules As Outlook.Rules: Set resortDefaultRules = Nothing
    On Error Resume Next
    Set resortDefaultRules = Application.Session.DefaultStore.GetRules()
    On Error GoTo 0
    If Not resortDefaultRules Is Nothing Then
        Call ResortRulesByLabel(resortDefaultRules)
    End If

    Call WriteTimestampsToCSV(csvPath, allLines, finalBatch, timestampString, keysProcessedThisBatch)
    
    Dim summaryMsg As String
    summaryMsg = "Full production run complete!" & vbCrLf & vbCrLf & _
                 processedCount & " sender domain(s) processed and timestamped." & vbCrLf & _
                 ruleActionCounter & " new consolidated rule(s) created."
    If harvestedDuplicateCount > 0 Then
        summaryMsg = summaryMsg & vbCrLf & harvestedDuplicateCount & _
                     " duplicate/prior rule(s) found and removed during consolidation."
    End If
    MsgBox summaryMsg, vbInformation, "Success"
End Sub

Private Sub BuildRulesFromMap(ByVal targetRules As Outlook.Rules, ByVal pathLookup As Object, ByVal mapKey As Variant, ByVal targetPath As String, ByVal ruleAccountBinding As String, ByVal activeAccount As Outlook.Account, ByRef domainsInThisBatch As Object, ByRef ruleActionCounter As Long, ByRef processedCount As Long, ByVal ns As Outlook.NameSpace, ByVal resolvedTargetStoreName As String, ByRef harvestedDuplicateCount As Long)
    Dim subDomainDict As Object: Set subDomainDict = pathLookup(mapKey)
    Dim domainList() As Variant: domainList = subDomainDict.Keys()
    Dim dCount As Long: dCount = UBound(domainList) + 1
    
    Dim manualSenders() As String
    Dim manualCount As Long: manualCount = 0
    ReDim manualSenders(0)
    
    Dim hasPreservedSubject As Boolean: hasPreservedSubject = False
    Dim preservedSubjectText As Variant
    Dim hasPreservedBody As Boolean: hasPreservedBody = False
    Dim preservedBodyText As Variant
    Dim hasPreservedBodyOrSubject As Boolean: hasPreservedBodyOrSubject = False
    Dim preservedBodyOrSubjectText As Variant
    Dim hasPreservedHeader As Boolean: hasPreservedHeader = False
    Dim preservedHeaderText As Variant
    Dim hasPreservedRecipient As Boolean: hasPreservedRecipient = False
    Dim preservedRecipientText As Variant
    Dim hasPreservedCategory As Boolean: hasPreservedCategory = False
    Dim preservedCategoryText As Variant
    Dim hasPreservedImportance As Boolean: hasPreservedImportance = False
    Dim preservedImportanceVal As Long
    Dim hasPreservedSensitivity As Boolean: hasPreservedSensitivity = False
    Dim preservedSensitivityVal As Long
    Dim hasPreservedAttachment As Boolean: hasPreservedAttachment = False
    Dim hasPreservedCc As Boolean: hasPreservedCc = False
    Dim hasPreservedOnlyToMe As Boolean: hasPreservedOnlyToMe = False
    Dim hasPreservedToOrCc As Boolean: hasPreservedToOrCc = False

    If dCount > 0 And CStr(domainList(0)) <> "" Then
        Dim dCounter As Long: dCounter = 0: Dim partCounter As Long: partCounter = 1
        Dim pathSegments() As String: pathSegments = Split(targetPath, "\")
        Dim baseFolderName As String: baseFolderName = Trim(pathSegments(UBound(pathSegments)))
        
        Dim sldRegex As Object: Set sldRegex = CreateObject("VBScript.RegExp")
        sldRegex.IgnoreCase = True
        sldRegex.Pattern = "^([a-zA-Z0-9_-]+)(?:\.[a-zA-Z]{2,})+$"

        On Error Resume Next
        Dim folderIdx As Long
        For folderIdx = targetRules.Count To 1 Step -1
            Dim candidateRule As Outlook.Rule: Set candidateRule = targetRules.item(folderIdx)
            If Not candidateRule Is Nothing Then
                Dim candidateFolderPath As String: candidateFolderPath = ""
                If candidateRule.Actions.MoveToFolder.Enabled Then
                    If Not candidateRule.Actions.MoveToFolder.folder Is Nothing Then
                        candidateFolderPath = GetFolderFullPathVBA(candidateRule.Actions.MoveToFolder.folder)
                        Dim storeCheck As Outlook.Store
                        For Each storeCheck In Application.Session.Stores
                            Dim storePrefixCandidate As String: storePrefixCandidate = storeCheck.DisplayName & "\"
                            If Left(candidateFolderPath, Len(storePrefixCandidate)) = storePrefixCandidate Then
                                candidateFolderPath = Mid(candidateFolderPath, Len(storePrefixCandidate) + 1)
                                Exit For
                            End If
                        Next storeCheck
                    End If
                End If
                
                If candidateFolderPath <> "" And candidateFolderPath = targetPath Then
                    If candidateRule.Conditions.senderAddress.Enabled Then
                        Dim liveAddresses As Variant: liveAddresses = candidateRule.Conditions.senderAddress.Address
                        Dim laIdx As Long
                        For laIdx = LBound(liveAddresses) To UBound(liveAddresses)
                            Dim liveItem As String: liveItem = LCase(Trim(liveAddresses(laIdx)))
                            If liveItem <> "" Then
                                Dim keepAsManual As Boolean: keepAsManual = True
                                
                                Dim csvIdx As Long
                                For csvIdx = 0 To UBound(domainList)
                                    Dim csvDomain As String: csvDomain = LCase(Trim(domainList(csvIdx)))
                                    
                                    If liveItem = csvDomain Then
                                        keepAsManual = False
                                        Exit For
                                    ElseIf InStr(liveItem, ".") = 0 Then
                                        Dim targetSLD As String: targetSLD = csvDomain
                                        If sldRegex.Test(csvDomain) Then
                                            targetSLD = sldRegex.Execute(csvDomain)(0).SubMatches(0)
                                        End If
                                        
                                        If liveItem = targetSLD Then
                                            keepAsManual = False
                                            Exit For
                                        End If
                                    End If
                                Next csvIdx
                                
                                If keepAsManual Then
                                    Dim checkDup As Long, isAlreadySaved As Boolean: isAlreadySaved = False
                                    For checkDup = 0 To manualCount - 1
                                        If manualSenders(checkDup) = liveItem Then isAlreadySaved = True: Exit For
                                    Next checkDup
                                    
                                    If Not isAlreadySaved Then
                                        ReDim Preserve manualSenders(manualCount)
                                        manualSenders(manualCount) = liveItem
                                        manualCount = manualCount + 1
                                    End If
                               End If
                            End If
                        Next laIdx
                    End If
                    
                    If candidateRule.Conditions.Subject.Enabled Then
                        hasPreservedSubject = True
                        preservedSubjectText = candidateRule.Conditions.Subject.text
                    End If
                    If candidateRule.Conditions.Body.Enabled Then
                        hasPreservedBody = True
                        preservedBodyText = candidateRule.Conditions.Body.text
                    End If
                    If candidateRule.Conditions.BodyOrSubject.Enabled Then
                        hasPreservedBodyOrSubject = True
                        preservedBodyOrSubjectText = candidateRule.Conditions.BodyOrSubject.text
                    End If
                    If candidateRule.Conditions.MessageHeader.Enabled Then
                        hasPreservedHeader = True
                        preservedHeaderText = candidateRule.Conditions.MessageHeader.text
                    End If
                    If candidateRule.Conditions.RecipientAddress.Enabled Then
                        hasPreservedRecipient = True
                        preservedRecipientText = candidateRule.Conditions.RecipientAddress.Address
                    End If
                    If candidateRule.Conditions.Category.Enabled Then
                        hasPreservedCategory = True
                        preservedCategoryText = candidateRule.Conditions.Category.Categories
                    End If
                    If candidateRule.Conditions.Importance.Enabled Then
                        hasPreservedImportance = True
                        preservedImportanceVal = candidateRule.Conditions.Importance.Importance
                    End If
                    If candidateRule.Conditions.Sensitivity.Enabled Then
                        hasPreservedSensitivity = True
                        preservedSensitivityVal = candidateRule.Conditions.Sensitivity.Sensitivity
                    End If
                    If candidateRule.Conditions.HasAttachment.Enabled Then
                        hasPreservedAttachment = True
                    End If
                    If candidateRule.Conditions.Cc.Enabled Then
                        hasPreservedCc = True
                    End If
                    If candidateRule.Conditions.OnlyToMe.Enabled Then
                        hasPreservedOnlyToMe = True
                    End If
                    If candidateRule.Conditions.ToOrCc.Enabled Then
                        hasPreservedToOrCc = True
                    End If
                    
                    harvestedDuplicateCount = harvestedDuplicateCount + 1
                    targetRules.Remove folderIdx
                End If
            End If
        Next folderIdx
        On Error GoTo 0
        
        domainList = subDomainDict.Keys()
        dCount = UBound(domainList) + 1
        
        Do While dCounter < dCount
            If ruleActionCounter >= 1000 Then Exit Do
            
            Dim chunkName As String
            chunkName = "Rule: [" & ruleAccountBinding & "] " & baseFolderName & " (Part " & partCounter & ")"
            
            Dim itemsInChunk As Long: itemsInChunk = 5
            If (dCount - dCounter) < 5 Then itemsInChunk = dCount - dCounter
            
            Dim totalItems As Long: totalItems = itemsInChunk
            If partCounter = 1 Then totalItems = itemsInChunk + manualCount
            
            Dim chunkArray() As String: ReDim chunkArray(totalItems - 1)
            Dim cIdx As Long
            For cIdx = 0 To itemsInChunk - 1
                chunkArray(cIdx) = CStr(domainList(dCounter))
                domainsInThisBatch(CStr(domainList(dCounter))) = True
                dCounter = dCounter + 1
            Next cIdx
            
            If partCounter = 1 And manualCount > 0 Then
                Dim mIdx As Long
                For mIdx = 0 To manualCount - 1
                    chunkArray(itemsInChunk + mIdx) = manualSenders(mIdx)
                Next mIdx
            End If
            
            Dim consolidatedRule As Outlook.Rule: Set consolidatedRule = targetRules.Create(chunkName, olRuleReceive)
            
            consolidatedRule.Actions.Stop.Enabled = True
            consolidatedRule.Conditions.OnLocalMachine.Enabled = True
            
            With consolidatedRule.Conditions.senderAddress
                .Enabled = True
                .Address = chunkArray
            End With
            
            If hasPreservedSubject Then
                With consolidatedRule.Conditions.Subject
                    .Enabled = True
                    .text = preservedSubjectText
                End With
            End If
            If hasPreservedBody Then
                With consolidatedRule.Conditions.Body
                    .Enabled = True
                    .text = preservedBodyText
                End With
            End If
            If hasPreservedBodyOrSubject Then
                With consolidatedRule.Conditions.BodyOrSubject
                    .Enabled = True
                    .text = preservedBodyOrSubjectText
                End With
            End If
            If hasPreservedHeader Then
                With consolidatedRule.Conditions.MessageHeader
                    .Enabled = True
                    .text = preservedHeaderText
                End With
            End If
            If hasPreservedRecipient Then
                With consolidatedRule.Conditions.RecipientAddress
                    .Enabled = True
                    .Address = preservedRecipientText
                End With
            End If
            If hasPreservedCategory Then
                With consolidatedRule.Conditions.Category
                    .Enabled = True
                    .Categories = preservedCategoryText
                End With
            End If
            If hasPreservedImportance Then
                With consolidatedRule.Conditions.Importance
                    .Enabled = True
                    .Importance = preservedImportanceVal
                End With
            End If
            If hasPreservedSensitivity Then
                With consolidatedRule.Conditions.Sensitivity
                    .Enabled = True
                    .Sensitivity = preservedSensitivityVal
                End With
            End If
            If hasPreservedAttachment Then
                consolidatedRule.Conditions.HasAttachment.Enabled = True
            End If
            If hasPreservedCc Then
                consolidatedRule.Conditions.Cc.Enabled = True
            End If
            If hasPreservedOnlyToMe Then
                consolidatedRule.Conditions.OnlyToMe.Enabled = True
            End If
            If hasPreservedToOrCc Then
                consolidatedRule.Conditions.ToOrCc.Enabled = True
            End If
            
            Dim targetArchiveFolderObj As Outlook.folder
            Dim resolvedArchiveRootForThisRule As Outlook.folder
            Set resolvedArchiveRootForThisRule = GetResolvedArchiveRoot(ns, resolvedTargetStoreName)
            If resolvedArchiveRootForThisRule Is Nothing Then
                Debug.Print "BuildRulesFromMap: could not resolve archive store '" & resolvedTargetStoreName & "' for '" & chunkName & "' -- skipping MoveToFolder action for this rule."
            Else
                Set targetArchiveFolderObj = GetOrCreateArchiveFolder(resolvedArchiveRootForThisRule, targetPath)
            End If
            
            If Not targetArchiveFolderObj Is Nothing Then
                On Error Resume Next
                Dim moveAction As Outlook.MoveOrCopyRuleAction: Set moveAction = consolidatedRule.Actions.MoveToFolder
                VBA.Interaction.CallByName moveAction, "Folder", VbSet, targetArchiveFolderObj
                moveAction.Enabled = True
                On Error GoTo 0
            End If
            
            On Error Resume Next
            VBA.Interaction.CallByName targetRules, "Save", VbMethod
            On Error GoTo 0
            
            DoEvents
            
            ruleActionCounter = ruleActionCounter + 1
            processedCount = processedCount + itemsInChunk
            partCounter = partCounter + 1
        Loop
    End If
End Sub

Sub WriteTimestampsToCSV(ByVal path As String, ByRef allLines As Variant, ByRef targetDomains As Object, ByVal tStamp As String, ByRef keysProcessedThisBatch As Object)
    Dim outAdoStream As Object: Set outAdoStream = CreateObject("ADODB.Stream")
    outAdoStream.Type = 2
    outAdoStream.Charset = "utf-8"
    outAdoStream.Open
    outAdoStream.WriteText allLines(0), 1
    Dim i As Long
    For i = LBound(allLines) + 1 To UBound(allLines)
        Dim currentLine As String: currentLine = allLines(i)
        If Trim(currentLine) <> "" Then
            Dim fields() As String: fields = ParseCSVLine(currentLine)
            
            If UBound(fields) >= 6 Then
                Dim rowCompositeKey As String: rowCompositeKey = LCase(Trim(fields(0))) & "|" & fields(5)
                Dim rowDomainMatched As Boolean: rowDomainMatched = keysProcessedThisBatch.Exists(rowCompositeKey)

                Dim existingTimestamp As String: existingTimestamp = Trim(fields(3))
                If rowDomainMatched And existingTimestamp = "" Then
                    fields(3) = tStamp
                End If

                ' Re-serialize all lines to guarantee robust escaping rules apply everywhere
                Dim fIdx As Long
                For fIdx = LBound(fields) To UBound(fields)
                    Dim val As String: val = fields(fIdx)
                    If InStr(val, ",") > 0 Or InStr(val, """") > 0 Or InStr(val, vbCr) > 0 Or InStr(val, vbLf) > 0 Then
                        val = """" & Replace(val, """", """""") & """"
                    End If
                    fields(fIdx) = val
                Next fIdx
                currentLine = Join(fields, ",")
            End If
            outAdoStream.WriteText currentLine, 1
        End If
    Next i
    outAdoStream.SaveToFile path, 2
    outAdoStream.Close
End Sub

Function GetOrCreateArchiveFolder(ByVal rootFolder As Object, ByVal relativePath As String) As Outlook.folder
    Dim pathSegments() As String: pathSegments = Split(relativePath, "\")
    Dim currentFolderObj As Outlook.folder: Set currentFolderObj = rootFolder
    Dim segmentIndex As Long
    For segmentIndex = 0 To UBound(pathSegments)
        Dim nextFolderName As String: nextFolderName = Trim(pathSegments(segmentIndex))
        If nextFolderName <> "" Then
            Dim nextFolderObj As Outlook.folder: Set nextFolderObj = Nothing
            On Error Resume Next
            Set nextFolderObj = currentFolderObj.Folders.item(nextFolderName)
            On Error GoTo 0
            If nextFolderObj Is Nothing Then
                On Error Resume Next
                Set nextFolderObj = currentFolderObj.Folders.Add(nextFolderName)
                On Error GoTo 0
            End If
            Set currentFolderObj = nextFolderObj
            If currentFolderObj Is Nothing Then Exit For
        End If
    Next segmentIndex
    Set GetOrCreateArchiveFolder = currentFolderObj
End Function

Function GetFolderFullPathVBA(ByVal targetFolder As Outlook.folder) As String
    On Error GoTo ErrHandler
    Dim pathParts As String: pathParts = targetFolder.Name
    Dim walker As Object: Set walker = targetFolder.Parent
    Do While Not walker Is Nothing
        If TypeName(walker) = "Folder" Or TypeName(walker) = "MAPIFolder" Then
            pathParts = walker.Name & "\" & pathParts
            On Error Resume Next
            Dim nextWalker As Object: Set nextWalker = Nothing
            Set nextWalker = walker.Parent
            On Error GoTo 0
            Set walker = nextWalker
        Else
            Exit Do
        End If
    Loop
    GetFolderFullPathVBA = pathParts
    Exit Function
ErrHandler:
    GetFolderFullPathVBA = ""
End Function