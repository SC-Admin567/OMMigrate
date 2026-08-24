Attribute VB_Name = "Module1"
'-------------------------------------------------------------------------
' OutlookMailMigrator (OMMigrate)
'-------------------------------------------------------------------------
' Originator & Architect:    Kirk Shallcross - Shallcross Consulting
' Implementation Specialist: Anthropic Claude AI
' Inception Date:            May 2026
' Version:                   1.5.2
'-------------------------------------------------------------------------
'===========================================================
' Delete duplicate Outlook items (Mail, Appointments, Tasks, Contacts, etc.)
' Duplicates are moved to a "Duplicates" subfolder inside the starting folder
'===========================================================

Option Explicit

' Global (module-level) variables
Public rootFolderPath As String
Public rootFolder As Outlook.folder
Public duplicateRootFolderPath As String
Public duplicateRootFolder As Outlook.folder
Public gTotalDuplicatesMoved As Long ' Global counter for the final popup

'===========================================================
' Entry point
'===========================================================
Sub DeleteDups()
    Dim folder As Outlook.MAPIFolder
    
    ' Reset global counter
    gTotalDuplicatesMoved = 0
    
    ' Select the starting folder
    Set folder = Application.Session.PickFolder
    
    If Not folder Is Nothing Then
        DeleteDupsInFolder folder, True
    End If
End Sub

' ==============================================================================
' ADDED 2026-07-11, Administrator direction: parameterized variant of DeleteDups that
' takes the starting folder as an argument instead of showing the PickFolder
' dialog. Built so Module7's CorrectArchiveFolders can call this directly
' (already-selected archive root passed in) with zero extra prompts, while
' DeleteDups() above remains completely unchanged for standalone toolbar use
' -- same picker, same behavior, nothing about that workflow is touched.
' ShowSummaryPopup lets the caller suppress the final MsgBox when this is
' being called as one step inside a larger macro (Module7 shows its own
' combined summary at the end instead).
' ==============================================================================
Sub DeleteDupsInFolder(ByVal folder As Outlook.MAPIFolder, Optional ByVal ShowSummaryPopup As Boolean = True)
    ' Reset global counter
    gTotalDuplicatesMoved = 0

    If Not folder Is Nothing Then
        ' Set the root path to be inside the selected folder to guarantee visibility in PSTs
        rootFolderPath = folder.folderPath
        Set rootFolder = folder
        
        ' Prepare Duplicates folder safely inside your selected starting folder
        duplicateRootFolderPath = rootFolderPath & "\Duplicates"
        Set duplicateRootFolder = CreateFolder(duplicateRootFolderPath)
        
        If duplicateRootFolder Is Nothing Then
            If ShowSummaryPopup Then
                MsgBox "Could not create the Duplicates folder. Execution stopped.", vbCritical
            End If
            Exit Sub
        End If
        
        Debug.Print "Started at " & Now
        LoopFolders folder, True
        Debug.Print "Finished at " & Now
        
        ' Final popup with the total duplicate counter -- suppressed when
        ' called from Module7, which reports this count as part of its own
        ' combined summary instead.
        If ShowSummaryPopup Then
            MsgBox "Duplicate removal completed." & vbCrLf & vbCrLf & _
                   "Total duplicate items moved: " & gTotalDuplicatesMoved, vbInformation, "Process Complete"
        End If
    End If
End Sub

'===========================================================
' Recursively process all subfolders
'===========================================================
Sub LoopFolders(currentFolder As Outlook.MAPIFolder, ByVal Recursive As Boolean)
    Dim subFolder As Outlook.MAPIFolder

    ' Skip the duplicates root folder itself to avoid infinite loops
    If currentFolder.folderPath = duplicateRootFolderPath Or _
       InStr(1, currentFolder.folderPath, duplicateRootFolderPath, vbTextCompare) > 0 Then
        Debug.Print "Skipped " & currentFolder.folderPath
        Exit Sub
    End If
    
    ' Process current folder
    DoFolderActions currentFolder

    ' Loop through subfolders if recursive is requested
    If Recursive And (currentFolder.Folders.Count > 0) Then
        For Each subFolder In currentFolder.Folders
            LoopFolders subFolder, True
        Next subFolder
    End If
End Sub

'===========================================================
' Perform deduplication actions for a single folder
'===========================================================
Private Sub DoFolderActions(folder As Outlook.MAPIFolder)
    Dim duplicateTargetFolderPath As String
    Dim duplicateTargetFolder As Outlook.folder
    
    duplicateTargetFolderPath = Replace(folder.folderPath, rootFolderPath, duplicateRootFolderPath)
    Set duplicateTargetFolder = CreateFolder(duplicateTargetFolderPath)
    
    If Not duplicateTargetFolder Is Nothing Then
        RemoveDuplicateItems folder, duplicateTargetFolder
    End If
End Sub

'===========================================================
' Remove duplicates and move them to target folder
'===========================================================
Sub RemoveDuplicateItems(objFolder As Outlook.folder, objTargetFolder As Outlook.folder)
    Dim objDictionary As Object
    Dim folderItems As Outlook.Items
    Dim objItem As Object
    Dim strKey As String
    Dim i As Long
    Dim localDuplicatesDetected As Long
    
    Dim itemsToMove As New Collection
    Dim entryID As String
    
    Set objDictionary = CreateObject("Scripting.Dictionary")
    
    If Not (objFolder Is Nothing) Then
        Set folderItems = objFolder.Items
        
        On Error Resume Next
        folderItems.Sort "[CreationTime]", True
        folderItems.Sort "[ReceivedTime]", True
        On Error GoTo 0
        
        Debug.Print Now & " | Deduplicating: " & objFolder.folderPath
        
        ' --- PASS 1: Identify duplicates ---
        For i = folderItems.Count To 1 Step -1
            Set objItem = folderItems.item(i)
            strKey = ""
            
            Select Case True
                Case TypeOf objItem Is Outlook.mailItem
                    Dim m As Outlook.mailItem
                    Set m = objItem
                    strKey = "MailItem" & m.Subject & m.Body & m.To & m.Cc & m.BCC & m.SenderEmailAddress & m.SentOn
                
                Case TypeOf objItem Is Outlook.AppointmentItem
                    Dim a As Outlook.AppointmentItem
                    Set a = objItem
                    strKey = "AppointmentItem" & a.Subject & a.Start & a.Duration & a.Location & a.Body
                
                Case TypeOf objItem Is Outlook.ContactItem
                    Dim c As Outlook.ContactItem
                    Set c = objItem
                    strKey = "ContactItem" & c.FullName & c.Email1Address & c.Email2Address & c.Email3Address
            
                Case TypeOf objItem Is Outlook.TaskItem
                    Dim t As Outlook.TaskItem
                    Set t = objItem
                    strKey = "TaskItem" & t.Subject & t.StartDate & t.DueDate & t.Body
                
                Case TypeOf objItem Is Outlook.MeetingItem
                    strKey = "MeetingItem" & objItem.Subject & objItem.Body & objItem.SentOn
                
                Case TypeOf objItem Is Outlook.ReportItem
                    strKey = "ReportItem" & objItem.Subject & objItem.Body
            End Select
            
            If Len(strKey) > 0 Then
                strKey = Replace(strKey, ", ", Chr(32))
                
                If objDictionary.Exists(strKey) Then
                    On Error Resume Next
                    entryID = objItem.entryID
                    If Err.Number = 0 And Len(entryID) > 0 Then
                        itemsToMove.Add entryID
                        localDuplicatesDetected = localDuplicatesDetected + 1
                    End If
                    On Error GoTo 0
                Else
                    objDictionary.Add strKey, True
                End If
            End If
            
            DoEvents
        Next i
        
        ' --- PASS 2: Safely execute moves ---
        If itemsToMove.Count > 0 Then
            Dim itemID As Variant
            Dim moveItem As Object
            Dim ns As Outlook.NameSpace
            Set ns = Application.Session
            
            For Each itemID In itemsToMove
                On Error Resume Next
                Set moveItem = ns.GetItemFromID(itemID)
                If Not moveItem Is Nothing Then
                    moveItem.Move objTargetFolder
                    gTotalDuplicatesMoved = gTotalDuplicatesMoved + 1 ' Increment global counter
                End If
                On Error GoTo 0
                DoEvents
            Next itemID
        End If
        
    End If
    
    Debug.Print "Found and moved " & localDuplicatesDetected & " duplicate item(s) from " & objFolder.Name
End Sub

'===========================================================
' Get folder by full path
'===========================================================
Function GetFolder(ByVal folderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
 
    On Error GoTo GetFolder_Error
    If Left(folderPath, 2) = "\\" Then
        folderPath = Right(folderPath, Len(folderPath) - 2)
    End If
    
    FoldersArray = Split(folderPath, "\")
    Set TestFolder = Application.Session.Folders.item(FoldersArray(0))
    
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Set TestFolder = TestFolder.Folders.item(FoldersArray(i))
            If TestFolder Is Nothing Then Exit For
        Next
    End If
     
    Set GetFolder = TestFolder
    Exit Function
 
GetFolder_Error:
    Set GetFolder = Nothing
End Function

'===========================================================
' Create a folder and return it
'===========================================================
Function CreateFolder(ByVal folderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
 
    On Error Resume Next
    If Left(folderPath, 2) = "\\" Then
        folderPath = Right(folderPath, Len(folderPath) - 2)
    End If
    
    FoldersArray = Split(folderPath, "\")
    Set TestFolder = Application.Session.Folders.item(FoldersArray(0))
    
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            
            Dim nextFolder As Outlook.folder
            Set nextFolder = Nothing
            Set nextFolder = SubFolders.item(FoldersArray(i))
            
            If nextFolder Is Nothing Then
                Set TestFolder = SubFolders.Add(FoldersArray(i))
            Else
                Set TestFolder = nextFolder
            End If
        Next
    End If
    
    On Error GoTo 0
    Set CreateFolder = TestFolder
End Function

