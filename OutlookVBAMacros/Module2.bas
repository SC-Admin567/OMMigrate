Attribute VB_Name = "Module2"
'-------------------------------------------------------------------------
' OutlookMailMigrator (OMMigrate)
'-------------------------------------------------------------------------
' Originator & Architect:    Kirk Shallcross - Shallcross Consulting
' Implementation Specialist: Anthropic Claude AI
' Inception Date:            May 2026
' Version:                   1.5.2
'-------------------------------------------------------------------------
Option Explicit

'===========================================================
' Main Entry Point: Empties all Deleted Items/Trash folders
'===========================================================
Sub EmptyAllTrashFolders()
    Dim objNamespace As Outlook.NameSpace
    Dim objStores As Outlook.Stores
    Dim objStore As Outlook.Store
    Dim objFolder As Outlook.folder
    Dim totalStoresChecked As Long
    Dim totalFoldersEmptied As Long
    
    ' Static counter retains its value across automatic sequential executions
    Static executionPass As Long
    
    ' Increment the pass tracking counter
    executionPass = executionPass + 1
    
    Set objNamespace = Application.GetNamespace("MAPI")
    Set objStores = objNamespace.Stores
    
    totalStoresChecked = 0
    totalFoldersEmptied = 0
    
    Debug.Print "=== Starting Global Trash Purge (Pass " & executionPass & "): " & Now & " ==="
    
    ' Loop through every email account/store in the active Outlook profile
    For Each objStore In objStores
        totalStoresChecked = totalStoresChecked + 1
        Set objFolder = Nothing
        
        Debug.Print "Checking Store: " & objStore.DisplayName
        
        ' 1. Attempt to resolve via default MAPI folder enumeration
        On Error Resume Next
        Set objFolder = objStore.GetDefaultFolder(olFolderDeletedItems)
        On Error GoTo 0
        
        ' 2. Fallback: Deep recursive search if default resolution returns Nothing
        If objFolder Is Nothing Then
            Dim rootFolder As Outlook.folder
            On Error Resume Next
            Set rootFolder = objStore.GetRootFolder
            On Error GoTo 0
            
            If Not rootFolder Is Nothing Then
                Set objFolder = FindTrashFolderRecursive(rootFolder)
            End If
        End If
       
        ' 3. Process the resolved target trash folder
        If Not objFolder Is Nothing Then
            Debug.Print " -> Found Trash Target: " & objFolder.folderPath
            
            ' Clear all items and nested folder structures recursively
            PurgeFolderContents objFolder
            
            totalFoldersEmptied = totalFoldersEmptied + 1
        Else
            Debug.Print " -> WARNING: No Deleted Items or Trash folder identified for this store."
        End If
        
        DoEvents
    Next objStore
    
    Debug.Print "=== Pass " & executionPass & " Complete ==="
    
    ' Automatically run a second verification pass if this was the first run
    If executionPass < 2 Then
        DoEvents
        Call EmptyAllTrashFolders
    Else
        ' Reset the static counter back to 0 so it's ready for the next time you use it manually
        executionPass = 0
        
        ' Final popup only shows after the second pass completes entirely
        MsgBox "Trash emptying process completed!" & vbCrLf & vbCrLf & _
               "Accounts/Stores processed: " & totalStoresChecked & vbCrLf & _
               "Trash folders successfully cleared: " & totalFoldersEmptied, vbInformation, "Execution Complete"
    End If
End Sub

'===========================================================
' Deep recursive search to find a trash folder anywhere in the tree
'===========================================================
Private Function FindTrashFolderRecursive(ByVal currentFolder As Outlook.folder) As Outlook.folder
    Dim subFolder As Outlook.folder
    Dim foundFolder As Outlook.folder
    
    If InStr(1, currentFolder.Name, "Trash", vbTextCompare) > 0 Or _
       InStr(1, currentFolder.Name, "Deleted Items", vbTextCompare) > 0 Or _
       InStr(1, currentFolder.Name, "Deleted Objects", vbTextCompare) > 0 Then
        Set FindTrashFolderRecursive = currentFolder
        Exit Function
    End If
    
    On Error Resume Next
    If currentFolder.Folders.Count > 0 Then
        For Each subFolder In currentFolder.Folders
            Set foundFolder = FindTrashFolderRecursive(subFolder)
            If Not foundFolder Is Nothing Then
                Set FindTrashFolderRecursive = foundFolder
                Exit Function
            End If
        Next subFolder
    End If
    On Error GoTo 0
    
    Set FindTrashFolderRecursive = Nothing
End Function

'===========================================================
' Purges all items and nested subfolders recursively
'===========================================================
Private Sub PurgeFolderContents(ByVal targetFolder As Outlook.folder)
    Dim folderItems As Outlook.Items
    Dim i As Long
    
    On Error Resume Next
    
    ' Clear items out of this specific level
    Set folderItems = targetFolder.Items
    If Not folderItems Is Nothing Then
        For i = folderItems.Count To 1 Step -1
            folderItems.item(i).Delete
            If i Mod 200 = 0 Then DoEvents
        Next i
    End If
    
    ' Mirror actions down through any nested sub-levels using a snapshot array
    Dim subFolderCount As Long
    subFolderCount = targetFolder.Folders.Count
    
    If subFolderCount > 0 Then
        Dim subFoldersArray() As Outlook.folder
        ReDim subFoldersArray(1 To subFolderCount)
        
        For i = 1 To subFolderCount
            Set subFoldersArray(i) = targetFolder.Folders.item(i)
        Next i
        
        For i = subFolderCount To 1 Step -1
            Dim subFolder As Outlook.folder
            Set subFolder = subFoldersArray(i)
            
            If Not subFolder Is Nothing Then
                PurgeFolderContents subFolder
                DoEvents
                subFolder.Delete
            End If
        Next i
    End If
    On Error GoTo 0
End Sub

