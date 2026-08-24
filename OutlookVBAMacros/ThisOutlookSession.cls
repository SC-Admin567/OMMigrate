VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
' OutlookMailMigrator (OMMigrate)
'-------------------------------------------------------------------------
' Originator & Architect:    Kirk Shallcross - Shallcross Consulting
' Implementation Specialist: Anthropic Claude AI
' Inception Date:            May 2026
' Version:                   1.5.2
'-------------------------------------------------------------------------
Public Sub SearchAllUnreadMail()
    Dim objSearch As String
    Dim olSearchScopeAllMailBoxes As Integer
    
    ' Assign the explicit code for searching all mailboxes
    olSearchScopeAllMailBoxes = 2
    objSearch = "isread:no"
    
    On Error Resume Next
    Dim currentExplorer As Outlook.Explorer
    Set currentExplorer = Application.ActiveExplorer
    
    If Not currentExplorer Is Nothing Then
        currentExplorer.Search objSearch, olSearchScopeAllMailBoxes
    End If
End Sub
Public Sub ForceAllMailboxSearch()
    Dim txtSearch As String
    Dim myExplorer As Outlook.Explorer
    
    ' Set search query text
    txtSearch = "isread:no"
    
    Set myExplorer = Application.ActiveExplorer
    If Not myExplorer Is Nothing Then
        ' Explicitly clear any existing folder limits by resetting view
        myExplorer.ClearSearch
        ' Execute search forcing the absolute highest mailbox scope index
        myExplorer.Search txtSearch, 2
    End If
End Sub
