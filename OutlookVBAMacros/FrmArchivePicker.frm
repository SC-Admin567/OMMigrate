VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmArchivePicker 
   Caption         =   "Select your Master Archive to Redistribute Email"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmArchivePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmArchivePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedStoreName As String
Public WasCancelled As Boolean

Public Sub PopulateStores(ByVal storeNames As Collection)
    Dim nm As Variant
    Me.lstStores.Clear
    For Each nm In storeNames
        Me.lstStores.AddItem CStr(nm)
    Next nm
    If Me.lstStores.ListCount > 0 Then Me.lstStores.ListIndex = 0
End Sub

Private Sub cmdOK_Click()
    If Me.lstStores.ListIndex = -1 Then
        MsgBox "Please select a store from the list.", vbExclamation, "No Selection"
        Exit Sub
    End If
    SelectedStoreName = Me.lstStores.value
    WasCancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    SelectedStoreName = ""
    WasCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    WasCancelled = True  ' default to cancelled unless OK is explicitly clicked
End Sub
