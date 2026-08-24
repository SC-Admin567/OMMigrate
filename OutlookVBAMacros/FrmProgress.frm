VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProgress 
   Caption         =   "OMMigrate Engine Progress"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "FrmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The maximum width the progress bar is allowed to expand to inside the form frame
Private Const MAX_BAR_WIDTH As Long = 220

Public Sub UpdateProgress(ByVal PhaseText As String, ByVal Percent As Single)
    Dim targetWidth As Single
    
    ' 1. Calculate safe, bounded control width
    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100
    
    targetWidth = (Percent / 100) * MAX_BAR_WIDTH
    If targetWidth < 5 Then targetWidth = 5
    
    ' 2. Safely modify control layout properties
    Me.lblStatus.Caption = PhaseText
    Me.lblBar.Width = targetWidth
    
    ' 3. Hard-paint the UI frame update to prevent Windows thread caching
    Me.Repaint
    DoEvents
End Sub

