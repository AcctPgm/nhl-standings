VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmStatusMessage 
   Caption         =   "NHL Standings"
   ClientHeight    =   1035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ufmStatusMessage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufmStatusMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ShowMessage(sMsg As String, Optional sTitle As String, Optional bHide As Boolean)
    If bHide Then
        Hide
        Exit Sub
    End If
    
    lblMessage.Caption = sMsg
    If sTitle <> "" Then
        ufmStatusMessage.Caption = sTitle
    End If
    
    Show vbModeless
    DoEvents
End Sub

