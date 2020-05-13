VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilSel 
   Caption         =   "DemSheets: File Selection"
   ClientHeight    =   2532
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5772
   OleObjectBlob   =   "frmFilSel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFilSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctlSelDir_Click()
gFilSelDir = True
Unload Me
End Sub

Private Sub ctlSelFil_Click()
gFilSelDir = False
Unload Me
End Sub
