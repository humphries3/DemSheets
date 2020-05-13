VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFunSel 
   Caption         =   "DemSheets: Functions"
   ClientHeight    =   2460
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5508
   OleObjectBlob   =   "frmFunSel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFunSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctlCombinFils_Click()
'gFunction = "combsheets"
Call CombSheets
'Unload Me
End Sub

Private Sub ctlDone_Click()
'gFunction = "done"
ThisWorkbook.Application.Quit
'Unload Me
End Sub
