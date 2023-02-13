VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Imports 
   Caption         =   "Import from another setup"
   ClientHeight    =   7440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "Imports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Imports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub ImportButton_Click()
    ImportSetup
End Sub

Private Sub LoadButton_Click()
    NewSetupPath
End Sub
