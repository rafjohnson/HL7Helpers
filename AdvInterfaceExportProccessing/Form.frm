VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "HL7 Search"
   ClientHeight    =   5664
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9624.001
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub btnSearchFields_Click()
    Dim results As String
    results = ThisWorkbook.getSheetValuesAtSegmentAndIndex(cmbSeg2.value, tbField2.Text)
    tbResults.Text = results
End Sub

Private Sub btnSearchMsgs_Click()
    'search for messages.
  Call ThisWorkbook.searchMessages(cmbSeg1.value, tbField1.Text, tbSearchVal.Text)
    
    
End Sub

Private Sub MultiPage1_Change()
    'load segments
    'get panel
    Debug.Print (MultiPage1.SelectedItem.Caption)
End Sub

Private Sub tbField1_Change()

End Sub

Private Sub tbField2_Change()

End Sub

Private Sub tbSearchVal_Change()

End Sub

Private Sub UserForm_Initialize()
    'load segment?
    Debug.Print ("Load Form")
    
    'load segments
    Dim segments() As String
    segments = ThisWorkbook.getSegmentNames()
    
    Dim segNameFromList As Variant
    For Each segNameFromList In segments
        cmbSeg1.AddItem (segNameFromList)
        cmbSeg2.AddItem (segNameFromList)
    Next segNameFromList
    
End Sub
