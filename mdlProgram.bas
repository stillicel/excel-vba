Attribute VB_Name = "mdlProgram"
Public Sub Main()

    Dim objRs As Recordset
    Set objRs = mdlOfflineRecordSet.CreateRecordsetFromDB()
    Call TestOfflineRecordsetForReading(objRs)
    Call TestOfflineRecordsetForWriting(objRs)
    Set objRs = Nothing

End Sub

Private Sub TestOfflineRecordsetForWriting(ByRef RecordsetObj As ADODB.Recordset)
    Dim objRs As Recordset
    Set objRs = RecordsetObj
    Set RecordsetObj = Nothing
    
    objRs.Save Application.ActiveWorkbook.Path & "\" & "original.xml", adPersistXML
    
    'Update the Contact_ID filed
    objRs.MoveFirst
    Dim strID As String
    Do Until objRs.EOF
        strID = objRs.Fields("Contact_ID").Value
        strID = "C_" & strID
        objRs.Fields.Item("Contact_ID").Value = strID
        objRs.Update
        objRs.MoveNext
    Loop
    objRs.Save Application.ActiveWorkbook.Path & "\" & "update.xml", adPersistXML
    
    'Delete the first record
    objRs.MoveFirst
    objRs.Delete
    objRs.Save Application.ActiveWorkbook.Path & "\" & "delete.xml", adPersistXML
    
    'Add a new record
    objRs.AddNew Array("Contact_ID", "Contact_Name", "Tel_Office", "Tel_Home", "Tel_Mobile", "Live_ID", "Skype_ID", "GMail"), _
                 Array("C_999999", "NewGuy", "021-99999999", "021-88888888", "13988888888", "NewGuy@live.com", "NewGuy", "New_Guy@gmail.com")
    objRs.Save Application.ActiveWorkbook.Path & "\" & "add.xml", adPersistXML
    
    Set objRs = Nothing
End Sub

Private Sub TestOfflineRecordsetForReading(ByRef RecordsetObj As ADODB.Recordset)
    Dim objRs As Recordset
    Set objRs = RecordsetObj
    Set RecordsetObj = Nothing
    
    'As a result of Client-side cursor,
    'we can invoke RecordCount properties
    Debug.Print "RecordCount = " & objRs.RecordCount
    
    'Print all the records
    objRs.MoveFirst
    Dim strVal As String
    Do Until objRs.EOF
        strVal = ""
        strVal = strVal & "Contact_ID = " & objRs.Fields("Contact_ID").Value & vbCrLf
        strVal = strVal & "Contact_Name = " & objRs.Fields("Contact_Name").Value & vbCrLf
        strVal = strVal & "Tel_Office = " & objRs.Fields("Tel_Office").Value & vbCrLf
        strVal = strVal & "GMail = " & objRs.Fields("GMail").Value & vbCrLf
        Debug.Print "TestOfflineRecordsetForReading(1st time)" & vbCrLf & strVal
        
        objRs.MoveNext
    Loop
    
    'Print all the records again
    'If the CursorLocation is adUseServer,
    'it is no way to move back
    objRs.MoveFirst
    Do Until objRs.EOF
        strVal = ""
        strVal = strVal & "Contact_ID = " & objRs.Fields("Contact_ID").Value & vbCrLf
        strVal = strVal & "Contact_Name = " & objRs.Fields("Contact_Name").Value & vbCrLf
        strVal = strVal & "Tel_Office = " & objRs.Fields("Tel_Office").Value & vbCrLf
        strVal = strVal & "GMail = " & objRs.Fields("GMail").Value & vbCrLf
        Debug.Print "TestOfflineRecordsetForReading(2nd time)" & vbCrLf & strVal
        
        objRs.MoveNext
    Loop
    
    'Save the records into a XML file
    Dim strXML As String
    strXML = Application.ActiveWorkbook.Path & "\Contact.xml"
    Call objRs.Save(strXML, adPersistXML)
    
    Set objRs = Nothing
End Sub
