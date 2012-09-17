Attribute VB_Name = "mdlOfflineRecordSet"
'Create an offline Recordset object from MS Access Database file
Public Function CreateRecordsetFromDB() As Recordset
    Dim strDB As String
    strDB = Application.ActiveWorkbook.Path & "\test.accdb"
    
    'Get the connection string for a MS Access 2007 database file
    Dim strCnn As String
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""@DB_File"";Persist Security Info=False"
    strCnn = Replace(strCnn, "@DB_File", strDB)

    'Connect the database file
    Dim objCnn As Connection
    Set objCnn = New Connection
    objCnn.Open strCnn
    
    'Get the SQL statement
    'For example, there is a TBL_Contact table in that database
    Dim strSQL As String
    strSQL = "SELECT * FROM TBL_Contact"
    
    'To get an offline Recordset object,
    'you should resolve the following operations
    Dim objRs As Recordset
    Set objRs = New Recordset
    objRs.CursorLocation = adUseClient
    objRs.CursorType = adOpenStatic
    objRs.LockType = adLockOptimistic
    
    'Disconnect the database after getting the data from database
    Set objRs.ActiveConnection = objCnn
    objRs.Open strSQL
    Set objRs.ActiveConnection = Nothing
    objCnn.Close
    Set objCnn = Nothing
    
    'It will return an offline Recordset object
    Set CreateRecordsetFromDB = objRs

End Function

Public Function CreateRecordsetFromCSV(ByVal CSVFileDir As String, ByVal CSVFileName As String) As Recordset
    Dim strCnn As String
    strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CSVFileDir & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"""

    Dim strSQL As String
    strSQL = "SELECT * FROM " & CSVFileName
    
    Dim objCnn As Connection
    Set objCnn = New Connection
    objCnn.Open strCnn
    
    Dim objRs As Recordset
    Set objRs = New Recordset
    objRs.CursorLocation = adUseClient
    objRs.CursorType = adOpenStatic
    objRs.LockType = adLockOptimistic
    
    Set objRs.ActiveConnection = objCnn
    objRs.Open strSQL
    
    Set objRs.ActiveConnection = Nothing
    
    objCnn.Close
    Set objCnn = Nothing

    Set CreateRecordsetFromCSV = objRs
End Function

Public Function CreateRecordsetFromXML(ByVal XMLFilePath As String) As Recordset

    Dim objRs As Recordset
    Set objRs = New Recordset
    objRs.Open XMLFilePath
    
    Set CreateRecordsetFromXML = objRs

End Function

Public Function CreateRecordsetFromMemory() As Recordset
    Dim objRs As Recordset
    Set objRs = New Recordset
    
    objRs.CursorLocation = adUseClient
    objRs.CursorType = adOpenKeyset
    objRs.LockType = adLockBatchOptimistic
    
    objRs.Fields.Append "A", adChar, 50
    objRs.Fields.Append "B", adChar, 50
    objRs.Fields.Append "C", adChar, 50
    objRs.Fields.Append "D", adChar, 50
    
    objRs.Open
    
    objRs.AddNew Array("A", "B", "C", "D"), _
                 Array("aaaaaaaaaaaaaaaa1", "bbbbbbbbbbbb1", "cccccccccc1", "ddddddddd1")

    objRs.AddNew Array("A", "B", "C", "D"), _
                 Array("aaaaaaaaaaaaaaaa2", "bbbbbbbbbbbb2", "cccccccccc2", "ddddddddd2")
                 
    objRs.AddNew Array("A", "B", "C", "D"), _
                 Array("aaaaaaaaaaaaaaaa3", "bbbbbbbbbbbb3", "cccccccccc3", "ddddddddd3")

    objRs.AddNew Array("A", "B", "C", "D"), _
                 Array("aaaaaaaaaaaaaaaa4", "bbbbbbbbbbbb4", "cccccccccc4", "ddddddddd4")

    objRs.AddNew Array("A", "B", "C", "D"), _
                 Array("aaaaaaaaaaaaaaaa5", "bbbbbbbbbbbb5", "cccccccccc5", "ddddddddd5")

    objRs.AddNew Array("A", "B", "C", "D"), _
                 Array("aaaaaaaaaaaaaaaa6", "bbbbbbbbbbbb6", "cccccccccc6", "ddddddddd6")
                 
    Set CreateRecordsetFromMemory = objRs
    
End Function


