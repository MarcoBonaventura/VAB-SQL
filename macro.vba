Dim ws As Object
Dim sheetUtente As Object
Dim sheet_userinfo As Object
Dim sheet_radcheck As Object
Dim sheet_radreply As Object

Dim userRange As Range

Dim valueCol As Integer
Dim xRow As Integer
Dim iRow As Integer
Dim ID As Integer
Dim i As Integer

Dim id_table_user As Long
Dim id_table_check As Long
Dim id_table_reply As Long
Dim k As Long

Dim mycell As String
Dim user As String
Dim passwd As String
Dim radiusIP As String
Dim gruppo As String
Dim saveToPath As String



Sub getUser()
    
    ID = 0
    iRow = 0
    xRow = 0
    id_table = 0
    
    'Set import_sheet = Worksheets("import")
    Set import_sheet = Worksheets("import")
    Set sheetUtente = Worksheets("tbl_utente")
    Set sheet_userinfo = Worksheets("tbl_userinfo_orig")
    Set sheet_radcheck = Worksheets("tbl_radcheck")
    Set sheet_radreply = Worksheets("tbl_radreply")
        
    'percorso dove creare o aggiornare i file .CSV
    saveToPath = "\\x.x.x.x\file_CSV\"
    
    'estraggo MAX ID delle tabelle
     id_table_user = CLng(sheet_userinfo.Range("A1").Value)
     id_table_check = CLng(sheet_radcheck.Range("A1").Value)
     id_table_reply = CLng(sheet_radreply.Range("A1").Value)
    
    'formattazione per nuovi inserimenti
    sheetUtente.Cells.ClearContents
    sheet_userinfo.Cells.ClearContents
    sheet_radcheck.Cells.ClearContents
    sheet_radreply.Cells.ClearContents
             
    'calcolo quande righe ci sono da leggere
    k = import_sheet.Range("A1048576").End(xlUp).Row
    'MsgBox k
       
    
     'MsgBox "id_table_user" & " " & id_table_user
     'MsgBox "id_table_radcheck" & " " & id_table_check
     'MsgBox "id_table_radreply" & " " & id_table_reply
     
     sheetUtente.Cells.ClearContents
     sheet_userinfo.Cells.ClearContents
     sheet_radcheck.Cells.ClearContents
     sheet_radreply.Cells.ClearContents
     
    'ciclo ogni riga per estrarre i valori
    i = 1
    Do While i < k + 1
    
        If i > 1 Then
            mycell = Cells(i, "A").Value
            'MsgBox mycell #per debug
            
            id_table_user = id_table_user + 1
            id_table_check = id_table_check + 1
            
            'estraggo username, password, gruppo e IP
             user = import_sheet.Cells(i, 2).Value
             passwd = import_sheet.Cells(i, 3).Value
             gruppo = import_sheet.Cells(i, 4).Value
             radiusIP = import_sheet.Cells(i, 5).Value
            
            'popolo tabella utente
             Call import_users(iRow, user)
            
            'popolo tabella userinfo_orig
             Call import_userinfo_orig(id_table_user, iRow, user)
            
            'popolo tabella radchek
             Call import_radcheck(id_table_check, iRow, user, passwd)
            
            'popolo tabella radcheck
             Call import_radreply(id_table_reply, xRow, user, radiusIP, gruppo)
            
        End If
        
        i = i + 1
        ID = ID + 1
        iRow = iRow + 1
        xRow = xRow + 1
    
    
    Loop
          
    ' alogoritmo vecchio con bug - rimane solo per documentazione
    'For Each userRange In import_sheet.UsedRange.Columns("A").Cells
    '    If ID > 1 Then
        
            'estraggo lo username
    '        user = import_sheet.Cells(ID, 2).Value
    '        passwd = import_sheet.Cells(ID, 3).Value
    '        radiusIP = import_sheet.Cells(ID, 4).Value
            
            'popolo tabella utente
    '        Call import_users(iRow, user)
            
            'popolo tabella userinfo_orig
    '        Call import_userinfo_orig(iRow, user)
            
            'popolo tabella radchek
    '        Call import_radcheck(iRow, user, passwd)
            
            'popolo tabella radcheck
    '        Call import_radreply(xRow, user, radiusIP)
                    
           
    '    End If
        
    '    ID = ID + 1
    '    iRow = iRow + 1
    '    xRow = xRow + 1

    'Next userRange
            
    MsgBox "importati " & ID & " nominativi"
        
End Sub


Sub import_users(iRow As Integer, user As String)

    sheetUtente.Cells(iRow, 1).Value = user
    sheetUtente.Cells(iRow, 2).Value = "NULL"
    sheetUtente.Cells(iRow, 3).Value = "NULL"
    sheetUtente.Cells(iRow, 4).Value = "0"
    sheetUtente.Cells(iRow, 5).Value = "NULL"
    sheetUtente.Cells(iRow, 6).Value = "0"
    sheetUtente.Cells(iRow, 7).Value = "0"
    sheetUtente.Cells(iRow, 8).Value = "0"
    
End Sub


Sub import_userinfo_orig(id_table_user As Long, iRow As Integer, user As String)

    sheet_userinfo.Cells(iRow, 1).Value = id_table_user
    sheet_userinfo.Cells(iRow, 2).Value = user
    sheet_userinfo.Cells(iRow, 3).Value = "-"
    sheet_userinfo.Cells(iRow, 4).Value = "-"
    sheet_userinfo.Cells(iRow, 5).Value = "-"
    sheet_userinfo.Cells(iRow, 6).Value = "-"
    sheet_userinfo.Cells(iRow, 7).Value = "-"

End Sub


Sub import_radcheck(id_table_check As Long, iRow As Integer, user As String, passwd As String)

    sheet_radcheck.Cells(iRow, 1).Value = id_table_check
    sheet_radcheck.Cells(iRow, 2).Value = user
    sheet_radcheck.Cells(iRow, 3).Value = "User-Password"
    sheet_radcheck.Cells(iRow, 4).Value = ":="
    sheet_radcheck.Cells(iRow, 5).Value = passwd


End Sub


Sub import_radreply(id_table_reply As Long, firstRow As Integer, user As String, radiusIP As String, gruppo As String)

    If firstRow = 1 Then
        secRow = firstRow + 1
    Else
        firstRow = firstRow + 1
        secRow = firstRow + 1
    End If
    
    'inserisco prima riga
    sheet_radreply.Cells(firstRow, 1).Value = id_table_reply + firstRow
    sheet_radreply.Cells(firstRow, 2).Value = user
    sheet_radreply.Cells(firstRow, 3).Value = "Framed-IP-Address"
    sheet_radreply.Cells(firstRow, 4).Value = ":="
    sheet_radreply.Cells(firstRow, 5).Value = radiusIP
    
    'inserisco seconda riga
    sheet_radreply.Cells(secRow, 1).Value = id_table_reply + secRow
    sheet_radreply.Cells(secRow, 2).Value = user
    sheet_radreply.Cells(secRow, 3).Value = "Framed-Filter-ID"
    sheet_radreply.Cells(secRow, 4).Value = ":="
    sheet_radreply.Cells(secRow, 5).Value = gruppo
       
End Sub



Sub WriteCSVFile_utente()

    Dim My_filenumber As Integer
    Dim logSTR As String
    Dim myData As Range
    
    Set sheetUtente = Worksheets("tbl_utente")
    
    My_filenumber = FreeFile
    'Open "\\192.168.102.4\ot-dati\Gestionale\SCHEDE TECNICHE\APN_CONDIVISO\Novareti\import_csv\utente.csv" For Output As #My_filenumber
    Open saveToPath + "utente.csv" For Output As #My_filenumber
    
    rrow = 1
    
    totRows = sheetUtente.UsedRange.Columns("A").Cells.Count
    
    For Each myData In sheetUtente.UsedRange.Columns("A").Cells
    
        logSTR = logSTR & sheetUtente.Cells(rrow, 1).Value & ";"
        logSTR = logSTR & sheetUtente.Cells(rrow, 2).Value & ";"
        logSTR = logSTR & sheetUtente.Cells(rrow, 3).Value & ";"
        logSTR = logSTR & sheetUtente.Cells(rrow, 4).Value & ";"
        logSTR = logSTR & sheetUtente.Cells(rrow, 5).Value & ";"
        logSTR = logSTR & sheetUtente.Cells(rrow, 6).Value & ";"
        logSTR = logSTR & sheetUtente.Cells(rrow, 7).Value & ";"
        
        'per evitare di inserire una nuova linea alla fine della tabella altrimenti SQL inserisce una riga vuota
        If rrow <= totRows - 1 Then
            logSTR = logSTR & sheetUtente.Cells(rrow, 8).Value & vbNewLine
        Else
            logSTR = logSTR & sheetUtente.Cells(rrow, 8).Value
        End If
        
        rrow = rrow + 1
                
    Next
    
    Print #My_filenumber, logSTR
    Close #My_filenumber
        
End Sub



Sub WriteCSVFile_userinfo()

    Dim My_filenumber As Integer
    Dim logSTR As String
    Dim myData As Range
    
    Set sheet_userinfo = Worksheets("tbl_userinfo_orig")
    
    My_filenumber = FreeFile
    Open saveToPath + "userinfo_orig.csv" For Output As #My_filenumber
    
    totRows = sheet_userinfo.UsedRange.Columns("A").Cells.Count
    
    rrow = 1
    
    For Each myData In sheet_userinfo.UsedRange.Columns("A").Cells
    
        logSTR = logSTR & sheet_userinfo.Cells(rrow, 1).Value & ";"
        logSTR = logSTR & sheet_userinfo.Cells(rrow, 2).Value & ";"
        logSTR = logSTR & sheet_userinfo.Cells(rrow, 3).Value & ";"
        logSTR = logSTR & sheet_userinfo.Cells(rrow, 4).Value & ";"
        logSTR = logSTR & sheet_userinfo.Cells(rrow, 5).Value & ";"
        logSTR = logSTR & sheet_userinfo.Cells(rrow, 6).Value & ";"
        
        If rrow <= totRows - 1 Then
            logSTR = logSTR & sheet_userinfo.Cells(rrow, 7).Value & vbNewLine
        Else
            logSTR = logSTR & sheet_userinfo.Cells(rrow, 7).Value
        End If
            
        rrow = rrow + 1
        
    Next
    
    Print #My_filenumber, logSTR
    Close #My_filenumber
        
End Sub



Sub WriteCSVFile_radcheck()

    Dim My_filenumber As Integer
    Dim logSTR As String
    Dim myData As Range
    
    Set sheet_radcheck = Worksheets("tbl_radcheck")
        
    My_filenumber = FreeFile
    Open saveToPath + "radcheck.csv" For Output As #My_filenumber
    
    totRows = sheet_radcheck.UsedRange.Columns("A").Cells.Count
    
    rrow = 1
    
    For Each myData In sheet_radcheck.UsedRange.Columns("A").Cells
    
        logSTR = logSTR & sheet_radcheck.Cells(rrow, 1).Value & ";"
        logSTR = logSTR & sheet_radcheck.Cells(rrow, 2).Value & ";"
        logSTR = logSTR & sheet_radcheck.Cells(rrow, 3).Value & ";"
        logSTR = logSTR & sheet_radcheck.Cells(rrow, 4).Value & ";"
        
        If rrow <= totRows - 1 Then
            logSTR = logSTR & sheet_radcheck.Cells(rrow, 5).Value & vbNewLine
        Else
            logSTR = logSTR & sheet_radcheck.Cells(rrow, 5).Value
        End If
        
        rrow = rrow + 1
        
    Next
    
    Print #My_filenumber, logSTR
    Close #My_filenumber
        
End Sub



Sub WriteCSVFile_radreply()

    Dim My_filenumber As Integer
    Dim logSTR As String
    Dim myData As Range
    
    Set sheet_radreply = Worksheets("tbl_radreply")
    
    My_filenumber = FreeFile
    Open saveToPath + "radreply.csv" For Output As #My_filenumber
    
    totRows = sheet_radreply.UsedRange.Columns("A").Cells.Count
    'questa tabella ha due righe per user quindi bisogna calcolare il doppio delle righe
    'totRows = totRows * 2
        
    rrow = 1
    
    For Each myData In sheet_radreply.UsedRange.Columns("A").Cells
    
        logSTR = logSTR & sheet_radreply.Cells(rrow, 1).Value & ";"
        logSTR = logSTR & sheet_radreply.Cells(rrow, 2).Value & ";"
        logSTR = logSTR & sheet_radreply.Cells(rrow, 3).Value & ";"
        logSTR = logSTR & sheet_radreply.Cells(rrow, 4).Value & ";"
        
        If rrow < totRows Then
            logSTR = logSTR & sheet_radreply.Cells(rrow, 5).Value & vbNewLine
        Else
            logSTR = logSTR & sheet_radreply.Cells(rrow, 5).Value
        End If
        
        rrow = rrow + 1
        
    Next
    
    Print #My_filenumber, logSTR
    Close #My_filenumber
        
End Sub


Sub eporta_tutto()

    Call WriteCSVFile_utente
    Call WriteCSVFile_userinfo
    Call WriteCSVFile_radcheck
    Call WriteCSVFile_radreply
    

End Sub


Sub sqlTest()

    Set sheet_userinfo = Worksheets("tbl_userinfo_orig")
    Set sheet_radcheck = Worksheets("tbl_radcheck")
    Set sheet_radreply = Worksheets("tbl_radreply")

   'credenziali MySql radius APN
    Const SERVER = "x.x.x.x"
    Const DB = "db_name"
    Const UID = "db_user"
    Const PWD = "db_password"
      
    'Define a connection and a recordset to hold extracted information
    Dim oConn As ADODB.Connection, rcSet As ADODB.Recordset
    Dim cnStr As String, n As Long, msg As String, e
    
    'connection string to connect to db4free.net
    cnStr = "Driver={MySQL ODBC 9.0 ANSI Driver};SERVER=" & SERVER & _
            ";PORT=3306;DATABASE=" & DB & _
            ";UID=" & UID & ";PWD=" & PWD & ";"
    
    'Test SQL query
    Const SQL_radcheck = "SELECT MAX(id) FROM radcheck"
    Const SQL_radreply = "SELECT MAX(id) FROM radreply"
    Const SQL_userinfo = "SELECT MAX(id) FROM userinfo_orig"
    
    ' connect
    Set oConn = New ADODB.Connection
    'oConn.CommandTimeout = 900
    
    On Error Resume Next
    oConn.Open cnStr
    If oConn.Errors.Count > 0 Then
        For Each e In oConn.Errors
            msg = msg & vbLf & e.Description
        Next
        MsgBox msg, vbExclamation, "ERROR - Connection Failed"
        Exit Sub
    Else
        MsgBox "Connected to database " & oConn.DefaultDatabase, vbInformation, "Success"
    End If
    
    'query legge MAX id tabella radcheck
    Set rcSet1 = oConn.Execute(SQL_radcheck, n)
    Set rcSet2 = oConn.Execute(SQL_radreply, n)
    Set rcSet3 = oConn.Execute(SQL_userinfo, n)
    
    If oConn.Errors.Count > 0 Then
        msg = ""
        For Each e In oConn.Errors
            msg = msg & vbLf & e.Description
        Next
        MsgBox msg, vbExclamation, "ERROR - Execute Failed"
    Else
        sheet_radcheck.Range("A1").CopyFromRecordset rcSet1
        sheet_radreply.Range("A1").CopyFromRecordset rcSet2
        sheet_userinfo.Range("A1").CopyFromRecordset rcSet3
        'MsgBox SQL_radcheck & " returned " & n & " records", vbInformation
    End If
    On Error GoTo 0
    
    oConn.Close
    
    MsgBox "table ID exported, closing SQL connection"
    
    Call getUser
    
End Sub



