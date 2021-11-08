Set fso = WScript.CreateObject("Scripting.FileSystemObject")

exportPath = "C:\Users\Dana satria\Documents\VBS\"
exportCsv = "Student.csv"

If fso.FolderExists(exportPath) Then
    Set Conn = createObject("ADODB.Connection")
    Set Rs = CreateObject("ADODB.recordset")

    StrConn = "Provider=SQLOLEDB; Data Source=MSI; initial catalog=UnivContoso; User ID=sa;password=dana10"
    Conn.Open StrConn

    Set Rs = Conn.execute("Select * From Student1")

    If Rs.EOF Then
        WScript.Echo "There is no data to export."
    Else
        Set csvFile = fso.CreateTextFile(exportPath & exportCsv, True)
        csvFile.WriteLine("""ID""|""LastName""|""FirstMidName""|""EnrollmentDate""")

        DataRow = ""
        Do while Not Rs.EOF
            dataRow = Chr(34) & Rs("ID") & Chr(34) & "|"
            dataRow = dataRow + Chr(34) & Rs("LastName") & Chr(34) & "|"
            dataRow = dataRow + Chr(34) & Rs("FirstMidName") & Chr(34) & "|"
            dataRow = dataRow + Chr(34) & Rs("EnrollmentDate") & Chr(34) 
            csvFile.WriteLine(dataRow)
        Rs.MoveNext
        Loop

        MsgBox "Data berhasil diExport"
    End If
Else

End If

