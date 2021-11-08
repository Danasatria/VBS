Set Conn = createObject("ADODB.Connection")
Set Rs = CreateObject("ADODB.recordset")

StrConn = "Provider=SQLOLEDB; Data Source=MSI; initial catalog=UnivContoso; User ID=sa;password=dana10"
Conn.Open StrConn

DelOrIn = MsgBox("Input Press YES | Delete Press NO | Update Press Cancel",vbYesNoCancel+vbQuestion,"Student")
if DelOrIn=vbYes Then
    SqlQueryIn = "Insert Into Student1 (LastName,FirstMidName,EnrollmentDate) Values ('Test','Test','2021/10/09')"
    Set Rs = Conn.execute(SqlQueryIn)
    MsgBox "Input Success"
Elseif DelOrIn=vbNo Then
    SqlQueryDel = "Delete from Student1 Where ID = 2"
    Set Rs = Conn.execute(SqlQueryDel)
    MsgBox "Delete Success"
Elseif DelOrIn=vbCancel Then
    SqlQueryUp = "UPDATE Student1 Set LastName='Tanto',FirstMidName='mamat',EnrollmentDate='2021/10/09' Where ID = 2"
    Set Rs = Conn.execute(SqlQueryUp)
    MsgBox "Update Success"
End if

Set Rs = Conn.execute("Select * From Student1")

str=""

Do while Not Rs.EOF
    val = Rs.fields.item("ID")
    val2 = Rs.fields.item("LastName")
    val3 = Rs.fields.item("FirstMidName")
    val4 = Rs.fields.item("EnrollmentDate")
    str=str&val&" "&val2&" "&val3&" "&val4&vbNewLine
Rs.MoveNext
Loop

MsgBox str