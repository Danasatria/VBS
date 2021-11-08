Set OutApp = CreateObject("Outlook.Application")
Set OutMail =OutApp.createItem(0) 

strbody = "<BODY style = font-size:12pt; font-familt:Arial>" & _ 
            "Hi, <br><br> Ini test coba cek attachment. <br><br>"

On Error Resume Next
    With OutMail
        .to = "xsatriax1002@gmail.com"
        .CC = ""
        .BCC = ""
        .Subject = "Test Training"
        .Display 
        .HTMLBody = strbody 
        .Attachments.Add "C:\Users\Dana satria\Documents\VBS\Student.csv"
    End With
On Error GoTo 0

Set OutMail = Nothing
    