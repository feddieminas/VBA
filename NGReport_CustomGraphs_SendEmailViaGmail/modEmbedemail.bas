Attribute VB_Name = "modEmbedemail"

' Worked in Windows 10 and Excel 2016

Sub MySendMailViaGmail()
' Sources
' https://www.rondebruin.nl/win/s1/cdo.htm
' https://forums.asp.net/t/2114608.aspx?The+server+rejected+the+sender+address+The+server+response+was+530+5+7+0+Must+issue+a+STARTTLS+command+first+h14sm17016366pgn+41+gsmtp

' Enabled VBA Reference Microsoft CDO for Windows 2000 Library

Dim strPath As String
Dim iMsg As Object
Dim iConf As Object
Dim Flds As Variant
    
With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With
    
On Error GoTo TheEnd
    
strPath = ThisWorkbook.Path & IIf(Right(ThisWorkbook.Path, 1) = "\", "", "\") & "imgs\"
If Not FileFolderExists(strPath) Then
MsgBox "No Images Exist. Macro Exits"
Exit Sub
End If

Dim CountPNGFiles As Integer
CountPNGFiles = CountFilesInFolder(strPath, "*png")
If CountPNGFiles < 12 Then
MsgBox "No Send Email. No 12 pictures appear for a full Status Report"
Exit Sub
End If

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

iConf.Load -1    ' CDO Source Defaults
Set Flds = iConf.Fields
With Flds ' Allow less secure apps on gmail, can try also creating an app password and use that as your password, especially if you use a 2-Factor authentication
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = ThisWorkbook.Worksheets("Buttons").Range("B20").Value
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ThisWorkbook.Worksheets("Buttons").Range("B21").Value
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 ' or 587
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
    .Update
End With

With iMsg
    Set .Configuration = iConf
    .To = ThisWorkbook.Worksheets("Buttons").Range("B24").Value
    .CC = ""
    .BCC = ""
    .From = ThisWorkbook.Worksheets("Buttons").Range("B20").Value
        
Dim objImage1 As Object
Dim myPic1 As String
       
Set objImage1 = .AddRelatedBodyPart(strPath & "mytestfile1.png", "mytestfile1.png", CdoReferenceTypeID)
objImage1.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile1.png>"
objImage1.Fields.Update
       
myPic1 = "<html><img src=""cid:mytestfile1.png""/></br></html>"
        
Dim objImage2 As Object
Dim myPic2 As String
        
Set objImage2 = .AddRelatedBodyPart(strPath & "mytestfile2.png", "mytestfile2.png", CdoReferenceTypeID)
objImage2.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile2.png>"
objImage2.Fields.Update

myPic2 = "<html><img src=""cid:mytestfile2.png""/></br></html>"
       
Dim objImage3 As Object
Dim myPic3 As String
 
Set objImage3 = .AddRelatedBodyPart(strPath & "mytestfile3.png", "mytestfile3.png", CdoReferenceTypeID)
objImage3.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile3.png>"
objImage3.Fields.Update
       
myPic3 = "<html><img src=""cid:mytestfile3.png""/></br></html>"
    
Dim objImage4 As Object
Dim myPic4 As String
    
Set objImage4 = .AddRelatedBodyPart(strPath & "mytestfile4.png", "mytestfile4.png", CdoReferenceTypeID)
objImage4.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile4.png>"
objImage4.Fields.Update
myPic4 = "<html><img src=""cid:mytestfile4.png""/></br></html>"
      
Dim objImage5 As Object
Dim myPic5 As String
      
Set objImage5 = .AddRelatedBodyPart(strPath & "mytestfile5.png", "mytestfile5.png", CdoReferenceTypeID)
objImage5.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile5.png>"
objImage5.Fields.Update
myPic5 = "<html><img src=""cid:mytestfile5.png""/></br></html>"
     
Dim objImage6 As Object
Dim myPic6 As String
     
Set objImage6 = .AddRelatedBodyPart(strPath & "mytestfile6.png", "mytestfile6.png", CdoReferenceTypeID)
objImage6.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile6.png>"
objImage6.Fields.Update
myPic6 = "<html><img src=""cid:mytestfile6.png""/></br></html>"
       
Dim objImage7 As Object
Dim myPic7 As String
             
Set objImage7 = .AddRelatedBodyPart(strPath & "mytestfile7.png", "mytestfile7.png", CdoReferenceTypeID)
objImage7.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile7.png>"
objImage7.Fields.Update
myPic7 = "<html><img src=""cid:mytestfile7.png""/></br></html>"
       
Dim objImage8 As Object
Dim myPic8 As String
                
Set objImage8 = .AddRelatedBodyPart(strPath & "mytestfile8.png", "mytestfile8.png", CdoReferenceTypeID)
objImage8.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile8.png>"
objImage8.Fields.Update
myPic8 = "<html><img src=""cid:mytestfile8.png""/></br></html>"
      
Dim objImage9 As Object
Dim myPic9 As String
         
Set objImage9 = .AddRelatedBodyPart(strPath & "mytestfile9.png", "mytestfile9.png", CdoReferenceTypeID)
objImage9.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile9.png>"
objImage9.Fields.Update
myPic9 = "<html><img src=""cid:mytestfile9.png""/></br></html>"
      
Dim objImage10 As Object
Dim myPic10 As String
      
Set objImage10 = .AddRelatedBodyPart(strPath & "mytestfile10.png", "mytestfile10.png", CdoReferenceTypeID)
objImage10.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile10.png>"
objImage10.Fields.Update
myPic10 = "<html><img src=""cid:mytestfile10.png""/></br></html>"
     
Dim objImage11 As Object
Dim myPic11 As String
     
Set objImage11 = .AddRelatedBodyPart(strPath & "mytestfile11.png", "mytestfile11.png", CdoReferenceTypeID)
objImage11.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile11.png>"
objImage11.Fields.Update
myPic11 = "<html><img src=""cid:mytestfile11.png""/></br></html>"
   
Dim objImage12 As Object
Dim myPic12 As String
   
Set objImage12 = .AddRelatedBodyPart(strPath & "mytestfile12.png", "mytestfile12.png", CdoReferenceTypeID)
objImage12.Fields.Item("urn:schemas:mailheader.Content.ID") = "<mytestfile12.png>"
objImage12.Fields.Update
myPic12 = "<html><img src=""cid:mytestfile12.png""/></br></html>"
        
    .Subject = "Italian Gas Market Status Report " & Format(ThisWorkbook.Worksheets("Sheet1").Range("K1").Value, "YYYYMMDD")
    
    .HTMLBody = myPic1 & myPic2 & myPic3 & myPic4 & myPic5 & myPic6 & myPic7 & myPic8 & myPic9 & myPic10 & myPic11 & myPic12 & myPic101 & myPic102
    .Send
End With
    
TheEnd:
If Err.Number <> 0 Then MsgBox "No Send Email : " & vbNewLine & "1. Check Email Sender and/or Password Input. It works with a gmail account " & _
vbNewLine & "2. Check your Receiver Emails. If multiple emails exist, make sure you have concat them using ;" & _
vbNewLine & "3. Make sure you have an Internet Connection"
Set objImage1 = Nothing: Set objImage2 = Nothing: Set objImage3 = Nothing: Set objImage4 = Nothing: Set objImage5 = Nothing: Set objImage6 = Nothing
Set objImage7 = Nothing: Set objImage8 = Nothing: Set objImage9 = Nothing: Set objImage10 = Nothing: Set objImage11 = Nothing: Set objImage12 = Nothing
Set Flds = Nothing
Set iConf = Nothing
Set iMsg = Nothing
    
With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub



 

 






