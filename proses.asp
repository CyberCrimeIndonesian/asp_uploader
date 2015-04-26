<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="uploader.asp" -->
<%

Dim Uploader, File
Set Uploader = New FileUploader

Uploader.Upload()

Response.Write "<b>^_^</b><br>"

If Uploader.Files.Count = 0 Then
   Response.Write "gak ke-upload"
Else
   For Each File In Uploader.Files.Items
      
      If Uploader.Form("saveto") = "disk" Then
   
         File.SaveToDisk "D:\" ' path penyimpanan hasil upload
   
      ElseIf Uploader.Form("saveto") = "database" Then
         
         Set RS = Server.CreateObject("ADODB.Recordset")
         RS.Open "MyUploadTable", "CONNECT STRING OR ADO.Connection", 2, 2
         RS.AddNew
         
         RS("filename")    = File.FileName
         RS("filesize")     = File.FileSize
         RS("contenttype") = File.ContentType
      
         File.SaveToDatabase RS("filedata")
         
         RS.Update
         RS.Close
      End If
      
      Response.Write "File Terupload: " & File.FileName & "<br>"
      Response.Write "Ukuran: " & File.FileSize & " bytes<br>"
      Response.Write "Jenis: " & File.ContentType & "<br><br>"
   Next
End If

%><%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="uploader.asp" -->
<%

Dim Uploader, File
Set Uploader = New FileUploader

Uploader.Upload()

Response.Write "<b>^_^</b><br>"

If Uploader.Files.Count = 0 Then
   Response.Write "gak ke-upload"
Else
   For Each File In Uploader.Files.Items
      
      If Uploader.Form("saveto") = "disk" Then
   
         File.SaveToDisk "D:\" ' path penyimpanan hasil upload
   
      ElseIf Uploader.Form("saveto") = "database" Then
         
         Set RS = Server.CreateObject("ADODB.Recordset")
         RS.Open "MyUploadTable", "CONNECT STRING OR ADO.Connection", 2, 2
         RS.AddNew
         
         RS("filename")    = File.FileName
         RS("filesize")     = File.FileSize
         RS("contenttype") = File.ContentType
      
         File.SaveToDatabase RS("filedata")
         
         RS.Update
         RS.Close
      End If
      
      Response.Write "File Terupload: " & File.FileName & "<br>"
      Response.Write "Ukuran: " & File.FileSize & " bytes<br>"
      Response.Write "Jenis: " & File.ContentType & "<br><br>"
   Next
End If

%>
