<!--#include file="config.asp"-->
<!--#include file="free-asp-upload\freeASPUpload.asp"-->
<%
'''
' A minimal PhotoBackup API endpoint developed in Classic ASP (brutal translated from the PHP one's by Martijn van der Ven).
'
' @version 1.0.0
' @author vespadj
' @copyright 2017
' @license http://opensource.org/licenses/MIT The MIT License
' 
'''
' The password required to upload to this server.
'
' The password is currently stored as clear text here. ASP code is not normally
' readable by third-parties so this should be safe enough. Many applications
' store database credentials in this way as well. A secondary and safer way is
' being considered for the next version.

' Edit config.asp for your personal configuration


''' -----------------------------------------------------------------------------
'' EXAMPLE
'' -----------------------------------------------------------------------------
''' to let test case "/test" url must be finish with ?, 
''' configure server link to this page like this:
'''	http://<server-ip>/upload/index.asp?

''' -----------------------------------------------------------------------------
'' NO CONFIGURATION NECCESSARY BEYOND THIS POINT.
'' -----------------------------------------------------------------------------

' for POST with enctype="multipart/form-data" - the Request.Form is blank!

' https://stackoverflow.com/questions/3649799/asp-request-form-is-not-returning-value
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=7361&lngWId=4#zip

' use FreeASPUpload library

Dim  Upload 
Set  Upload = New FreeASPUpload	' Class on free-asp-upload\freeASPUpload.asp

Upload.Upload	' Load, now you can read: Upload.Form("filesize")
' you must use Upload.Form("password") instaed of Request.Form("password")

' TODO: comment this after debug
' for debug:
Upload.DumpData() ' response passed form-data

'** == it's not need
' Establish what HTTP version is being used by the server.

'If Request.ServerVariables("SERVER_PROTOCOL") <> "" Then
'	protocol = Request.ServerVariables("SERVER_PROTOCOL")
'Else
'	protocol = "HTTP/1.0"
'End If

'**
' Find out if the client is requesting the test page.

testing = Not( Right(Request.ServerVariables("QUERY_STRING"), 5) <> "/test" )

'**
' If no data has been POSTed to the server and the client did not request the
' test page, exit imidiately.

If False And Request.TotalBytes=0 And Not testing Then
	Response.End
End If
'**
' If we are testing the server and see that no password has been set, exit with
' HTTP code 401.

If testing And Password="" Or (VarType(Password) <> vbString) Then
	Response.Status = "401 Unauthorized"
	Response.Write(Response.Status)
	Response.End
End If

'**
' Exit with HTTP code 403 if no password has been set on the server, or if the
' client did not submit a password, or the submitted password did not match
' this server"s password.
' *** DISABLED FOR DEBUG ***
If Password="" _
	Or VarType(Password) <> vbString _
	Or Upload.Form("password") = "" _
	Or Upload.Form("password") <> HashSHA512Managed(Password) _
Then
	Response.Status = "403 Forbidden"
	Response.Write(Response.Status)
	Response.Write(HashSHA512Managed(Password))
	Response.End
End If

'**
' If the upload destination folder has not been configured, does not exist, 
' exit with HTTP code 500.

' If folder is not writeable will appair a Permission Denied error.
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Debug:
'Set objFolder = fso.GetFolder(".")
'Response.write(objFolder.Path & "<br>") ' returns C:\WINDOWS\system32\inetsrv

Dim cwd	' current working directory for this script page
cwd = fso.GetParentFolderName( Request.ServerVariables("PATH_TRANSLATED") )

If MediaRoot<>"" And (VarType(MediaRoot) = vbString) Then
	' if MediaRoot is not a net path and is not an absolute path, 
	' then set base path as cwd (the path of this ASP page)
	If Left(MediaRoot, 2) <> "\\" And Not Instr(MediaRoot, ":") Then
		' this is necessary for not searching in C:\WINDOWS\system32\inetsrv
		MediaRoot = cwd & "\" & MediaRoot
	End If
	
	If Not fso.FolderExists(MediaRoot) Then
		Response.Status = "500 Internal Server Error! MediaRoot Path not exists: " & MediaRoot
		Response.Write(Response.Status)
		Response.End
	End If
Else
	Response.Status = "500 Internal Server Error. Please, set MediaRoot variable on server!"
	Response.Write(Response.Status)
	Response.End
End If

'**
' If we were only supposed to test the server, end here.

If testing = True Then
	Response.write("testing is OK")
	Response.End
End If

'**
' If the client did not submit a filesize, exit with HTTP code 400.

If Upload.Form("filesize") = "" Then
	Response.Status = "400 Bad Request"
	Response.Write(Response.Status)
	Response.End
End If

'** === I don't know how check it in ASP ===
' If the client did not upload a file, or something went wrong in the upload
' process, exit with HTTP code 401.

'If !isset($_FILES["upfile"]) Or $_FILES["upfile"]["error"] !== UPLOAD_ERR_OK Or !is_uploaded_file($_FILES["upfile"]["tmp_name"]) Then
'	Response.Status = "401 Unauthorized"
'	Response.Write(Response.Status)
'	Response.End
'End If

'**
' If the client submitted filesize did not match the uploaded file"s size, exit
' with HTTP code 411.

If CLng(upload.Form("filesize")) <> Upload.UploadedFiles("upfile").Length Then
	Response.Status = "411 Length Required" ' & ". Check: " & Upload.UploadedFiles("upfile").Length
	Response.Write(Response.Status)
	Response.End
End If

' Check Type: image/*
If Left(Upload.UploadedFiles("upfile").ContentType , 6) <> "image/" Then
	Response.Status = "412 Invalid Type"  & ". Passed: " & Upload.UploadedFiles("upfile").ContentType
	Response.Write(Response.Status)
	Response.End
End If

filename = Upload.UploadedFiles("upfile").FileName
extension = GetFileExtension(filename)

' TODO: invalid extensions for ASP server (asp,...)
iArray = Split("asp,asa,inc,exe,msi,cmd,bat,com,vbs,wfs,zip,7z" , ",")

found = false
for i = 0 to ubound(iArray)
    if iArray(i) = extension then
        found = true
    end if
next

If found Then
	Response.Status = "412 Invalid Extension" 
	Response.Write(Response.Status)
	Response.End
End If

'**
' Sanitize the file name to maximise server operating system compatibility and
' minimize possible attacks against this implementation.
Dim objRegExpr
Set objRegExpr = New RegExp
objRegExpr.IgnoreCase = true
objRegExpr.Global = true
'objRegExpr.Pattern = "([^0-9a-z._- ]+)"
objRegExpr.Pattern = "[/\\|<>:\*\?""]"
filename = objRegExpr.Replace(filename, "_")
target = MediaRoot & "\" & filename


'**
' If a file with the same name and size exists, treat the new upload as a
' duplicate and exit.
' TODO: this part may be cause issues because freeASPUpload already numerate and increase same-name file.
conflict = false
If fso.FileExists(target) Then
	' TODO: I'm not sure this is right because same file may be overwrite
	If fso.GetFile(target).Size = Upload.Form("filesize") Then
		conflict = true
	End If
End If

If conflict Then
	Response.Status = "409 Conflict"
	Response.Write(Response.Status)
	Response.End
End If

'**
' Move the uploaded file into the target directory. 
Call upload.SaveOne(MediaRoot, 0, filename, filename)


' If anything did not work,
' exit with HTTP code 500.
' === I dont' know how check it in ASP ===
'If (!move_uploaded_file($_FILES["upfile"]["tmp_name"], $target)) Then
'	Response.Status = "500 Internal Server Error"
'	Response.Write(Response.Status)
'	Response.End
'}
'Response.End


Set fso = Nothing

' ==== Other Functions ===

' from: https://stackoverflow.com/questions/28314564/how-to-replicate-asp-classic-sha512-hash-function-in-php
Function HashSHA512Managed(saltedPassword)
	'Dim objMD5, objUTF8
	Dim arrByte
	Dim strHash
	Set objUnicode = CreateObject("System.Text.UnicodeEncoding")
	Set objSHA512 = Server.CreateObject("System.Security.Cryptography.SHA512Managed")

	arrByte = objUnicode.GetBytes_4(saltedPassword)
	strHash = objSHA512.ComputeHash_2((arrByte))

	HashSHA512Managed = ToBase64(strHash)
	
End Function

Function ToBase64(rabyt)
    Dim xml: Set xml = CreateObject("MSXML2.DOMDocument.3.0")
    xml.LoadXml "<root />"
    xml.documentElement.dataType = "bin.base64"
    xml.documentElement.nodeTypedValue = rabyt
    ToBase64 = xml.documentElement.Text
End Function

Function GetFileExtension(strPath)
    If Right(strPath, 1) <> "." And Len(strPath) > 0 Then
        GetFileExtension = GetFileExtension(Left(strPath, Len(strPath) - 1)) + Right(strPath, 1)
    End If
End Function
%>