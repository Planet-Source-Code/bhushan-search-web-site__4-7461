<div align="center">

## Search web site


</div>

### Description

Search the web site
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bhushan\-](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bhushan.md)
**Level**          |Intermediate
**User Rating**    |4.8 (67 globes from 14 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bhushan-search-web-site__4-7461/archive/master.zip)





### Source Code

```
As the website increases you should provide an option for the user to trace out what he is looking for, in a very short time. For this reason search this website help you a lot.
Open your favorite editor and paste this code their. save the file as textsearch.asp.
 <B>Search Results for :-<font color=blue > <%=Request("SearchText")%></font></B><BR>
<%
Const fsoForReading = 1
Dim strSearchText
strSearchText = Request("SearchText")
'Now, we want to search all of the files
Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Dim objFolder
Set objFolder = objFSO.GetFolder(Server.MapPath("/news"))
Dim objFile, objTextStream, strFileContents, bolFileFound
bolFileFound = False
dim count
count=0
For Each objFile in objFolder.Files
If Response.IsClientConnected then
Set objTextStream = objFSO.OpenTextFile(objFile.Path,fsoForReading)
strFileContents = objTextStream.ReadAll
If InStr(1,strFileContents,strSearchText,1) then
count=count+1
Response.Write "<LI><A HREF=""/news/" & objFile.Name & _
""">" & objFile.Name & "</A><BR>"
'This program will do the search in the path specified only. if you want it to search through all your 'directories/folders use a different logic. here i have specified the path as /news. change the path accordingly.
Response.Write "<a href=http://www.plnaet-source-code.com" & objFile.Name & "> www.plnaet-source-code.com" & objFile.Name & "</a><br> "
Response.Write ("<br>")
bolFileFound = True
End If
objTextStream.Close
End If
Next
if Not bolFileFound then
Response.Write "No matches found..."
else
Response.Write "Total no of pages found=" & count & "<br>"
'Response.Write "click for more information:-.." & "<br>"
end if
Set objTextStream = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
%>
Next open another page and type in the following code;
<html>
<body>
<form method=post action="textsearch.asp">
<input type=text name=SearchText>
<input type=submit value=search>
</form>
</body>
</html>
```

