<div align="center">

## Simple Download Method with Streams


</div>

### Description

Enables you to put a certain degree of protection in files available for download in you IIS Server by making them inaccessible by a direct URL. Hides them in other directory, out of wwwroot.
 
### More Info
 
Pass the file name.

Get the file written in binary format into Request.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Verzellesi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-verzellesi.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-verzellesi-simple-download-method-with-streams__4-7832/archive/master.zip)





### Source Code

```
<%
	'-- DOWNLOAD.ASP
	'
	'-- Simple Download Method with Streams
	' Daniel Verzellesi
	'
	' As an example I'm getting the file I want to download
	' from the Request (//.../dowload.asp?fname=myfile1.doc).
	' You can change it to get the file name from a DB or
	' anything else...
	'
	Dim p, st, f
	'-- my "secret" path
	p = "c:\files\"
	'-- file name
	f = Request.QueryString("fname")
	'-- get file into stream
	Set st = CreateObject("ADODB.Stream")
	st.Open
	st.Type = 1 'binary
	st.LoadFromFile p & f
	'-- send stream to response
	Response.ContentType = "application/my-download"
	Response.AddHeader "Content-Disposition", "attachment; filename=""" & f & """"
	Response.BinaryWrite st.Read(-1) 'read all
	Response.End
	'-- close the stream
	set st = nothing
%>
```

