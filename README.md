<div align="center">

## Lewies URLDecode


</div>

### Description

Xao Xiong originally posted his version

of DecodeURL and asked for a reply for

a better way. I replied, but it didn't

turn out too well, so i'm posting it here

for him to see.

The approach i took is that when you look

at an encoded URL string, 2 hex characters

follow a percent sign. XAO listed a few

to translate into actual characters, but

he didn't account for all 256 of them.

Check it out!

See http://www.planetsourcecode.com/xq/ASP/

txtCodeId.6728/lngWId.4/qx/vb/scripts/

ShowCode.htm
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Intermediate
**User Rating**    |4.6 (41 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-lewies-urldecode__4-6729/archive/master.zip)





### Source Code

```
<%
Response.Write DecodeURL("http://www.lewismoten.com/1%2B2%3D3.asp")
Function DecodeURL(ByRef pstrURL)
	Dim llngIndex
	Dim llngMaxIndex
	Dim lstrChar
	Dim lstrResult
	llngMaxIndex = Len(pstrURL)
	llngIndex = 1
	Do While llngIndex <= llngMaxIndex
		lstrChar = Mid(pstrURL, llngIndex, 1)
		If lstrChar = "%" Then
			lstrResult = lstrResult & chr("&h" & Mid(pstrURL, llngIndex + 1, 2))
			llngIndex = llngIndex + 3
		ElseIf lstrChar = "+" Then
			lstrResult = lstrResult & " "
			llngIndex = llngIndex + 1
		Else
			lstrResult = lstrResult & lstrChar
			llngIndex = llngIndex + 1
		End If
	Loop
	DecodeURL = lstrResult
End Function
%>
```

