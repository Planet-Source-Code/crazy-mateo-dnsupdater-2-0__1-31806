<!-- #INCLUDE FILE="OpenConnection.asp" --><%
dim temp1(1)
dim temp2()
dim output
temp1(0) = request.querystring("table")
temp1(1) = request.querystring("exclude")
redim temp2(0)
if temp1(0) <> vbnullstring then  
	set rs = db.execute("SELECT * FROM " & temp1(0))
    parsestringtoarray temp2, temp1(1), ";"
    do until rs.eof
		select case lcase(temp1(0))
			case "routers"
   				if inarray(temp2, rs("name")) = false then
   					if output <> vbnullstring then output = output & ";"
    				output = output & rs("name") & "^" & rs("login") & "^" & rs("logout") & "^" & rs("status") & "^" & rs("keyword")
				end if
		    case "services"
    			if inarray(temp2, rs("name")) = false then
    				if output <> vbnullstring then output = output & ";"
    				output = output & rs("name") & "^" & rs("address") & "^" & rs("fields") & "^" & rs("keyword")
				end if
    		case else
     			if inarray(temp2, rs("name")) = false then
     				if output <> vbnullstring then output = output & ";"
    				output = output & rs("name") & "^" & rs("value")
				end if   
    	end select
    	rs.movenext
    loop
    response.write output
else
    response.write "error"
end if

Function parsestringtoarray(arr, str, delimiter)
Dim x, y
If len(str) = 0 Then Exit Function
if isarray(arr) = false then exit function
x = 1: y = 1
Do Until x > Len(str)
    If Mid(str, x, 1) = delimiter And x > 1 Then
        addtoarray arr, Mid(str, y, x - y)
        y = x + 1
    End If
    x = x + 1
Loop
If x > 1 Then
    addtoarray arr, Mid(str, y, Len(str) - y + 1)
End If
End Function

function addtoarray(arr, value)
if isarray(arr) = false then exit function
redim preserve arr(ubound(arr) + 1)
arr(ubound(arr)) = value
end function

function inarray(arr, value)
dim x
if isarray(arr) = false then exit function
x = 0
do until x > ubound(arr)
	if lcase(arr(x)) = lcase(value) then
		inarray = true
		exit function
	end if
	x = x + 1
loop
end function 
%><!-- #INCLUDE FILE="CloseConnection.asp" -->