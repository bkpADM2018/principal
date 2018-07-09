<%
'###############################################################################
' ASP Mathematics Library - Floor and Ceiling Functions
' Copyright (c) 2008, reusablecode.blogspot.com; some rights reserved.
' This work is licensed under the Creative Commons Attribution License. To view
' a copy of this license, visit http://creativecommons.org/licenses/by/3.0/ or
' send a letter to Creative Commons, 559 Nathan Abbott Way, Stanford, California
' 94305, USA.
'###############################################################################
' Returns the largest integer less than or equal to the specified number.
function Floor(x)
	dim temp
	temp = Round(x)
	if temp > x then temp = temp - 1	
	floor = temp
end function
' Returns the smallest integer greater than or equal to the specified number.
function Ceil(x)
	dim temp
	temp = Round(x)
	if temp < x then temp = temp + 1	
	ceil = temp
end function
%>
