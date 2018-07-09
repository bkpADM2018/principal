<%
Dim proc, dic, key
Dim myLink

Set dic=Server.CreateObject("Scripting.Dictionary")
 
'Intento descargar los datos del query string
for each key in Request.Querystring
    if (not dic.Exists(key)) then dic.Add key, Request.Querystring(key)
next

'Intento descargar los datos del form
for each key in Request.Form
    if (not dic.Exists(key)) then dic.Add key, Request.Form(key)    
next

'Tomo los cvalores de la session
For Each key in Session.Contents
    dic.Add "SESSION_" & key, Session.Contents(key) 
next


if (dic.Item("proc") <> "") then 
    For each key in dic
        if (myLink="") then
            myLink = "?"
        else
            myLink = myLink & "&"        
        end if        
        myLink = myLink & key & "=" & dic.Item(key) 
    next
    response.Redirect "http://bai-vm-intra-1/actisaintra/" & dic.Item("proc") & ".aspx" & myLink
end if    
 %>
