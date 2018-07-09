<%
'           ****************************
'           **   FUNCIONES GRAFICAS   **
'           ****************************
'---------------------------------------------------------------------------------------------            
'   Interfaz:
'
'   GF_CILINDRO_3D(P_intCapacidad,P_intUtilizado,P_imgFondo,P_imgFrente,Xo,Yo,P_Wth,
'                  P_Hgt,P_Desviacion)
'---------------------------------------------------------------------------------------------            
function GF_CILINDRO_3D(P_intCapacidad,P_intUtilizado,P_imgFondo,P_imgFrente,Xo,Yo,P_Wth,P_Hgt,P_Desviacion)

Dim intPorcentaje
DIm imgTop,imgBody,imgBottom

'Armo las imagenes
imgTop=Left(P_imgFrente,len(P_imgFrente)-4) & "_top" & Right(P_imgFrente,4)
imgBody=Left(P_imgFrente,len(P_imgFrente)-4) & "_body" & Right(P_imgFrente,4)
imgbottom=Left(P_imgFrente,len(P_imgFrente)-4) & "_bottom" & Right(P_imgFrente,4)
'Calculo el porcentaje de utilizacion en pixels.
if (P_intUtilizado > P_intCapacidad) then P_intUtilizado=P_intCapacidad
intPorcentaje= CInt((P_intUtilizado*(P_Hgt-(2*P_Desviacion)))/P_intCapacidad)
'***GRAFICO***
'Imagen de fondo.
Response.Write("<img src='images/" & P_imgFondo & "' style='POSITION:relative;")
Response.Write("TOP:" & Yo & "px; LEFT:" & int(Xo+(1.5*P_Wth)) & "px;'>")
'Imagen de frente
Response.Write("<img src='images/" & imgBody & "' style='POSITION:relative;")
Response.Write("TOP:" & int(Yo-(P_desviacion))  & "px; LEFT:" & int(Xo+(0.5*P_Wth)-1) & "px;")
Response.Write("WIDTH:" & int(P_Wth+1) & "px; HEIGHT:" & int(intPorcentaje) & "px;'") 
Response.Write("height='" & int(P_Hgt-(2*P_Desviacion)) & "'>")

Response.Write("<img src='images/" & imgBottom & "' style='POSITION:relative;")
Response.Write("TOP:" & int(Yo) & "px; LEFT:" & int(Xo-(0.5*P_Wth)-2) & "px;'>")

Response.Write("<img src='images/" & imgTop & "' style='POSITION:relative;")
Response.Write("TOP:" & int(Yo-(intPorcentaje)) & "px; LEFT:" & int(Xo-(1.5*P_Wth)-3) & "px;'>")

end function 


%>
