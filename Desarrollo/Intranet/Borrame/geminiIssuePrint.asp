<!-- #include file="Includes/procedimientosParametros.asp"-->
<!-- #include file="Includes/procedimientosCompras.asp"-->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosGemini.asp"-->
<!-- #include file="Includes/procedimientosJSON.asp"-->
<%
Const VERDE  = "#517B4A"
Const NEGRO  = "#000000"
Const BLANCO = "#FFFFFF"

Const ALTURA_RENGLON = 15
Const SEPARATION = 10
Const PAGE_HEIGHT_SIZE = 800
Const INIT_PAGE = 70

Dim oPDF,nroPagina,rsReporte
'--------------------------------------------------------------------------------------------------
Function dibujarReporte(p_IdTarea)
	Dim indexY
    Call dibujarEncabezado()
	Set rsReporte = getIssue(p_IdTarea)
    if (not rsReporte.Eof) then
        indexY = INIT_PAGE
        Call dibujarCabecera(rsReporte, indexY)
        Call dibujarFirmas(rsReporte, indexY)
        if (CInt(rsReporte("templateid")) = TEMPLATE_BUG_TRACKING) then
            'Call dibujarObjetosAS400(p_IdTarea, indexY)
            Call dibujarArchivos(p_IdTarea, indexY)        
        end if            
    end if
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarTituloArchivos(ByRef p_IndexY)
    Call GF_setFont(oPDF,"ARIAL", 8, FONT_STYLE_BOLD)
    p_IndexY = p_IndexY + ALTURA_RENGLON
    Call GF_squareBox(oPDF, 10, p_IndexY, 575, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
    Call GF_setFontColor(BLANCO)
    p_IndexY = p_IndexY + 3
    Call GF_writeTextAlign(oPDF, 15, p_IndexY , "Windows Files", 570,PDF_ALIGN_LEFT)
    Call GF_setFontColor(NEGRO)    
    p_IndexY = p_IndexY + ALTURA_RENGLON    
    'Se dibujan los titulos de los datos propios de los archivos.
    Call GF_setFont(oPDF,"COURIER", 8, FONT_STYLE_BOLD)
    Call GF_writeTextAlign(oPDF,  15, p_IndexY, "Filename",  435,PDF_ALIGN_LEFT)            
    Call GF_writeTextAlign(oPDF, 485, p_IndexY, "Developer", 105,PDF_ALIGN_LEFT)
    p_IndexY = p_IndexY + SEPARATION
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarArchivos(p_IdTarea, ByRef p_IndexY)
    Dim rsFile, sourceFile, developer, nameFile
    
    Call GF_BD_GEMINI(rsFile, "OPEN", "Select cc.data, cc.fullname, gu.username from gemini_codecommits cc inner join gemini_issueresources ir on cc.issueid=ir.issueid inner join gemini_users gu on gu.userid=ir.userid where cc.issueid=" & p_IdTarea & " order by cc.created desc")
    Call dibujarTituloArchivos(p_IndexY)
    Call GF_setFont(oPDF,"COURIER", 6, FONT_STYLE_NORMAL)
    if (not rsFile.eof) then
        'Se arma la lista de todos los archivos modificados eliminanado los duplicados, 
        'como se ordeno la consulta por fecha descendente siempre me quedo con el ultimo desarrollador que modifico el archivo.        
        Set fileDic = CreateObject("Scripting.Dictionary")
        Set oJSON = New JSONReader        
        while (not rsFile.eof)
            oJSON.loadJSON(rsFile("data"))
            For Each sourceFile In oJSON.data("Files")            
                Set this = oJSON.data("Files").item(sourceFile)
                nameFile = Replace(this.item("Filename"),"/trunk/","")
                if (Len(nameFile) > 130) then nameFile = ".." & Right(nameFile,125)
                developer = rsFile("fullname")
				if (developer = "JAS") then developer = rsFile("username")
                if (not fileDic.Exists(nameFile)) then fileDic.Add nameFile, developer               
            Next
            rsFile.MoveNext()
        wend
        'Se imprimen los archivos.
        For Each nameFile In fileDic           
            if (p_IndexY > PAGE_HEIGHT_SIZE) then	
                    Call GF_setFont(oPDF,"COURIER", 8, FONT_STYLE_NORMAL)		            
                    Call GF_writeTextAlign(oPDF, 15, p_IndexY, "-----  Continue on next page  -----", 555,PDF_ALIGN_CENTER)
                    Call nuevaPagina(p_IndexY)
                    Call dibujarTituloArchivos(p_IndexY)
                    Call GF_setFont(oPDF,"COURIER", 6, FONT_STYLE_NORMAL)
            end if            
            developer = fileDic(nameFile)
            Call GF_writeTextAlign(oPDF,  15, p_IndexY, nameFile, 435,PDF_ALIGN_LEFT)
            Call GF_writeTextAlign(oPDF, 485, p_IndexY, getUserDescription(developer),  105,PDF_ALIGN_LEFT)
            p_IndexY = p_IndexY + SEPARATION
        Next
        Call GF_setFont(oPDF,"COURIER", 8, FONT_STYLE_NORMAL)
        Call GF_writeTextAlign(oPDF, 15, p_IndexY, "-----  End of Windows Files  -----", 555,PDF_ALIGN_CENTER)
        Set oJSON = Nothing
    else
        Call GF_setFont(oPDF,"COURIER", 8, FONT_STYLE_NORMAL)
        Call GF_writeTextAlign(oPDF, 15, p_IndexY, "-----  No files were modified  -----", 555,PDF_ALIGN_CENTER)
    end if
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarTituloObjetosAS400(ByRef p_IndexY)
    Call GF_setFont(oPDF,"ARIAL", 8, FONT_STYLE_BOLD)
    p_IndexY = p_IndexY + ALTURA_RENGLON
    Call GF_squareBox(oPDF, 10, p_IndexY, 575, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
    Call GF_setFontColor(BLANCO)
    p_IndexY = p_IndexY + 3
    Call GF_writeTextAlign(oPDF, 15, p_IndexY , "iSeries Objects", 570,PDF_ALIGN_LEFT)
    Call GF_setFontColor(NEGRO)
    p_IndexY = p_IndexY + ALTURA_RENGLON  
    'Se dibujan los titulos de los datos propios de los archivos.
    Call GF_setFont(oPDF,"COURIER", 8, FONT_STYLE_BOLD)
    Call GF_writeTextAlign(oPDF,  15, p_IndexY, "Filename",  435,PDF_ALIGN_LEFT)            
    Call GF_writeTextAlign(oPDF, 485, p_IndexY, "Developer", 105,PDF_ALIGN_LEFT)
    p_IndexY = p_IndexY + SEPARATION      
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarObjetosAS400(p_IdTarea, ByRef p_IndexY)
    Dim rsFile
    
    Call executeQuery(rsFile, "OPEN", "Select DISTINCT SYBIBL, SYOBJE from SYSFL.SYS001F1 where SYISSU like '%" & p_IdTarea & "%'")
    Call dibujarTituloObjetosAS400(p_IndexY)
    Call GF_setFont(oPDF,"COURIER", 8, FONT_STYLE_NORMAL)
    if (not rsFile.eof) then       
        while (not rsFile.eof)
            if (p_IndexY > PAGE_HEIGHT_SIZE) then			            
                    Call GF_writeTextAlign(oPDF, 15, p_IndexY, "-----  Continue on next page  -----", 555,PDF_ALIGN_CENTER)
                    Call nuevaPagina(p_IndexY)
                    Call dibujarTituloObjetosAS400(p_IndexY)
                    Call GF_setFont(oPDF,"COURIER", 8, FONT_STYLE_NORMAL)
            end if                
            nameFile = Trim(rsFile("SYBIBL")) & "  /  " & rsFile("SYOBJE")            
            Call GF_writeTextAlign(oPDF, 15, p_IndexY, nameFile, 435,PDF_ALIGN_LEFT)
            Call GF_writeTextAlign(oPDF, 485, p_IndexY, getUserAndIssue(p_IdTarea),  105,PDF_ALIGN_LEFT)            
            p_IndexY = p_IndexY + SEPARATION			                        
            rsFile.MoveNext()
        wend
        Call GF_writeTextAlign(oPDF, 15, p_IndexY, "-----  End of iSeries objects  -----", 555,PDF_ALIGN_CENTER)        
    else
        Call GF_writeTextAlign(oPDF, 15, p_IndexY, "-----  No iSeries objects were modified  -----", 555,PDF_ALIGN_CENTER)
    end if
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarFirmas(p_Rs, ByRef p_IndexY)
    Dim indiceY
    
    p_IndexY = p_IndexY + ALTURA_RENGLON

    Call GF_squareBox(oPDF, 10,  p_IndexY, 100, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
    Call GF_squareBox(oPDF, 110, p_IndexY, 190, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
	Call GF_squareBox(oPDF, 300, p_IndexY, 180, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
	Call GF_squareBox(oPDF, 480, p_IndexY, 105, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)

    p_IndexY = p_IndexY + ALTURA_RENGLON

    Call GF_squareBox(oPDF, 10,  p_IndexY, 100, ALTURA_RENGLON * 2, 0,  VERDE, NEGRO, 1, 0)

    p_IndexY = p_IndexY + (ALTURA_RENGLON * 2)

    Call GF_squareBox(oPDF, 10,  p_IndexY, 100, ALTURA_RENGLON * 2, 0,  VERDE, NEGRO, 1, 0)
    
    p_IndexY = p_IndexY + (ALTURA_RENGLON * 2)
       

    Call GF_setFontColor(BLANCO)
	Call GF_setFont(oPDF,"ARIAL", 8, FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF, 110, 134, "USER", 190,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 300, 134, "MMTO", 180,PDF_ALIGN_CENTER)	
    Call GF_writeTextAlign(oPDF, 480, 134, "HKEY", 105,PDF_ALIGN_CENTER)
    
    if (CInt(p_Rs("templateid")) = TEMPLATE_BUG_TRACKING) then                    
        Call GF_squareBox(oPDF, 10,  p_IndexY, 100, ALTURA_RENGLON * 2, 0,  VERDE, NEGRO, 1, 0)
        p_IndexY = p_IndexY + (ALTURA_RENGLON * 2)
                    
        Call GF_writeTextAlign(oPDF, 10, 155, "TESTING", 100,PDF_ALIGN_CENTER)
	    Call GF_writeTextAlign(oPDF, 10, 185,"APROBACION", 100,PDF_ALIGN_CENTER)
	    Call GF_writeTextAlign(oPDF, 10, 215,"PUBLICACION", 100,PDF_ALIGN_CENTER)
	    
	    Call GF_squareBox(oPDF, 110,  205, 190, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)	    
        Call GF_squareBox(oPDF, 300,  205, 180, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)
        Call GF_squareBox(oPDF, 480,  205, 105, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)    
    else
        Call GF_writeTextAlign(oPDF, 10, 155, "TECNICO", 100,PDF_ALIGN_CENTER)
	    Call GF_writeTextAlign(oPDF, 10, 185, "SOLICITANTE", 100,PDF_ALIGN_CENTER)
    end if	    
    
    Call GF_squareBox(oPDF, 110,  145, 190, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)
    Call GF_squareBox(oPDF, 300,  145, 180, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)  
    Call GF_squareBox(oPDF, 480,  145, 105, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)
    
    Call GF_squareBox(oPDF, 110,  175, 190, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)
    Call GF_squareBox(oPDF, 300,  175, 180, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)    
    Call GF_squareBox(oPDF, 480,  175, 105, ALTURA_RENGLON * 2, 0,  BLANCO, NEGRO, 1, 0)    

    Call GF_setFontColor(NEGRO)
    indiceFirmasY = 155
    Set rsFirmas = getIssueSign(p_Rs("issueid"))
    'Vienen ordenadas con la nueva secuencia: 
    '   1)Tester/Tecnico - 2)Solicitante - 3)Publicador
    while (not rsFirmas.Eof)
        Call GF_writeTextAlign(oPDF, 115, indiceFirmasY, getUserDescription(rsFirmas("CDUSUARIO")), 190,PDF_ALIGN_LEFT)
	    if (rsFirmas("MMTO") <> "") then Call GF_writeTextAlign(oPDF, 300, indiceFirmasY, GF_FN2DTE(rsFirmas("MMTO")), 180,PDF_ALIGN_CENTER)
	    Call GF_writeTextAlign(oPDF, 480, indiceFirmasY, rsFirmas("HKEY"), 105,PDF_ALIGN_CENTER)
        indiceFirmasY = indiceFirmasY + (ALTURA_RENGLON * 2)
        rsFirmas.MoveNext()
    wend
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarEncabezado()
	
	Call GF_squareBox(oPDF,3,5,590 ,830,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	
	Call GF_writeImage(oPDF, Server.MapPath("Images\kogge64.gif"),6, 6, 60, 60, 0)
	
	Call GF_setFont(oPDF,"ARIAL", 16 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,0, 15, GF_TRADUCIR("I.T. Department") , 550,PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,70,40,515)
	Call GF_writeTextAlign(oPDF,0, 45, GF_TRADUCIR("Project Control") , 590,PDF_ALIGN_CENTER)
	
	Call GF_setFont(oPDF,"ARIAL",8 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina "& nroPagina, 580 , PDF_ALIGN_RIGHT)

End Function
'--------------------------------------------------------------------------------------------------
Function dibujarCabecera(p_Rs, ByRef p_IndexY)
	
	'Se dibuja los boxes de la cabecera.
	Call GF_squareBox(oPDF,  10, p_IndexY, 100, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
	Call GF_squareBox(oPDF, 110, p_IndexY, 475, ALTURA_RENGLON, 0, BLANCO, NEGRO, 1, 0)
	'Call GF_squareBox(oPDF, 462, p_IndexY,  50, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
	'Call GF_squareBox(oPDF, 512, p_IndexY,  73, ALTURA_RENGLON, 0, BLANCO, NEGRO, 1, 0)
	
    p_IndexY = p_IndexY + ALTURA_RENGLON
	
    Call GF_squareBox(oPDF,  10, p_IndexY, 100, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
	Call GF_squareBox(oPDF, 110, p_IndexY, 220, ALTURA_RENGLON, 0, BLANCO, NEGRO, 1, 0)
	Call GF_squareBox(oPDF, 330, p_IndexY,  72, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
	Call GF_squareBox(oPDF, 402, p_IndexY, 183, ALTURA_RENGLON, 0, BLANCO, NEGRO, 1, 0)

    p_IndexY = p_IndexY + ALTURA_RENGLON

    Call GF_squareBox(oPDF,  10, p_IndexY, 100, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
	Call GF_squareBox(oPDF, 110, p_IndexY, 220, ALTURA_RENGLON, 0, BLANCO, NEGRO, 1, 0)
	'Call GF_squareBox(oPDF, 330, p_IndexY,  72, ALTURA_RENGLON, 0,  VERDE, NEGRO, 1, 0)
	'Call GF_squareBox(oPDF, 402, p_IndexY, 183, ALTURA_RENGLON, 0, BLANCO, NEGRO, 1, 0)
	
    p_IndexY = p_IndexY + ALTURA_RENGLON

	'Se escriben los labels fijos.
	Call GF_setFontColor(BLANCO)
	Call GF_setFont(oPDF,"ARIAL", 10, FONT_STYLE_BOLD)	
	Call GF_writeTextAlign(oPDF,  10, 72,           GF_TRADUCIR("Title"), 100, PDF_ALIGN_CENTER)
	'Call GF_writeTextAlign(oPDF, 462, 72,       GF_TRADUCIR("Code"),  50, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,  10, 87,  GF_TRADUCIR("Creator"), 100, PDF_ALIGN_CENTER)		    
	Call GF_writeTextAlign(oPDF, 330, 87,          GF_TRADUCIR("Creation date"),  72, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,  10, 102,  GF_TRADUCIR("Code"), 100, PDF_ALIGN_CENTER)		
    'Call GF_writeTextAlign(oPDF, 330, 102,          GF_TRADUCIR("System"),  72, PDF_ALIGN_CENTER)

    'SE escriben los datos.
	Call GF_setFontColor(NEGRO)
	Call GF_writeTextAlign(oPDF, 115, 72, Trim(p_Rs("SUMMARY")), 336, PDF_ALIGN_LEFT)
	'Call GF_writeTextAlign(oPDF, 522, 72, getGeminiTaskCode(p_Rs("ISSUEID")),  63, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF, 115, 87, getGeminiUserFullName(p_Rs("CREATOR")), 205, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF, 407, 87, Left(p_Rs("CREATED"),10), 180, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, 115, 102, getGeminiTaskCode(p_Rs("ISSUEID")), 205, PDF_ALIGN_LEFT)
	'Call GF_writeTextAlign(oPDF, 407, 102, getDsProjectByIssue(p_Rs("ISSUEID")),    180, PDF_ALIGN_LEFT)
		
	Call GF_setFont(oPDF,"ARIAL", 10 , FONT_STYLE_NORMAL)
End Function
'-------------------------------------------------------------------------------------------------------------------------------
Function nuevaPagina(p_IndexY)
	Call GF_newPage(oPDF)
	nroPagina = nroPagina + 1
	Call dibujarEncabezado()
    p_IndexY = INIT_PAGE
    Call dibujarCabecera(rsReporte, p_IndexY)
end Function
'-------------------------------------------------------------------------------------------------------------------------------
Function crearReporte(p_idTarea)
    Dim pathReport
    pathReport = Server.MapPath("temp\TAREA-" & p_idTarea & ".pdf")
    nroPagina = 1
    Set oPDF = GF_createPDF(pathReport)
    call GF_setPDFMode(PDF_FILE_MODE)
    Call dibujarReporte(p_idTarea)
    Call GF_closePDF(oPDF)
    crearReporte = pathReport
End function
'*****************************************************************************************************************************
Function crearReporteStream(p_idTarea)
    Dim pathReport
    pathReport = Server.MapPath("temp\TAREA-" & p_idTarea & ".pdf")
    nroPagina = 1
    Set oPDF = GF_createPDF(pathReport)
    call GF_setPDFMode(PDF_STREAM_MODE)
    Call dibujarReporte(p_idTarea)
    Call GF_closePDF(oPDF)    
End function

%>