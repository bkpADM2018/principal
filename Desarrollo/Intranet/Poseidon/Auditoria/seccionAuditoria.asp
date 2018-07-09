<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<%
Dim pto
pto = GF_PARAMETROS7("pto", "", 6)
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Poseidon - Auditoria</title>
<link rel="stylesheet" href="../../css/ActiSAIntra-1.css" type="text/css">

<style type="text/css">
.table, th, td {
	vertical-align:top;
	}

	#cell .boxround {
		padding: 7px;
	}
}
</style>

<script type="text/javascript">	
	function irControlBalanzaCamiones(){				
		window.open('controlBalanzaCamiones.asp?pto=<%=pto%>','Control Balanza','width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
	}
	function irAjustePuerto(){		
		window.open('../AdministracionAjustes.asp?Pto=<%=pto%>','Control Balanza','width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
	}
	function irMermaVolatil() {
	    window.open("../mermaVolatilAdministrar.asp?Pto=<%=g_strPuerto%>", '<%=g_strPuerto%>', "width=1180,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO");
	}
</script>
</head>

<body>

<table width="500px" border="0" align="center" cellpadding="6" cellspacing="0">
<tr>
<td width="50%" valign="top">
<section id="cell">
            <div class="boxround">
	            <table width="480">
                    <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="irControlBalanzaCamiones()">
	                    <tr>
	                        
	                        <td class="title_sec_section"><% =GF_TRADUCIR("Control de balanza de Camiones")%></td>
	                    </tr>
	                    <tr>
	                        <td class="textoSeccion"><% =GF_TRADUCIR("Controle el peso de camiones por las distintas balanzas del Puerto.") %></td>
	                    </tr>
                    </tbody>
                </table>
            </div>
			<div class="boxround">
				<table width="480">
					<tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="irMermaVolatil()">
						<tr>
							<td class="title_sec_section"><% =GF_TRADUCIR("Administrar Ratios de Merma volatil")%></td>
						</tr>
						<tr>
							<td class="textoSeccion"><% =GF_TRADUCIR("Lista la informaci&oacuten de las mermas volatil que tendrán los camiones que se van cargando en las plantas.") %></td>
						</tr>
					</tbody>
				</table>
			</div>
            <div class="boxround">
	            <table width="480">
                    <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="irAjustePuerto()">
	                    <tr>
	                        
	                        <td class="title_sec_section"><% =GF_TRADUCIR("Ajustes de Stock")%></td>
	                    </tr>
	                    <tr>
	                        <td class="textoSeccion"><% =GF_TRADUCIR("Administrar los ajustes de stock realizados en la planta.") %></td>
	                    </tr>
                    </tbody>
                </table>
            </div>
</section>   
</td>
</tr>
</table>
</body>
</html>
