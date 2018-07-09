<h2><img align="absMiddle" src="images/compras/<%= imagen %>"> <%= titulo %></h2>
<table width="100%">
	<tr valign="top">
		<td>
			<div id="toolBar<% =titulo %>"></div>	
			<form name="frmSel">
				<div id="busqueda<% =seccion %>" class="divOculto">
					<table width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
					   <input type="hidden" name="accion" id="accion" value="">
					   <tr>
						   <td width="8"><img src="images/marco_r1_c1.gif"></td>
						   <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
						   <td width="8"><img src="images/marco_r1_c3.gif"></td>
						   <td width="73%"><td>
						   <td></td>
					   </tr>
					   <tr>
						   <td width="8"><img src="images/marco_r2_c1.gif"></td>
						   <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></td>
						   <td width="8"><img src="images/marco_r2_c3.gif"></td>
						   <td></td>
						   <td></td>
					   </tr>
					   <tr>
						   <td><img src="images/marco_r2_c1.gif" height="8"  width="8"></td>
						   <td></td>
						   <td valign="top" align="right"><img src="images/marco_r1_c2.gif" height="8" width="2"></td>
						   <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
						   <td width="8"><img src="images/marco_r1_c3.gif"></td>
					   </tr>
					   <tr>
						   <td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
						   <td colspan="3">
								<div id="busqueda<% =seccion %>TD"></div>			
							</td>
							   <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
						   </tr>
						   <tr>
								<td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
								<td colspan="3" align="center">
									<input type="button" value="<% =GF_TRADUCIR("Buscar") %>" onClick="javascript:doBuscar(<% =seccion %>)">
								</td>
								<td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
							</tr>
						   <tr>
							   <td width="8"><img src="images/marco_r3_c1.gif"></td>
							   <td width="100%" align=center colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
							   <td width="8"><img src="images/marco_r3_c3.gif"></td>
						   </tr>
					</table>
				</div>
			</form>			
			<div id="paginacion<% =seccion %>"></div>			
			<div id="seccion<% =seccion %>">
				<table width="100%" height="100%" class="reg_header" cellspacing="2" cellpadding="1">
					<tr class="reg_header_nav">
						<td align="center">.</td>
						<td width="15%">Nombre</td>
						<td>Descripcion</td>
						<td align="center">.</td>
					</tr>
					<tr><td colspan="4" align="center"><img src="images/loading_blocks_green.gif"></td></tr>						
				</table>
			</div>				
		</td>
	</tr>
</table>