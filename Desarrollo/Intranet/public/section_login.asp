<!--#include file="loginController.asp"-->
	<div class="o-grid-col"></div>
	<div class="c-tier c-tier_baseOffset">
		<div class="o-wrapper">
			<div class="o-grid">
				
				<div class="o-grid-col  o-grid-col_0of12@small o-grid-col_3of12@medium">
					<div class="c-feature">
						<div class="c-feature-bd">
							<img src="images/watermark-home.png" />
						</div>
					</div>
				</div>
				
				<div class="o-grid-col o-grid-col_6of12@small o-grid-col_4of12@medium">
					<div class="c-feature">
						<div class="c-feature-hd">
							<h2 class="c-hdg c-hdg_4"><p>Iniciar sesi&oacuten.</p></h2>
						</div>
						<div class="c-feature-bd">
							Acceda al Centro de Aplicaciones, sitio desde el cual podr&aacute aprovechar todos los servicios que le brinda nuestra empresa.
						</div>
					</div>
				</div>
				
				<div class="o-grid-col  o-grid-col_12of12@small o-grid-col_5of12@medium">
					<div class="c-feature">
						<div class="c-feature-hd">
							
						</div>
						<div class="c-feature-bd">
							<p>
								<p>Usuario</p>
								<p><input Type="password" Name="Username" Id="Username" maxlength="10"></p>
								<p>&nbsp;</p>
								<p>Contrase&ntildea</p>
								<p><input Type="password" Name="Password" Id="Password" maxlength="30" onkeypress="checkEnter(event)"></p>
								<p>&nbsp;</p>
								<p><input type="button" name="btnLogin" onclick="login();" value="Entrar" /></p>
							</p>									
						</div>
						<div id="msg"></div>
					</div>
				</div>
																	
			</div>
		</div>
	</div>
	<input Type="hidden" Name="llave" Id="llave" value="<% =generarLlave() %>">