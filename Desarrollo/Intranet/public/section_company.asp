<script type="text/javascript">
	$(document).ready(function() {
		$('span')		
			.click(function() {			
				var source = "public/" + $(this).parent().attr('id') + ".asp";
				$("#mainContent").css('display', 'none');
				$("#mainContent").load(source);
				$("#mainContent").delay(0).fadeIn(900);					   
			});		    
	});
</script>			
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
						
						<div class="o-grid-col  o-grid-col_12of12@small o-grid-col_5of12@medium">
							<div class="c-feature">
								<div class="c-feature-hd">
									<h2 class="c-hdg c-hdg_component"><p>ADM Argentina</p></h2>
								</div>
								<div class="c-feature-bd">
									<p>
										<p>Desde comienzos de 1999, ADM Argentina ha crecido para convertirse en uno de los mayores exportadores de ma&iacutez, sorgo y soja del mundo.</p>
										<p>&nbsp;</p>	
										<p>50 empleados en ventas, log&iacutestica y administraci&oacuten atienden a nuestros clientes desde las oficinas en Buenos Aires y Rosario, y administran acuerdos para cargar los productos en muchos puertos de la Argentina.</p>
										<p>&nbsp;</p>
										<p>Manteniendo nuestro compromiso de responsabilidad corporativa, ADM provee asistencia a Caritas Argentina, quien trabaja con 64 organizaciones de caridad para proveer ayuda de emergencia a tres millones de pobres y promover la paz, los derechos humanos y el cuidado del medio ambiente.</p>
									</p>
								</div>
							</div>
						</div>
												
						
						<div class="o-grid-col o-grid-col_6of12@small o-grid-col_4of12@medium">
							<div class="c-feature">
								<div class="c-feature-hd">
									<h2 class="c-hdg c-hdg_4"><p> Datos &Uacutetiles </p></h2>
								</div>
								<div class="c-feature-bd">
									<p>
										<p><a href="Documentos/Politicas/POLITICA_CALIDAD_2018.pdf" target="_blank"> Pol&iacutetica de Calidad </a></p>
										<p>&nbsp;</p>
										<p><a id="section_tax" href="#" class="c-txt c-txt_subAction c-mix-txt_blocked"><span>Datos Impositivos</span></a></p>
									</p>
								</div>
							</div>
						</div>
					</div>
				</div>
			</div>