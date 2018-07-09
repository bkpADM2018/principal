	
var UPLOAD_FILE = "file";
var UPLOAD_IMG = "imgUpld";
var UPLOAD_FRAME = "ifrm";
var UPLOAD_FORM = "formUpld";
var UPLOAD_MAX_CHECK = 120;

var $ulm_ch = new channel();
var $ulm_handlers = new Array();

/*******************************
  *****    UPLOAD HANDLER     *****
  *******************************
  ** Clase encargada de administrar **
  ** un  archivo submitido.                 **
  *******************************/
function UploadHandler(theDiv, folder) {

	this.id = "F" + theDiv;
	this.theDiv = theDiv;
	this.folder = folder.replace(/\\/g,"$");
	this.clientFile = "";
	this.canDelete = true;
	this.isReady = true;
	this.form = undefined;
	this.isError = false;
	this.checkCount = 0;
	
	this.upload = 	function() {
						var iFile = document.getElementById(UPLOAD_FILE + this.id);
						if (iFile.value != "") {
							document.getElementById(UPLOAD_IMG + this.id).className="imgLoadingOn";
							iFile.className="ifrmStyle";
							this.form.appendChild(iFile);
							this.form.submit();
							setTimeout("uploadCheck('" + this.id + "')",1000);
							this.isReady = false;
						}
					}
	this.remove = 	function() {											
						document.getElementById(UPLOAD_IMG + this.id).className="imgLoadingOn";									
						$ulm_ch.bind("uploadSubmitFile2.asp?accion=remove&folder=" + this.folder + "&file=" + this.clientFile, "removeFile_callBack('" + this.id + "')");
						$ulm_ch.send();							
						this.isReady = false;						
					}
					
	this.checkUpload = 	function() {
							this.checkCount++;
							var path = document.getElementById(UPLOAD_FILE + this.id).value;
							this.clientFile = path;
							var pos = path.lastIndexOf("\\");
							if (pos != -1) this.clientFile = path.substring(pos+1);
							$ulm_ch.bind("uploadStatus.asp?accion=upload&folder=" + this.folder + "&file=" + this.clientFile,"uploadFile_callBack('" + this.id + "')");
							$ulm_ch.send();
						}					
	this.setFolder = 	function(aFolder) {
							this.folder = aFolder.replace(/\\/g,"$");	
						}
						
	this.setFile =	function(aFileName) {
								if (aFileName != "") {									
									this.clientFile = aFileName;																							
								}
							}
	this.blockDelete = 	function() {
							this.canDelete = false;
						}
	
	this.draw =	function() 	{
								if (this.clientFile != "") {
									document.getElementById(this.theDiv).innerHTML = this.drawFlat();
								} else {							
									if (!this.form) this._buildForm();
									document.getElementById(this.theDiv).innerHTML = this.drawInput();									
								}																
							}
						
	this.drawInput =	function() {	
							this.isError = false;
							this.checkCount = 0;
							var html = "<input type='file' name='"  + UPLOAD_FILE + this.id + "' id='" + UPLOAD_FILE + this.id + "' onblur=\"javascript:uploadFile('" + this.id + "')\" />&nbsp;";
							html += "<img id='" + UPLOAD_IMG + this.id + "' src='images/loading_small_green.gif' class='loadingOff'>";
							html += "<iframe class='ifrmStyle' name='" + UPLOAD_FRAME + this.id + "' id='" + UPLOAD_FRAME  + this.id + "'></iframe>";
							return html;
				}
	
	this.drawFlat =	function() {
						var html = this.clientFile + "&nbsp;";
						if (this.isError) {
							html += "<img src='images/upload/upload_error.png'>";
						} else {
							html += "<img src='images/upload/upload_ok.png'>";
						}
						if (this.canDelete) {
							html += "&nbsp;&nbsp;<span style=\"cursor: pointer\" onClick=\"javascript:removeFile('" + this.id + "')\">-Cambiar-</span>&nbsp;<img id='" + UPLOAD_IMG + this.id + "' src='images/upload/loading_small_green.gif' class='loadingOff'>";
						}						
						return html;
					}
	this._buildForm =	function() {
							this.form = document.createElement("form");
							this.form.method = "POST";
							this.form.encoding = "multipart/form-data";
							this.form.name= UPLOAD_FORM + this.id;
							this.form.id= UPLOAD_FORM + this.id;
							this.form.action = "uploadSubmitFile2.asp?accion=upload&folder=" + this.folder;
							this.form.target = UPLOAD_FRAME + this.id;
							document.body.appendChild(this.form);
						}
	this.confirmUpload =function() {				
							this.draw();
							this.isReady = true;
						}
										
	this.confirmError = function(err) {
							this.isError = true;
							this.draw();
							this.isReady = true;
							alert(err);
						}						
						
	this.confirmRemove =function() {
							this.clientFile = "";
							var obj = document.getElementById(UPLOAD_FILE + this.id);		
							if (obj) {
								this.form.removeChild(obj);
							}
							this.draw();
							this.isReady = true;
						}
	
	this.getFileName = 	function() {
							return this.clientFile;
						}
	this.setFileName = 	function(file) {
							this.clientFile = file;
						}	
	$ulm_handlers[this.id] = this;
}
	
/*******************************
  *****  GLOBAL FUNCTIONS *****
  *******************************/
function uploadFile_callBack(id) {
	var resp = $ulm_ch.response();	
	var respC = resp.split('|');
	if (respC[0] == "DONE") {					
		$ulm_handlers[id].confirmUpload();
	} else if (respC[0] == "ERROR") {			
		$ulm_handlers[id].confirmError(respC[1]);		
	} else {
		if ($ulm_handlers[id].checkCount == UPLOAD_MAX_CHECK) {
			$ulm_handlers[id].confirmError("ERROR: TIME OUT");			
		} else {
			setTimeout("uploadCheck('" + id + "')",1000);
		}
	}
}
function uploadCheck(id) {	
	$ulm_handlers[id].checkUpload();
}

function removeFile_callBack(id) {
	$ulm_handlers[id].confirmRemove();	
}
function removeFile(id) {	
	if ($ulm_handlers[id]) $ulm_handlers[id].remove();			
}

function uploadFile(id) {	
	if ($ulm_handlers[id]) $ulm_handlers[id].upload();			
}
	