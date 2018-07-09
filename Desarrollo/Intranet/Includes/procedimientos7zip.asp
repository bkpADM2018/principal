<%
	Const TYPE_7ZIP = "7z"
	Const TYPE_ZIP = "zip"

	Class zip
		private m_zipCmd
		private m_files2Add
		private m_fileName
		private m_filePath
		private m_filePathLogic
		private m_zipType
		private m_zipError

		

		Private Sub Class_Initialize
			set m_zipCmd = new cmd
			Call clearAll()
		End Sub

		public property Let fileName(pValue)
			if (instr(pValue,".7z")>0) then m_zipType = TYPE_7ZIP
			if (instr(pValue,".zip")>0) then m_zipType = TYPE_ZIP

			m_fileName = replace(pValue,"\","")

			if ( instr(m_fileName,".7z")=0 and instr(m_fileName,".zip") = 0) then
				'no tiene extencion, se la agrego'
				m_fileName = m_fileName & "." & TYPE_ZIP
			end if
		end property

		public property get fileName()
			fileName = m_fileName
		End Property

		public property Let filePath(pValue)
			Dim auxPath
			m_filePath = server.MapPath(pValue)
			if (right(m_filePath,1) <> "\") then m_filePath = m_filePath & "\"

			auxPath = split(m_filePath,"actisaintra\")
			m_filePathLogic = auxPath(1)
		end property

		public property get filePath()
			filePath = m_filePath
		end property

		public property get filePathLogic()
			filePathLogic = m_filePathLogic
		end property

		public property get zipError()
			zipError = m_zipError
		End property

		'Agrega archivos o carpetas, pPath debe ser el path fisico del archivo/carpeta a comprimir'
		Function add(pPath)
			redim preserve m_files2Add(ubound(m_files2Add)+1)
			
			m_files2Add(ubound(m_files2Add)) = Chr(34) & pPath & Chr(34)
		End Function


		'En el proceso se reemplazara cualquier archivo que tenga el mismo nombre y extension del que se intenta crear'
		Function zip()
			Dim auxFiles,rtrn
			auxFiles = ""

			rtrn = true
			for i = 0 to ubound(m_files2Add)
				auxFiles = auxFiles & " " & m_files2Add(i)
			next

			if (trim(auxFiles) <> "") then

				if (fso.FileExists(m_filePath & m_fileName)) then
					response.write m_filePath & m_fileName
					fso.deleteFile(m_filePath & m_fileName)
				end if

				m_zipCmd.exec(" 7z a -t"&TYPE_ZIP&" " & Chr(34) & m_filePath & m_fileName & Chr(34) & " " & auxFiles)


				if (m_zipCmd.hayError()) then
					rtrn = false
					m_zipError = m_zipCmd.errorMsg()
				end if

			end if
			zip = rtrn
			'Call clearAll()
			redim m_files2Add(0)
		End Function


		private function clearAll()
			redim m_files2Add(0)
			m_fileName = session("mmtoSistema")
			m_filePath = server.MapPath(".")
			m_filePathLogic = "."
			m_zipType = TYPE_7ZIP
			m_zipError = ""
		End Function

	End Class
%>
