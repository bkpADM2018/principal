<% 
Class cmd
	
	private m_wshell
	private m_proc
	private m_hayError
	private m_errorMsg
	private m_response

	Private Sub Class_Initialize
		'propiedades por defecto
		set m_wshell = CreateObject("WScript.Shell") 
		m_hayError = false
	End Sub

	public property get hayError()
		hayError = m_hayError
	end property

	public property get errorMsg()
		errorMsg = m_errorMsg
	end property

	public property get response()
		response = m_response
	end property


	Function exec(pCommand)
		Dim auxError
		'response.write "cmd.exe /c "& pCommand

		set m_proc =  m_wshell.exec("cmd.exe /c "& pCommand)

		m_errorMsg = m_proc.Stderr.readall

		if (m_errorMsg <> "") then 
			m_hayError = true
		else
			m_response =  m_proc.StdOut.readall
			m_hayError = false
		end if
	End Function

End Class
%>