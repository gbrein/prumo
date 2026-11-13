-- Chamado pelo suplemento via AppleScriptTask("Prumo.scpt", ...).
-- Instalado em ~/Library/Application Scripts/com.microsoft.Powerpoint/Prumo.scpt por panel/build.sh.

-- Nome completo do usuário do Mac, para as iniciais dos comentários.
on nomeUsuario(arg)
	return long user name of (system info)
end nomeUsuario
