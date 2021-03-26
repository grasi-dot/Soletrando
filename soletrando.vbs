dim palavra(30), n, i, cont, resp, resp2, resp3
dim validador, audio, escolha, nome, vitoria, acerto, pular, ouvir

call carregar_audio

sub carregar_audio ()
set audio=createobject("SAPI.SPVOICE")
	audio.volume=100
	audio.rate=1
	validador = 0
	call inicio
end sub

sub inicio()
	acerto = 0
	pular = 1
	nome = UCase(inputbox("ATENCAO A PALAVRA SERA FALADA A SEGUIR" + vbNewLine &_
						  "=======================================" + vbNewLine &_
						  "DIGITE SEU NOME"))
	if nome = "" then
		validador = validador + 1
		if validador < 2 then
			msgbox("POR FAVOR ENTRE COM UM NOME")
			call inicio
		else 
			wscript.quit
		end if
	end if
	call sorteio
end sub

sub sorteio ()
ouvir = 1
palavra(1) = "CAPCIOSO"
palavra(2) = "REIVINDICAR"
palavra(3) = "RECAUCHUTAR"
palavra(4) = "IDEM"
palavra(5) = "NEFASTO"
palavra(6) = "ESTENDER"
palavra(7) = "HESITAR"
palavra(8) = "MERITOCRACIA"
palavra(9) = "TACITURNO"
palavra(10) = "METICULOSIDADE"
palavra(11) = "IDONEIDADE"
palavra(12) = "DISSIMULADO"
palavra(13) = "CADEIRA"
palavra(14) = "DISCUTIR"
palavra(15) = "PEJORATIVO"
palavra(16) = "CONSENSO"
palavra(17) = "TANGENTE"
palavra(18) = "EXCLUSIVIDADE"
palavra(19) = "EMPECILHO"
palavra(20) = "PISCINA"
palavra(21) = "PREZADO"
palavra(22) = "PERIQUITO"
palavra(23) = "ASTERISCO"
palavra(24) = "BALBUCIAR"
palavra(25) = "CAMISETA"
palavra(26) = "COCEIRA"
palavra(27) = "PRESSA"
palavra(28) = "EXPLODIR"
palavra(29) = "ASSUSTADO"
palavra(30) = "ESTUPEFATO"
for n = 1 to 30 step 1 
    randomize(second(time))
    i=int(rnd * 30) + 1	
next
	call voz
end sub

sub voz()
audio.speak ("A palavra " + palavra(i))
call entrada
end sub

sub entrada()
resp = Ucase(inputbox("DIGITE A PALAVRA OUVIDA"+ vbNewLine + vbNewLine &_
					      "NOME DO JOGADOR: "& nome & "" + vbNewLine + vbNewLine &_
						  "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + vbNewLine &_
                          "[O]UVIR NOVAMENTE A PALAVRA" + vbNewLine &_
                          "[P]ULAR A PALAVRA UMA UNICA VEZ" + vbNewLine &_
                          "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~","SOLETRANDO"))
	if  resp = "O" then
		if ouvir > 0 then
			ouvir = ouvir -1
	        call voz
		else 
	        msgbox("VOCE JA USOU ESTA OPCAO"), vbinformation + vbOKOnly, "ATENCAO"
	        call entrada
	    end if
	elseif resp = "P" Then
		if pular > 0 then
			pular = pular - 1
			call sorteio
		else
			msgbox("VOCE JA USOU ESTA OPCAO"), vbinformation + vbOKOnly, "ATENCAO"
	        call entrada
		end if	
	elseif  resp = palavra(i) then
		acerto = acerto + 1
		msgbox("PARABENS VOCE ACERTOU!" + vbNewLine &_
			   "QUANTIDADE DE ACERTOS: "& acerto & " DE 12"),vbInformation+vbOKOnly,"SOLETRANDO"	
		if acerto = 12 then
			msgbox("PARABENS VOCE CHEGOU ATE O FINAL")
			resp3 =(msgbox( "QUANTIDADE DE ACERTOS: "& acerto & " DE 12"+ vbNewLine &_
							"DESEJA JOGAR NOVAMENTE? ",vbQuestion + vbYesNo,"FIM DE JOGO"))
			if resp3 = vbYes then
				call inicio
			else
				wscript.quit
			end if	
		end if							
		call sorteio
	elseif resp <> (palavra(i)) then
		msgbox("VOCE PERDEU "& nome & "")
		resp2 =(msgbox("VOCE ERROU!"+ vbNewLine &_
					   "QUANTIDADE DE ACERTOS: "& acerto & " DE 12"+ vbNewLine &_
					   "DESEJA JOGAR NOVAMENTE? ",vbQuestion + vbYesNo,"FIM DE JOGO"))
		if resp2 = vbYes then
			call inicio
		else
			wscript.quit
		end if	
	end if
end sub		


