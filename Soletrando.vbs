dim palavras(20), resp, audio, cont, pontos
dim nomeUsuario, i, n, nivel
pontos=0
call inicioUsuario
sub inicioUsuario()
	nomeUsuario=CStr(InputBox("Digite seu nome usuario: "))
	call carregar_voz
end sub


call carregar_voz
sub carregar_voz()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1 
call jogo
end sub


sub jogo()
cont=0
nivel = nivel + 1
for i=1 to 5 step 1
	
	palavras(1) = "Jaca"
	palavras(2) = "Banana"
	palavras(3) = "Limao"
	palavras(4) = "Macarrao"
	palavras(5) = "Feijao"
	
	do while i <= 5 
		Randomize(Second(time))
		n=int(Rnd * 5) + 1
		audio.speak(palavras(n))
		resp=CStr(InputBox("[O] Ouvir novamente" + vbNewLine &_
					   "[P] pular palavra"))
			
		'MsgBox(palavras(n)),vbInformation,"QUANTIDADE DE PALAVRAS: " & i & ""
		if cont < 3 Then
			if resp = palavras(n) then
				cont = cont + 1
				pontos = pontos + 1
			else 
				MsgBox("Voce erro!!!"),vbExclamation+vbOKOnly, "ATENCAO"
				call placarFinal
			end if
		Else
			if cont = 3 Then
				call nivel2
			end if
		end if
		i = i + 1 
	Loop
Next
end sub

sub nivel2()
	cont=0
	nivel = nivel + 1
	MsgBox("Voce chegou no nivel 2 parabens!!!" + vbNewLine &_
		   "Porem, cuidado jovem gafanhoto, as coisas vao comecar a ficar dificeis!!! :D"),vbInformation+vbOKOnly, "AVISO"
for i=1 to 5 step 1
	
	palavras(1) = "Agnostico"
	palavras(2) = "Alvissaras"
	palavras(3) = "Deleterio"
	palavras(4) = "Desasnado"
	palavras(5) = "Filantropo"
	
	do while i <= 5 
		Randomize(Second(time))
		n=int(Rnd * 5) + 1
		audio.speak(palavras(n))
		resp=CStr(InputBox("[O] Ouvir novamente" + vbNewLine &_
					   "[P] pular palavra"))
		'MsgBox(palavras(n)),vbInformation,"QUANTIDADE DE PALAVRAS: " & i & ""
		if cont < 3 Then
			if resp = palavras(n) then
				cont = cont + 1
				pontos = pontos + 1
			else 
				MsgBox("Voce erro!!!"),vbExclamation+vbOKOnly, "ATENCAO"
				call placarFinal
			end if
		Else
			if cont = 3 Then
				call nivel3
			end if
		end if
		i = i + 1 
	Loop
Next
end sub

sub nivel3()
	cont=0
	nivel = nivel + 1
	MsgBox("Voce chegou no nivel 2 parabens!!!" + vbNewLine &_
		   "Porem, cuidado jovem gafanhoto, as coisas vao comecar a ficar dificeis!!! :D"),vbInformation+vbOKOnly, "AVISO"
for i=1 to 5 step 1
	
	palavras(1) = "Agnostico"
	palavras(2) = "Alvissaras"
	palavras(3) = "Deleterio"
	palavras(4) = "Desasnado"
	palavras(5) = "Filantropo"
	
	do while i <= 5 
		Randomize(Second(time))
		n=int(Rnd * 5) + 1
		audio.speak(palavras(n))
		resp=CStr(InputBox("[O] Ouvir novamente" + vbNewLine &_
					   "[P] pular palavra"))
		'MsgBox(palavras(n)),vbInformation,"QUANTIDADE DE PALAVRAS: " & i & ""
		if cont < 3 Then
			if resp = palavras(n) then
				cont = cont + 1
				pontos = pontos + 1
			else 
				MsgBox("Voce erro!!!"),vbExclamation+vbOKOnly, "ATENCAO"
				call placarFinal
			end if
		Else
			if cont = 3 Then
				call nivel3
			end if
		end if
		i = i + 1 
	Loop
Next
end sub

sub nivel4()
	cont=0
	nivel = nivel + 1
	MsgBox("Voce chegou no nivel 2 parabens!!!" + vbNewLine &_
		   "Porem, cuidado jovem gafanhoto, as coisas vao comecar a ficar dificeis!!! :D"),vbInformation+vbOKOnly, "AVISO"
for i=1 to 5 step 1
	
	palavras(1) = "Hebdomadário"
	palavras(2) = "Iconoclasta"
	palavras(3) = "Idiossincrasia"
	palavras(4) = "Lancinante"
	palavras(5) = "Nitidificar"
	
	do while i <= 5 
		Randomize(Second(time))
		n=int(Rnd * 5) + 1
		audio.speak(palavras(n))
		resp=CStr(InputBox("[O] Ouvir novamente" + vbNewLine &_
					   "[P] pular palavra"))
		'MsgBox(palavras(n)),vbInformation,"QUANTIDADE DE PALAVRAS: " & i & ""
		if cont < 3 Then
			if resp = palavras(n) then
				cont = cont + 1
				pontos = pontos + 1
			else 
				MsgBox("Voce erro!!!"),vbExclamation+vbOKOnly, "ATENCAO"
				call placarFinal
			end if
		Else
			if cont = 3 Then
				call nivel3
			end if
		end if
		i = i + 1 
	Loop
Next
end sub

sub placarFinal()
	MsgBox("Nome: "& nomeUsuario &"" + vbNewLine &_
		   "Pontuacao: "& pontos &"" + vbNewLine &_
		   "Ultimo nivel: "& nivel &""),vbInformation+vbOKOnly, "HEL-HEL-HEL"
	resp=CStr(InputBox("Deseja jogar novamente? Digite Sim para reiniciar o game."),vbInformation+vbOKOnly, "ATECAO")
	if resp = "Sim" Then
		call inicioUsuario
	else 
		wscript.quit
	end if
end Sub

sub GG()
	MsgBox("Parabens jovem gafanhoto, chegaste ao mais alto nivel desse game!!!" + vbNewLine &_
			"Aposto que nao foi facil, mas meus parabens!!!"),vbInformation+vbOKOnly, "PARABENS"
	call placarFinal
end Sub