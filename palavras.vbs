dim palavras, resp, pontuacao, ponto, ouvirM
dim i, n, audio, val

call carregar_voz
sub carregar_voz()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1 'Velocidade da fala
call teste
end sub


sub teste()
	palavras = Array("alicate","ameixa","amiga","amizade","amora", "chapéu", "tesoura", "pássaro", "barriga", "quiabo")
	
	For i=0 to 10 Step 1
		Randomize(Second(time))
		n=int(Rnd * 9) + 1
		audio.speak(palavras(n))
		resp=CStr(InputBox("[O]uvir novamente" + vbNewLine &_
						   "[P]ular palavra"))
		if resp = "O" Then
			audio.speak(palavras(n))
			ouvirM = ouvirM + 1
		Else
			if resp = "P" Then
				i = i + 1
			Else
				if ponto >= 3 Then
					i = 0
					i = 5
				Else
					if resp = palavras(n) Then
						ponto = ponto + 1
					Else
						MsgBox("Erro")
					end if
				end if
			end if
		end if 
	Next
	MsgBox("Fim do sorteio"), vbInformation+vbOKOnly,"AVISO"
end Sub

