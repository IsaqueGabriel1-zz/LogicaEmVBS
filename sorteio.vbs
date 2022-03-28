dim n, palavra, cont
dim facil(5), audio
dim medio(5), dificil(5), veryhard(5)

call carregar_voz
sub carregar_voz()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1 'Velocidade da fala
call sorte
end sub

sub fucFacil()
	facil(1)="alicate"
	facil(2)="ameixa"
	facil(3)="amiga"
	facil(4)="amizade"
	facil(5)="amora"
	Randomize(Second(time))
	n=int(Rnd * 5) + 1
	audio.speak(facil(n))
end sub

sub fucMedio()
	medio(1)="quiabo;"
	medio(2)="barriga;"
	medio(3)="pássaro;"
	medio(4)="tesoura;"
	medio(5)="chapéu;"
	Randomize(Second(time))
	n=int(Rnd * 5) + 1
	audio.speak(medio(n))
end sub

sub fucDificil()
	dificil(1)="cabeleireiro;"
	dificil(2)="advogado;"
	dificil(3)="absurdo;"
	dificil(4)="desenhar;"
	dificil(5)="família;"
	Randomize(Second(time))
	n=int(Rnd * 5) + 1
	audio.speak(dificil(n))
end sub



sub sorte()
call fucFacil
palavra=CStr(InputBox("[1] ouvir novamente" + vbNewLine &_
					  "[2] pular palavra"))
					  
if palavra = facil(n) Then
	call fucMedio
	palavra2=CStr(InputBox("[1] ouvir novamente" + vbNewLine &_
					  "[2] pular palavra")) 
					  
					  
	if palavra2 = medio(n) Then
		call fucDificil
		palavra=CStr(InputBox("[1] ouvir novamente" + vbNewLine &_
						  "[2] pular palavra"))
						  
						  
		if palavra = dificil(n) Then
			MsgBox("ok")
			
			
		Else
			MsgBox("Erro if3")
		end if
	Else
		MsgBox("Erro if2")
	end if
Else
	MsgBox("Deu ruim")
end if
end sub