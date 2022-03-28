dim nome(5),n,i
call sorteio
sub sorteio()
for i=1 to 10 step 1 
	nome(1)="Moquidesia"
	nome(2)="Jurema"
	nome(3)="Lindolfo"
	nome(4)="ademir"
	nome(5)="Azriel"
	Randomize(Second(time))
	n=int(Rnd * 5) + 1
	MsgBox(nome(n)),vbInformation+vbOKOnly,"Qtde Sorteio: "& i &""
next
MsgBox("Fim do sorteio"), vbInformation+vbOKOnly,"AVISO"
end Sub