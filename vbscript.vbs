dim cor, farol, resp
call escolha

sub escolha()
farol=cint(inputbox("[1] verde" + vbnewline &_
                    "[2] amarelo" + vbnewline & _
                    "[3] vermelho", "cores do semaforo" + vbnewline &_
                    "[0 ou 10] Encerrar"))
select case farol
        case 1:Loyu
            cor="Siga em frente"
        case 2:
            cor="Atencao"
        case 3: 
            cor="Vermelho - pare"
        case 0,10
            resp=msgbox("Deseja encerrar?",vbQuestion + vbYesNo, "ATENCAO")
            if resp=vbYes then
                wscript.quit
            else 
                call escolha
            end if
        case else
            cor="Cor invalida!"
end select
msgbox(""& cor &""),vbInformation+vbokonly, "cores do semaforo"
call escolha
end sub
