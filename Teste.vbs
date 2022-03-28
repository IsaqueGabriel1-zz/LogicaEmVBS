dim n1, n2
call teste

sub teste()
n1=CStr(InputBox("N1: "))
n2=CStr(InputBox("n2: "))

if n1 = n2 then
	MsgBox("Igual")
Else
	MsgBox("Diferente")
end if
end sub