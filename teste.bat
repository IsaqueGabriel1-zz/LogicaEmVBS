@echo off
title joken-po
mode 60, 20
color 0a
:inicio
cls
set /a cont=0
set /a cont=%cont% + 1
set /a sorteMaquina=(%random% %%5) + 1
set /p nome=Digite o seu nome: 
set /a OpMaquina=PEDRA
set /a opJogador=PEDRA
set /a res=EMPATE
echo [1] PEDRA
echo [2] PAPEL
echo [3] TESOURA
echo [4] LARGATO
echo [5] SPOCK
echo.
set /p golpe=Digite o golpe desejado: 
echo.

if %golpe% == 1 (
    set opJogador=PEDRA
)
if %golpe% == 2 (
    set opJogador=PAPEL
)
if %golpe% == 3 (
    set opJogador=TESOURA
)
if %golpe% == 4 (
    set opJogador=LARGATO
)
if %golpe% == 5 (
    set opJogador=SPOCK
)

if %sorteMaquina% == 1 (
    set OpMaquina=PEDRA
)
if %sorteMaquina% == 2 (
    set OpMaquina=PAPEL
)
if %sorteMaquina% == 3 (
    set OpMaquina=TESOURA
)
if %sorteMaquina% == 4 (
    set OpMaquina=LARGATO
)
if %sorteMaquina% == 5 (
    set OpMaquina=SPOCK
)

if %opJogador% == PEDRA (
    if %OpMaquina% == PEDRA (
        set res=EMPATE
    )
)

if %opJogador% == PEDRA (
    if %OpMaquina% == PAPEL (
        set res=MAQUINA
    )
)

if %opJogador% == PEDRA (
    if %OpMaquina% == TESOURA (
        set res=JOGADOR  
    )
)

if %opJogador% == PEDRA (
    if %OpMaquina% == LARGATO (
        set res=JOGADOR
    )
)

if %opJogador% == PEDRA (
    if %OpMaquina% == SPOCK (
        set res=MAQUINA
    )
)

 
if %opJogador% == PAPEL (
    if %OpMaquina% == PEDRA (
        set res=JOGADOR
    )
)

if %opJogador% == PAPEL (
    if %OpMaquina% == PAPEL (
        set res=EMPATE
    )
)

if %opJogador% == PAPEL (
    if %OpMaquina% == TESOURA (
        set res=MAQUINA  
    )
)

if %opJogador% == PAPEL (
    if %OpMaquina% == LARGATO (
        set res=MAQUINA
    )
)

if %opJogador% == PAPEL (
    if %OpMaquina% == SPOCK (
        set res=JOGADOR
    )
)




if %opJogador% == TESOURA (
    if %OpMaquina% == PEDRA (
        set res=MAQUINA
    )
)

if %opJogador% == TESOURA (
    if %OpMaquina% == PAPEL (
        set res=JOGADOR
    )
)

if %opJogador% == TESOURA (
    if %OpMaquina% == TESOURA (
        set res=EMPATE  
    )
)

if %opJogador% == TESOURA (
    if %OpMaquina% == LARGATO (
        set res=JOGADOR
    )
)

if %opJogador% == TESOURA (
    if %OpMaquina% == SPOCK (
        set res=MAQUINA
    )
)



if %opJogador% == LARGATO (
    if %OpMaquina% == PEDRA (
        set res=MAQUINA
    )
)

if %opJogador% == LARGATO (
    if %OpMaquina% == PAPEL (
        set res=JOGADOR
    )
)

if %opJogador% == LARGATO (
    if %OpMaquina% == TESOURA (
        set res=MAQUINA  
    )
)

if %opJogador% == LARGATO (
    if %OpMaquina% == LARGATO (
        set res=EMPATE
    )
)

if %opJogador% == LARGATO (
    if %OpMaquina% == SPOCK (
        set res=JOGADOR
    )
)



if %opJogador% == SPOCK (
    if %OpMaquina% == PEDRA (
        set res=JOGADOR
    )
)

if %opJogador% == SPOCK (
    if %OpMaquina% == PAPEL (
        set res=MAQUINA
    )
)

if %opJogador% == SPOCK (
    if %OpMaquina% == TESOURA (
        set res=JOGADOR  
    )
)

if %opJogador% == SPOCK (
    if %OpMaquina% == LARGATO (
        set res=MAQUINA
    )
)

if %opJogador% == SPOCK (
    if %OpMaquina% == SPOCK (
        set res=JOGADOR
    )
)

echo %nome% escolheu %opJogador%
echo mamaco escolheu %OpMaquina%

if %res% == JOGADOR (
    echo %nome% Ganhou!!!
) else (
    if %res% == MAQUINA (
        echo %nome% Perdeu!!!
    ) else (
        echo EMPATE
    )
)
pause > nul
goto inicio