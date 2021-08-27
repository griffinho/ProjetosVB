dim Palavra(10), p1,audio,a,p,v,pon
call carregar_audio
sub carregar_audio()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1 'Velocidade da fala
call sortear_palavra
end sub

sub sortear_palavra()

palavra(1)="azul"
palavra(2)="bolo"
palavra(3)="pé"
palavra(4)="sal"
palavra(5)="bolo"
palavra(6)="sol"
palavra(7)="lua"
palavra(8)="ovo"
palavra(9)="eu"
palavra(10)="dia"

msgbox("Nivel 1"), vbinformation+vbokony, "Soletrando" 

pon=0
v=1
a=0

do while a<4
		
		randomize(second(second (time)))
		p=int (rnd *10)+1 
		audio.speak(palavra(p))
		
		opcao=msgbox("pular palavra?",vbyesno,"ATENÇÃO")
		
		if opcao=vbyes then 
		do while a<4
		
		randomize(second(second (time)))
		p=int (rnd *10)+1 
		audio.speak(palavra(p))
		
		p1=(inputbox("Digite a palavra","ATENC ÃO"))
                     
					 
		 if p1=(palavra(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		a=a+1
		pon=pon+50
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-50
		
		end if
		
		randomize(second(second (time)))
		p=int (rnd *9)+1 
		audio.speak(palavra(p))
		
		p1=(inputbox("Digite a palavra","ATENC ÃO"))
                     
					 
		 if p1=(palavra(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		a=a+1
		pon=pon+50
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-50
		
		end if
		
		randomize(second(second (time)))
		p=int (rnd *7)+1 
		audio.speak(palavra(p))
		
		
		p1=(inputbox("Digite a palavra","ATENCÃO"))
            
					 
		 if p1=(palavra(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		a=a+1
		pon=pon+50
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-50
		
		end if
		
		randomize(second(second (time)))
		p=int (rnd *5)+1
		audio.speak(palavra(p))
		
		p1=(inputbox("Digite a palavra","ATENC ÃO"))
                     
					 
		 if p1=(palavra(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		a=a+1
		pon=pon+50
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-50
		
		end if
		loop
		
		else
		
	
		p1=(inputbox("Digite a palavra", "ATENÇAÕ") )
		if p1=(palavra(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa") 
		a=a+1
		pon=pon+50
					 
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-50
		end if
		
		if v=0 then
		msgbox("Sua pontuação foi: "& pon &""), vbinformation+vbokony, "VOCÊ PERDEU" 
		wscript.quit 'Encerrar o programa
		end if
			
		end if

loop
end sub


dim PalavraN2(10),p2,audioN2,aN2,pN2
call carregar_audioN2
sub carregar_audioN2()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1 'Velocidade da fala
call sortear_palavraN2
end sub

sub sortear_palavraN2()

palavraN2(1)="andar"
palavraN2(2)="chuva"
palavraN2(3)="tarde"
palavraN2(4)="passos"
palavraN2(5)="perna"
palavraN2(6)="rato"
palavraN2(7)="coelho"
palavraN2(8)="sacola"
palavraN2(9)="sapo"
palavraN2(10)="noite"

msgbox("Nivel 2"), vbinformation+vbokony, "Soletrando"
aN2=0	
do while aN2<4
		randomize(second(second (time)))
		p=int (rnd *10)+1 
		audio.speak(palavraN2(p))
		
		opcao=msgbox("pular palavra?",vbyesno,"ATENÇÃO")
		
		if opcao=vbyes then 
		 
		do while aN2<4
		randomize(second(second (time)))
		p=int (rnd *10)+1 
		audio.speak(palavraN2(p))
		
		p2=(inputbox("Digite a palavra","ATENCÃO"))
                     
					 
		if p2=(palavraN2(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		aN2=aN2+1
		pon=pon+100
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-100
		
		end if

		randomize(second(second (time)))
		p=int (rnd *9)+1 
		audio.speak(palavraN2(p))
		
		p2=(inputbox("Digite a palavra","ATENCÃO"))
                     
					 
		if p2=(palavraN2(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		aN2=aN2+1
		pon=pon+100
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-100
		
		end if
		randomize(second(second (time)))
		p=int (rnd *7)+1 
		audio.speak(palavraN2(p))
		
		p2=(inputbox("Digite a palavra","ATENCÃO"))
                     
					 
		if p2=(palavraN2(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		aN2=aN2+1
		pon=pon+100
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-100
		
		end if
		randomize(second(second (time)))
		p=int (rnd *5)+1 
		audio.speak(palavraN2(p))
		
		p2=(inputbox("Digite a palavra","ATENCÃO"))
                     
					 
		if p2=(palavraN2(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		aN2=aN2+1
		pon=pon+100
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-100
		
		end if
		
		loop
		else
		
	
		p2=(inputbox("Digite a palavra","ATENÇÃO") )
		if p2=(palavraN2(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa") 
		aN2=aN2+1
		pon=pon+100
					 
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-100
		end if
		
		if v=0 then
		msgbox("Sua pontuação foi: "& pon &""), vbinformation+vbokony, "VOCÊ PERDEU" 
		wscript.quit 'Encerrar o programa
		end if
			
		end if
		
		
		
loop
end sub

dim PalavraN3(10),p3,audioN3,aN3,pN3
call carregar_audioN3
sub carregar_audioN3()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1 'Velocidade da fala
call sortear_palavraN3
end sub

sub sortear_palavraN3()

palavraN3(1)="abacaxi"
palavraN3(2)="pessego"
palavraN3(3)="jardim"
palavraN3(4)="computador"
palavraN3(5)="plantar"
palavraN3(6)="celular"
palavraN3(7)="camiseta"
palavraN3(8)="maquina"
palavraN3(9)="ventilador"
palavraN3(10)="apagador"

msgbox("Nivel 3"), vbinformation+vbokony, "Soletrando"

aN3=0
do while aN3<4
		
		randomize(second(second (time)))
		p=int (rnd *10)+1 
		audio.speak(palavraN3(p))
		
		opcao=msgbox("pular palavra?",vbyesno,"ATENÇÃO")
		
		if opcao=vbyes then 
		
		do while aN3<4
		randomize(second(second (time)))
		p=int (rnd *10)+1 
		audio.speak(palavraN3(p))
		
		p3=(inputbox("Digite a palavra","ATENCÃO"))
                     
					 
		if p3=(palavraN3(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		aN3=aN3+1
		pon=pon+500 
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-500
		
		end if
		loop
		else
		
	
		p3=(inputbox("Digite a palavra", "ATENÇAO") )
		
		if p3=(palavraN3(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa") 
		aN3=aN3+1
		pon=pon+500
					 
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-500
		end if
		
		if v=0 then
		msgbox("Sua pontuação foi: "& pon &""), vbinformation+vbokony, "VOCÊ PERDEU" 
		wscript.quit 'Encerrar o programa
		end if
			
		end if
		
		
loop
end sub

dim PalavraN4(10),p4,audioN4,aN4,pN4
call carregar_audioN4
sub carregar_audioN4()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1 'Velocidade da fala
call sortear_palavraN4
end sub

sub sortear_palavraN4()

palavraN4(1)="fenecimento"
palavraN4(2)="admoesta"
palavraN4(3)="engodar"
palavraN4(4)="piscicopata"
palavraN4(5)="belicoso"
palavraN4(6)="hipopotamo"
palavraN4(7)="ardiloso"
palavraN4(8)="calunia"
palavraN4(9)="alcunha"
palavraN4(10)="alarido"

msgbox("Nivel 4"), vbinformation+vbokony, "Soletrando" 
aN4=0
do while aN4<4
	
		randomize(second(second (time)))
		p=int (rnd *10)+1 
		audio.speak(palavraN4(p))
		
		opcao=msgbox("pular palavra?",vbyesno,"ATENÇÃO")
		
		if opcao=vbyes then 
		
		do while aN4<4
		randomize(second(second (time)))
		p=int (rnd *10)+1 
		audio.speak(palavraN4(p))
		
		p4=(inputbox("Digite a palavra","ATENCÃO"))
                     
					 
		   if p4=(palavraN4(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa")
		aN4=aN4+1
		pon=pon+1000
		
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-1000
		
		end if
		loop
		else
		
	
		p4=(inputbox("Digite a palavra", "ATENÇAO") )
		if p4=(palavraN4(p))then
		msgbox("Resposta certa"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta certa") 
		aN4=aN4+1
		pon=pon+1000
					 
		else
		msgbox("Resposta errada"), vbinformation+vbokony, "ATENÇAO" 
		audio.speak("Resposta errada")
		v=v-1
		pon=pon-1000
		end if
		
		if v=0 then
		msgbox("Sua pontuação foi: "& pon &""), vbinformation+vbokony, "VOCÊ PERDEU" 
		wscript.quit 'Encerrar o programa
		end if
			
		end if
		
		
		
		
loop
end sub

msgbox("PARABENS VOCÊ GANHOU O SOLETRANDO SUA PONTUAÇÃO FOI: "& pon &""), vbinformation+vbokony, "ATENÇÃO" 