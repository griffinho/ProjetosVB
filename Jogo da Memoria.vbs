Option Explicit
dim  n1,n2,n3,audio,r,numeros(),i,n,list,numero,qt,resp,j
call carregar_audio
sub carregar_audio()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1 'Velocidade da fala
call sortear_numero
end sub
sub sortear_numero()

Set list = CreateObject("System.Collections.ArrayList")
msgbox("Jogo da memoria "),vbinformation+vbokonly
randomize(second(second (time)))
n1 = int (rnd *100)+1
audio.speak(n1)
list.Add n1

randomize(second(second (time)))
n2 = int (rnd *100)+1
audio.speak(n2)
list.Add n2

randomize(second(second (time)))
n3 = int (rnd *100)+1
audio.speak(n3)
list.Add n3

n = cstr (n1 & n2 & n3) 
r = cstr(inputbox("digite a sequencia de numeros :"))


qt = 0


do while r = n
	randomize(second(second (time)))
	n1 = int (rnd *100)+1
	n1 = n1 + 1
	list.add n1
	for each numero in list'descompactar lista
		audio.speak(numero)
		
	next
	
	n = n & n1
	r = cstr(inputbox("digite a sequencia de numero :"))
	if (r=n) then
		msgbox("Parabens voce acertou  "),vbinformation+vbokonly
	qt = qt + 1
	end if
loop
msgbox("voce errou "),vbinformation+vbokonly
msgbox("A sequencia de numeros foi :" &n),vbinformation+vbokonly

resp=msgbox("Deseja reiniciar ?",vbquestion+vbyesno,"ATENCAO")
if resp=vbyes then
	call sortear_numero
else
	wscript.quit 'Encerra o programa
end if


end sub
