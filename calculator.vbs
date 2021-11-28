num1 = inputbox("Enter First Number")
equ = inputbox("enter your variable ie. (/ * - +)")
num2 = inputbox("Enter Second Number")

Set voice = CreateObject("SAPI.SpVoice")
voice.Rate = 1
voice.Volume = 90
Set voice.Voice = voice.GetVoices.Item(0)


if equ = "+" then voice.Speak(num1 -- num2) : msgbox(num1 -- num2)  

if equ = "-" then voice.Speak(num1 - num2) : msgbox(num1 - num2) 

if equ = "*" then voice.Speak(num1 * num2) : msgbox(num1 * num2) 

if equ = "/" then voice.Speak(num1 / num2) : msgbox(num1 / num2) 
