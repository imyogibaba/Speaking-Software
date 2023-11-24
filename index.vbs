Dim Message, Speak
Message=InputBox("Welcome to VMACE.in speaking software Entertext","Speak")
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak Message
