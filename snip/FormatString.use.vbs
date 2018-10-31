REM Copy and paste class or just @mini into your code to use.
REM You can even set @mini to a user environment variable with 
REM @mini as the value, and then import it like so:
REM Execute(CreateObject("WScript.Shell").Environment("User")("ENV_VAR_NAME_HERE"))

Option Explicit      
Dim name, age
name = "Jeremy"
age = 97

Class FormatString:Private l,r,n,q:Sub Class_Initialize:Tokens="{}":n=vbLf:q="""":End Sub:Property Let Tokens(s):l=Left(s,1):r=Right(s,1):End Property:Property Let Append(f,s):f=f&Format(s):End Property:Public Default Function Format(s):Dim p,d:Format=s:p=InStr(1,Format,l)+1:If(p=1)Then:Exit Function:End If:d=Mid(Format,p,InStr(p,Format,r)-p):Format=Format(Replace(Format,(l&d&r),Eval(d))):End Function:End Class
Dim F : Set F = New FormatString

msgbox F.Format("Welcome to the FormatString Class.")

Dim msg, title
title = F("Hello {name}. Happy {age}th birthday!")

F.Append(msg) = "{name} is {age} years old today.{n}"
F.Append(msg) = "He's been an adult for {age - 18} years now."
F.Append(msg) = "{n}{q}Wow that's old!{q}, said Kate."

msgbox msg, vbOkOnly, title

F.Tokens = ";"
msgbox F("Tokens can be changed.;n;So you can include {brackets} in your message.")
