# Glossary
A collection of classes and functions that I have made.

---
### Classes
- [FormatString](./FormatString.vbs) - [Code example](./FormatString.use.vbs)

---
### Functions
- [FormatDateAs](./FormatDateAs.vbs) - [Code example](./FormatDateAs.use.vbs)

---
### @Mini
Each class or function has a @mini version that can be used to reduce the space the code takes up, but also can be assigned to an environment variable for easy importation into any of your scripts.
#### How to assign @mini to an environment variable
```
KEY:
<> = title of window
{} = keyboard keys
[] = button, or input
() = copy-paste this value.
// = value inside is not raw; it represents something else
__ = type something here
```

> {WinKey + R}  

> \<Run>  
>> [Open:] (rundll32 sysdm.cpl,EditEnvironmentVariables)   
>> [OK]  

> \<Environment Variables>  
>> [New...]  

> \<New User Variable>  
>> [Variable name:] _ENV_NAME_HERE_  
>> [Variable value:] (/@Mini/)  
>> [OK]   

> \<Environment Variables>  
>> [OK]  

#### How to import environment variable code into vbscript
```CS
Execute(CreateObject("WScript.Shell").Environment("User")("ENV_NAME_HERE"))

REM your code here...
```
