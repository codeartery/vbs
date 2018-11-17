# Glossary
A collection of classes and functions that I have made.

---
### Classes
- [FormatString](./FormatString.vbs), [example](./FormatString.use.vbs)

---
### Functions
- [FormatDateAs](./FormatDateAs.vbs), [example](./FormatDateAs.use.vbs)

---
### @Mini
Each class or function has a @mini version that can be used to reduce the space the code takes up, but also can be assigned to an environment variable for easy importation into any of your scripts.
#### How to assign @mini to an environment variable
```ini
KEY:
<> = title of window
{} = keyboard keys
[] = button, or input
() = copy-paste this value.
\\ = value inside is not raw; it represents something else
__ = type something here
```

1. {WINDOWS_KEY + R}
2. <Run>
   [Open:] (rundll32 sysdm.cpl,EditEnvironmentVariables)  
   [OK]  
3. <Environment Variables> 
   [New...]  
4. <New User Variable>
   [Variable name:] _NAME_HERE_  
   [Variable value:] (\@Mini\)  
   [OK]  
5. <Environment Variables>
   [OK]
