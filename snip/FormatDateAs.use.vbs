Function FormatDateAs(i,f,z):Dim d,dp,m,mp,y,yp:dp=Len(f)-Len(Replace(f,"d","")):mp=Len(f)-Len(Replace(f,"m","")):yp=Len(f)-Len(Replace(f,"y","")):f=Replace(Replace(Replace(f,String(dp,"d"),"d"),String(mp,"m"),"m"),String(yp,"y"),"y"):If(z)Then:z="000":Else:z="":End If:d=Right(z&DatePart("d",i),dp):m=Right(z&DatePart("m",i),mp):y=Right(z&DatePart("yyyy",i),yp):f=Replace(Replace(Replace(f,"d",d),"m",m),"y",y):FormatDateAs=f:End Function

REM example with today's date zero padded
dim testExample
testExample = FormatDateAs( Date(), "yyyy\mm\dd", True )
MsgBox testExample, vbInformation, "FormatDateAs() Output:"

REM used to switch between zero and no padding
Dim padZeroSwitch
padZeroSwitch = False

REM more examples with pad zero turned on and off
Call TestExamples( Date(), "dd.mm.yyyy" )
Call TestExamples( #01/02/2003#, "mm/d/yy" )
Call TestExamples( #Jan 4, 2010#, "yyyy-mm-dd" )
Call TestExamples( #25 April 1992#, "d:yy:m" )
Call TestExamples( "2017/03/14", "yyyy\mmmm\dddd" )

REM looks better if run with cscript.exe
Sub TestExamples( dInput, dFormat )  
  padZeroSwitch = Not padZeroSwitch
  
  REM get formatted date
  dOutput = FormatDateAs( dInput, (dFormat), (padZeroSwitch) )
    
  REM show the user
  WScript.Echo _
    "Original :" & dInput & vbLf & _
    "Format   :" & dFormat & vbLf & _
    "Padded   :" & padZeroSwitch & vbLf & _
    "Output   :" & dOutput & vbLf & _
    "--------------------------------------"    
  
  REM show again, but without zero padding
  If( padZeroSwitch )Then    
    Call TestExamples( dInput, dFormat ) 
    WScript.Echo vbLf   
  End If
End Sub
