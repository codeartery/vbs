Function FormatDateAs( dInput, sFormat, padZero )
  REM @auth::Jeremy England ( CodeArtery™ )
  REM @info::Format a date based off of a user defined 'd','m','y' pattern.
  REM @mini::Function FormatDateAs(i,f,z):Dim d,dp,m,mp,y,yp:dp=Len(f)-Len(Replace(f,"d","")):mp=Len(f)-Len(Replace(f,"m","")):yp=Len(f)-Len(Replace(f,"y","")):f=Replace(Replace(Replace(f,String(dp,"d"),"d"),String(mp,"m"),"m"),String(yp,"y"),"y"):If(z)Then:z="000":Else:z="":End If:d=Right(z&DatePart("d",i),dp):m=Right(z&DatePart("m",i),mp):y=Right(z&DatePart("yyyy",i),yp):f=Replace(Replace(Replace(f,"d",d),"m",m),"y",y):FormatDateAs=f:End Function
  Dim dPrt, dPad, mPrt, mPad, yPrt, yPad
  
  dPad = Len( sFormat ) - Len( Replace( sFormat, "d", "") )
  mPad = Len( sFormat ) - Len( Replace( sFormat, "m", "") )
  yPad = Len( sFormat ) - Len( Replace( sFormat, "y", "") )

  sFormat = Replace( sFormat, String(dPad,"d"), "d")
  sFormat = Replace( sFormat, String(mPad,"m"), "m")
  sFormat = Replace( sFormat, String(yPad,"y"), "y")
  
  If( padZero )Then padZero = "000" Else padZero = ""
  dPrt = Right( padZero & DatePart( "d", dInput ), dPad)
  mPrt = Right( padZero & DatePart( "m", dInput ), mPad)
  yPrt = Right( padZero & DatePart( "yyyy", dInput ), yPad)

  sFormat = Replace( sFormat, "d", dPrt )
  sFormat = Replace( sFormat, "m", mPrt )
  sFormat = Replace( sFormat, "y", yPrt )
  FormatDateAs = sFormat
End Function
