Function FormatDateAs( dInput, sFormat )
  REM @Auth::Jeremy England( CodeArtery )
  REM @Info::Format a date based off of a user defined 'd','m','y' pattern.
  REM @Mini::Function FormatDateAs(i,f):Dim d,dp,m,mp,y,yp:dp=Len(f)-Len(Replace(f,"d","")):mp=Len(f)-Len(Replace(f,"m","")):yp=Len(f)-Len(Replace(f,"y","")):f=Replace(Replace(Replace(f,String(dp,"d"),"d"),String(mp,"m"),"m"),String(yp,"y"),"y"):d=Right("000"&DatePart("d",i),dp):m=Right("000"&DatePart("m",i),mp):y=Right("000"&DatePart("yyyy",i),yp):f=Replace(Replace(Replace(f,"d",d),"m",m),"y",y):FormatDateAs=f:End Function  
  Dim dPrt, dPad, mPrt, mPad, yPrt, yPad
  dPad = Len( sFormat ) - Len( Replace( sFormat, "d", "") )
  mPad = Len( sFormat ) - Len( Replace( sFormat, "m", "") )
  yPad = Len( sFormat ) - Len( Replace( sFormat, "y", "") )

  sFormat = Replace( sFormat, String(dPad,"d"), "d")
  sFormat = Replace( sFormat, String(mPad,"m"), "m")
  sFormat = Replace( sFormat, String(yPad,"y"), "y")

  dPrt = Right( "0000" & DatePart( "d", dInput ), dPad)
  mPrt = Right( "0000" & DatePart( "m", dInput ), mPad)
  yPrt = Right( "0000" & DatePart( "yyyy", dInput ), yPad)

  sFormat = Replace( sFormat, "d", dPrt )
  sFormat = Replace( sFormat, "m", mPrt )
  sFormat = Replace( sFormat, "y", yPrt )
  FormatDateAs = sFormat
End Function
