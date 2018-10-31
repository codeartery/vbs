Class FormatString
  REM @owner Jeremy England( CodeArtery )
  REM @info   Evals data in strings between {brackets} or user defined tokens, and allows for easy appending.
  REM @mini   Class FormatString:Private l,r,n,q:Sub Class_Initialize:Tokens="{}":n=vbLf:q="""":End Sub:Property Let Tokens(s):l=Left(s,1):r=Right(s,1):End Property:Property Let Append(f,s):f=f&Format(s):End Property:Public Default Function Format(s):Dim p,d:Format=s:p=InStr(1,Format,l)+1:If(p=1)Then:Exit Function:End If:d=Mid(Format,p,InStr(p,Format,r)-p):Format=Format(Replace(Format,(l&d&r),Eval(d))):End Function:End Class
  Private tokenL, tokenR, n, q

  Sub Class_Initialize
    Tokens = "{}"
    n = vbLf
    q = """"
  End Sub

  Property Let Tokens( str )
    tokenL = Left( str, 1 )
    tokenR = Right( str, 1 )
  End Property

  Property Let Append( ByRef ref, str )
    ref = ref & Format( str )
  End Property

  Public Default Function Format( str )
    Dim tPos, tData
    Format = str
    tPos = InStr( 1, Format, tokenL ) + 1
    If( tPos = 1 )Then Exit Function
    tData = Mid( Format, tPos, InStr( tPos, Format, tokenR ) - tPos )
    Format = Format( Replace( Format, (tokenL & tData & tokenR), Eval(tData) ) )
  End Function
End Class
