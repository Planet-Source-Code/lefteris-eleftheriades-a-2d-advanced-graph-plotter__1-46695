Attribute VB_Name = "Module1"
Function Sec(ByVal X As Currency) As Currency: Sec = 1 / Cos(X): End Function
Function Cosec(ByVal X As Currency) As Currency: Cosec = 1 / Sin(X): End Function
Function Cotan(ByVal X As Currency) As Currency: Cotan = 1 / Tan(X): End Function
Function Arcsin(ByVal X As Currency) As Currency: Arcsin = Atn(X / Sqr(-X * X + 1)): End Function
Function Arccos(ByVal X As Currency) As Currency: Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1): End Function
Function Arcsec(ByVal X As Currency) As Currency: Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1)): End Function
Function Arccosec(ByVal X As Currency) As Currency: Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1)): End Function
Function Arccotan(ByVal X As Currency) As Currency: Arccotan = Atn(X) + 2 * Atn(1): End Function
Function HSin(ByVal X As Currency) As Currency: HSin = (Exp(X) - Exp(-X)) / 2: End Function
Function HCos(ByVal X As Currency) As Currency: HCos = (Exp(X) + Exp(-X)) / 2: End Function
Function HTan(ByVal X As Currency) As Currency: HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X)): End Function
Function HSec(ByVal X As Currency) As Currency: HSec = 2 / (Exp(X) + Exp(-X)): End Function
Function HCosec(ByVal X As Currency) As Currency: HCosec = 2 / (Exp(X) - Exp(-X)): End Function
Function HCotan(ByVal X As Currency) As Currency: HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X)): End Function
Function HArcsin(ByVal X As Currency) As Currency: HArcsin = Log(X + Sqr(X * X + 1)): End Function
Function HArccos(ByVal X As Currency) As Currency: HArccos = Log(X + Sqr(X * X - 1)): End Function
Function HArctan(ByVal X As Currency) As Currency: HArctan = Log((1 + X) / (1 - X)) / 2: End Function
Function HArcsec(ByVal X As Currency) As Currency: HArcsec = Log((Sqr(-X * X + 1) + 1) / X): End Function
Function HArccosec(ByVal X As Currency) As Currency: HArccosec = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X): End Function
Function HArccotan(ByVal X As Currency) As Currency: HArccotan = Log((X + 1) / (X - 1)) / 2: End Function
Function Ð() As Currency: Ð = 3.14159265358979: End Function
Function Pi() As Currency: Pi = 3.14159265358979: End Function
Function Drad() As Currency: Drad = 1.74532925199433E-02: End Function
Function Rdeg() As Currency: Rdeg = 57.2957795130823: End Function
Function LogN(ByVal B As Currency, ByVal r As Currency) As Currency
     If r > 0 And B > 0 Then LogN = Log(r) / Log(B)
End Function

