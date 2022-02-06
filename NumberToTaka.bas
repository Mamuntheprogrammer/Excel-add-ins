Attribute VB_Name = "NumberToTaka"
Function taka(ByVal N As Currency) As String
 
   Const Thousand = 1000@
   Const Lac = 100000@
   Const Crore = 10000000@
 
   If (N = 0@) Then BangladeshTaka = "Zero": Exit Function
 
   Dim Buf As String: If (N < 0@) Then Buf = "Negative " Else Buf = ""
   Dim Frac As Currency: Frac = Abs(N - Fix(N))
   If (N < 0@ Or Frac <> 0@) Then N = Abs(Fix(N))
   Dim AtLeastOne As Integer: AtLeastOne = N >= 1
 
   If (N >= Crore) Then
      Buf = Buf & BangladeshTakaDigitGroup(Int(N / Crore)) & " Crore"
      N = N - Int(N / Crore) * Crore
      If (N >= 1@) Then Buf = Buf & " "
   End If
 
   If (N >= Lac) Then
      Buf = Buf & BangladeshTakaDigitGroup(Int(N / Lac)) & " Lac"
      N = N - Int(N / Lac) * Lac
      If (N >= 1@) Then Buf = Buf & " "
   End If
 
   If (N >= Thousand) Then
      Buf = Buf & BangladeshTakaDigitGroup(N \ Thousand) & " Thousand"
      N = N Mod Thousand
      If (N >= 1@) Then Buf = Buf & " "
   End If
 
   If (N >= 1@) Then
      Buf = Buf & BangladeshTakaDigitGroup(N)
   End If
 
   If (Frac = 0@) Then
      Buf = Buf & " Taka Only"
       
   Else
      If AtLeastOne Then Buf = Buf & " Taka"
      Buf = Buf & " and " & BangladeshTakaDigitGroup(Frac * 100) & " Paisa Only"
   End If
 
   taka = Buf
End Function
 
Private Function BangladeshTakaDigitGroup(ByVal N As Integer) As String
   Const Hundred = " Hundred"
   Const One = "One"
   Const Two = "Two"
   Const Three = "Three"
   Const Four = "Four"
   Const Five = "Five"
   Const Six = "Six"
   Const Seven = "Seven"
   Const Eight = "Eight"
   Const Nine = "Nine"
   Dim Buf As String: Buf = ""
   Dim Flag As Integer: Flag = False
 
   Select Case (N \ 100)
      Case 0: Buf = "": Flag = False
      Case 1: Buf = One & Hundred: Flag = True
      Case 2: Buf = Two & Hundred: Flag = True
      Case 3: Buf = Three & Hundred: Flag = True
      Case 4: Buf = Four & Hundred: Flag = True
      Case 5: Buf = Five & Hundred: Flag = True
      Case 6: Buf = Six & Hundred: Flag = True
      Case 7: Buf = Seven & Hundred: Flag = True
      Case 8: Buf = Eight & Hundred: Flag = True
      Case 9: Buf = Nine & Hundred: Flag = True
   End Select
 
   If (Flag <> False) Then N = N Mod 100
   If (N > 0) Then
      If (Flag <> False) Then Buf = Buf & " "
   Else
      BangladeshTakaDigitGroup = Buf
      Exit Function
   End If
 
   Select Case (N \ 10)
      Case 0, 1: Flag = False
      Case 2: Buf = Buf & "Twenty": Flag = True
      Case 3: Buf = Buf & "Thirty": Flag = True
      Case 4: Buf = Buf & "Forty": Flag = True
      Case 5: Buf = Buf & "Fifty": Flag = True
      Case 6: Buf = Buf & "Sixty": Flag = True
      Case 7: Buf = Buf & "Seventy": Flag = True
      Case 8: Buf = Buf & "Eighty": Flag = True
      Case 9: Buf = Buf & "Ninety": Flag = True
   End Select
 
   If (Flag <> False) Then N = N Mod 10
   If (N > 0) Then
      If (Flag <> False) Then Buf = Buf & " "
   Else
      BangladeshTakaDigitGroup = Buf
      Exit Function
   End If
 
   Select Case (N)
      Case 0:
      Case 1: Buf = Buf & One
      Case 2: Buf = Buf & Two
      Case 3: Buf = Buf & Three
      Case 4: Buf = Buf & Four
      Case 5: Buf = Buf & Five
      Case 6: Buf = Buf & Six
      Case 7: Buf = Buf & Seven
      Case 8: Buf = Buf & Eight
      Case 9: Buf = Buf & Nine
      Case 10: Buf = Buf & "Ten"
      Case 11: Buf = Buf & "Eleven"
      Case 12: Buf = Buf & "Twelve"
      Case 13: Buf = Buf & "Thirteen"
      Case 14: Buf = Buf & "Fourteen"
      Case 15: Buf = Buf & "Fifteen"
      Case 16: Buf = Buf & "Sixteen"
      Case 17: Buf = Buf & "Seventeen"
      Case 18: Buf = Buf & "Eighteen"
      Case 19: Buf = Buf & "Nineteen"
   End Select
 
   BangladeshTakaDigitGroup = Buf
 
End Function

