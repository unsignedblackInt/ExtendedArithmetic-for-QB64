'Dim a As _Unsigned _Integer64
'Print ExtDiv$(1, 998001, 200)
'a = 1 / 998001
'Print a
'Print ExtFact(20)
'Print ExtAdd(ExtAdd("8", "8"), "800")
Print ExtMult("56890664327653534689", "56890664327653534689")

Function ExtDiv$ (n, d, f)
    n$ = Str$(n): d$ = Str$(d)
    res = n \ d
    r = n - (d * res)
    res$ = _Trim$(Str$(res))
    If r = 0 GoTo extDivison

    res$ = res$ + "."
    s = 0
    Do
        If r < d Then r = r * 10
        If r > d Then
            res = r \ d: s = s + 1
            res$ = res$ + _Trim$(Str$(res))
            r = r - (d * res): If r = 0 GoTo extDivison
        Else
            res$ = res$ + "0"
        End If
    Loop Until s = f
    extDivison: ExtDiv$ = res$
End Function

Function ExtMult$ (m$, n$)
    m$ = _Trim$(m$): n$ = _Trim$(n$)
    If Len(m$) < Len(n$) Then
        mm$ = n$: nn$ = m$
        m$ = mm$: n$ = nn$
    End If
    Print n$, m$
    For j = Len(n$) To 1 Step -1
        l_n = Val(Mid$(n$, j, 1))
        For i = Len(m$) To 1 Step -1
            l_d = Val(Mid$(m$, i, 1))
            sres$ = _Trim$(Str$((l_n * l_d) + c))
            res$ = Right$(sres$, 1) + res$: c = Val(Left$(sres$, Len(sres$) - 1))
            '?res$,sres$:sleep
        Next
        If c <> 0 Then res$ = _Trim$(Str$(c)) + res$
        res$ = res$ + String$(Len(n$) - j, "0"): k = k + 1
        'Print res$, Len(n$) - j: Sleep
        If k = 2 Then
            prdct$ = ExtAdd$(pres$, res$)
        ElseIf k < 2 Then
            pres$ = res$: prdct$ = pres$
        Else
            prdct$ = ExtAdd$(prdct$, res$)
        End If
        res$ = "": c = 0
    Next
    ExtMult$ = prdct$
End Function

Function ExtAdd$ (m$, n$)
    m$ = _Trim$(m$): n$ = _Trim$(n$)

    'case 2, both rational numbers
    If InStr(m$, ".") <> 0 And InStr(n$, ".") <> 0 Then
        dm$ = Mid$(m$, InStr(m$, ".") + 1): ldm = Len(dm$)
        dn$ = Mid$(n$, InStr(n$, ".") + 1): ldn = Len(dn$)
        If ldn < ldm Then
            res$ = Right$(dm$, ldm - ldn)
            dm$ = Left$(dm$, ldn)
        Else
            res$ = Right$(dn$, ldn - ldm)
            dn$ = Left$(dn$, ldm)
        End If
        cm$ = Left$(m$, InStr(m$, ".") - 1): cn$ = Left$(n$, InStr(n$, ".") - 1)
        m$ = dm$: n$ = dn$: d = 1

        'case 3, any one is rational
    ElseIf InStr(m$, ".") <> 0 Or InStr(n$, ".") <> 0 Then
        dm$ = Mid$(m$, InStr(m$, ".")): dn$ = Mid$(n$, InStr(n$, "."))
        If dm$ = m$ Then
            res$ = dn$
            n$ = Left$(n$, InStr(n$, ".") - 1)
        Else
            res$ = dm$
            m$ = Left$(m$, InStr(m$, ".") - 1)
        End If
    End If

    'case 1, both integer
    Add: Do
        l_m = Val(Right$(m$, 1))
        l_n = Val(Right$(n$, 1))
        sum$ = _Trim$(Str$(l_m + l_n + c))
        res$ = Right$(sum$, 1) + res$: c = Val(Left$(sum$, Len(sum$) - 1))
        m$ = Left$(m$, Len(m$) - 1): n$ = Left$(n$, Len(n$) - 1)
    Loop Until m$ = "" And n$ = ""
    If d = 1 Then
        d = -1: m$ = cm$: n$ = cn$
        res$ = "." + res$
        GoTo Add
    Else
        If c <> 0 Then res$ = _Trim$(Str$(c)) + res$
    End If

    ExtAdd$ = res$

End Function

Function ExtFact (n)
    f = 1
    For i = n To 1 Step -1
        f = f * i
    Next
    ExtFact = f
End Function
