'$Console:Only
'_Dest _Console
'Dim a As _Unsigned _Integer64
'Print EDiv$(1, 998001, 200)
'a = 1 / 998001
'Print a
'Do
'    i = i + 1
'    _Limit 5
'    Print EFact(i)
'Loop
'Print EAdd$(ExtAdd("8", "8"), "800")
'Print EMult("12345.67890", "12345678.90")

Function EDiv$ (n, d, f)
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
    extDivison: EDiv$ = res$
End Function

Function EMult$ (m$, n$)
    m$ = _Trim$(m$): n$ = _Trim$(n$)

    'For Decimals
    d_n = InStr(n$, "."): d_m = InStr(m$, ".")
    n_d = Len(n$) - d_n: m_d = Len(m$) - d_m
    If d_m = 0 Then m_d = 0: If d_n = 0 Then n_d = 0
    Decimal = n_d + m_d
    m$ = Left$(m$, d_m - 1) + Right$(m$, Len(m$) - d_m)
    n$ = Left$(n$, d_n - 1) + Right$(n$, Len(n$) - d_n)

    'Main Start
    If Len(m$) < Len(n$) Then
        mm$ = n$: nn$ = m$
        m$ = mm$: n$ = nn$
    End If
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
            prdct$ = EAdd$(pres$, res$)
        ElseIf k < 2 Then
            pres$ = res$: prdct$ = pres$
        Else
            prdct$ = EAdd$(prdct$, res$)
        End If
        res$ = "": c = 0
    Next
    If Decimal <> 0 Then
        prdct$ = String$(Decimal - Len(prdct$), "0") + prdct$
        prdct$ = Left$(prdct$, Len(prdct$) - Decimal) + "." + Right$(prdct$, Decimal)
    End If
    EMult$ = prdct$
End Function

Function EAdd$ (m$, n$)
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

    EAdd$ = res$

End Function

Function EFact$ (n)
    f$ = "1"
    For i = n To 1 Step -1
        f$ = EMult(f$, Str$(i))
    Next
    EFact$ = f$
End Function
