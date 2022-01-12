Rem EDiv$(Numerator, Denominator, Number of Decimal Places)
Function EDiv$ (numTbDiv, numDiviTb, numDeciplc)

    ObtedRest = numTbDiv \ numDiviTb
    rremainber = numTbDiv - (numDiviTb * ObtedRest)
    ObtedRest$ = _Trim$(Str$(ObtedRest))
    If rremainber = 0 GoTo extDivison

    ObtedRest$ = ObtedRest$ + "."
    numDeciplcS = 0
    Do
        If rremainber < numDiviTb Then rremainber = rremainber * 10
        If rremainber > numDiviTb Then
            ObtedRest = rremainber \ numDiviTb: numDeciplcS = numDeciplcS + 1
            ObtedRest$ = ObtedRest$ + _Trim$(Str$(ObtedRest))
            rremainber = rremainber - (numDiviTb * ObtedRest): If rremainber = 0 GoTo extDivison
        Else
            ObtedRest$ = ObtedRest$ + "0"
        End If
    Loop Until numDeciplcS = numDeciplc

    extDivison: EDiv$ = ObtedRest$

End Function

Rem EMult("number 1"|str$(num1), "number 2"|str$(num2))
Function EMult$ (mult1num$, mult2num$)

    mult1num$ = _Trim$(mult1num$): mult2num$ = _Trim$(mult2num$)

    'For Decimals
    demult2num = InStr(mult2num$, "."): demult1num = InStr(mult1num$, ".")
    dimult2num = Len(mult2num$) - demult2num: dimult1num = Len(mult1num$) - demult1num
    If demult1num = 0 Then dimult1num = 0: If demult2num = 0 Then dimult2num = 0
    DigDeciPlcs = dimult2num + dimult1num
    mult1num$ = Left$(mult1num$, demult1num - 1) + Right$(mult1num$, Len(mult1num$) - demult1num)
    mult2num$ = Left$(mult2num$, demult2num - 1) + Right$(mult2num$, Len(mult2num$) - demult2num)

    'Main Start
    If Len(mult1num$) < Len(mult2num$) Then
        rmdmulvar$ = mult2num$: rndmulvar$ = mult1num$
        mult1num$ = rmdmulvar$: mult2num$ = rndmulvar$
    End If

    For j = Len(mult2num$) To 1 Step -1
        lstmult2num = Val(Mid$(mult2num$, j, 1))

        For i = Len(mult1num$) To 1 Step -1
            lstmult1num = Val(Mid$(mult1num$, i, 1))
            sObtedRest$ = _Trim$(Str$((lstmult2num * lstmult1num) + crryFrmult))
            ObtedRest$ = Right$(sObtedRest$, 1) + ObtedRest$: crryFrmult = Val(Left$(sObtedRest$, Len(sObtedRest$) - 1))
        Next
        If crryFrmult <> 0 Then ObtedRest$ = _Trim$(Str$(crryFrmult)) + ObtedRest$
        ObtedRest$ = ObtedRest$ + String$(Len(mult2num$) - j, "0"): k = k + 1


        If k = 2 Then
            ObtedPrdct$ = EAdd$(pObtedRest$, ObtedRest$)
        ElseIf k < 2 Then
            pObtedRest$ = ObtedRest$: ObtedPrdct$ = pObtedRest$
        Else
            ObtedPrdct$ = EAdd$(ObtedPrdct$, ObtedRest$)
        End If
        ObtedRest$ = "": crryFrmult = 0

    Next


    If DigDeciPlcs <> 0 Then
        ObtedPrdct$ = String$(DigDeciPlcs - Len(ObtedPrdct$), "0") + ObtedPrdct$
        ObtedPrdct$ = Left$(ObtedPrdct$, Len(ObtedPrdct$) - DigDeciPlcs) + "." + Right$(ObtedPrdct$, DigDeciPlcs)
    End If

    EMult$ = ObtedPrdct$

End Function

Rem EAdd("number 1"|str$(num1), "number 2"|str$(num2))
Function EAdd$ (mult1num$, mult2num$)

    mult1num$ = _Trim$(mult1num$): mult2num$ = _Trim$(mult2num$)


    'case 2, both rational numbers
    If InStr(mult1num$, ".") <> 0 And InStr(mult2num$, ".") <> 0 Then
        dmult1num$ = Mid$(mult1num$, InStr(mult1num$, ".") + 1): ledmult1num = Len(dmult1num$)
        dmult2num$ = Mid$(mult2num$, InStr(mult2num$, ".") + 1): ledmult2num = Len(dmult2num$)
        If ledmult2num < ledmult1num Then
            ObtedRest$ = Right$(dmult1num$, ledmult1num - ledmult2num)
            dmult1num$ = Left$(dmult1num$, ledmult2num)
        Else
            ObtedRest$ = Right$(dmult2num$, ledmult2num - ledmult1num)
            dmult2num$ = Left$(dmult2num$, ledmult1num)
        End If
        cmult1num$ = Left$(mult1num$, InStr(mult1num$, ".") - 1): cmult2num$ = Left$(mult2num$, InStr(mult2num$, ".") - 1)
        mult1num$ = dmult1num$: mult2num$ = dmult2num$: ContDecipnt = 1


        'case 3, any one is rational
    ElseIf InStr(mult1num$, ".") <> 0 Or InStr(mult2num$, ".") <> 0 Then
        dmult1num$ = Mid$(mult1num$, InStr(mult1num$, ".")): dmult2num$ = Mid$(mult2num$, InStr(mult2num$, "."))
        If dmult1num$ = mult1num$ Then
            ObtedRest$ = dmult2num$
            mult2num$ = Left$(mult2num$, InStr(mult2num$, ".") - 1)
        Else
            ObtedRest$ = dmult1num$
            mult1num$ = Left$(mult1num$, InStr(mult1num$, ".") - 1)
        End If
    End If


    'case 1, both integer
    Add: Do
        lstmult1num = Val(Right$(mult1num$, 1))
        lstmult2num = Val(Right$(mult2num$, 1))
        SabkaSum$ = _Trim$(Str$(lstmult1num + lstmult2num + crryFrmult))
        ObtedRest$ = Right$(SabkaSum$, 1) + ObtedRest$: crryFrmult = Val(Left$(SabkaSum$, Len(SabkaSum$) - 1))
        mult1num$ = Left$(mult1num$, Len(mult1num$) - 1): mult2num$ = Left$(mult2num$, Len(mult2num$) - 1)
    Loop Until mult1num$ = "" And mult2num$ = ""


    If ContDecipnt = 1 Then
        ContDecipnt = -1: mult1num$ = cmult1num$: mult2num$ = cmult2num$
        ObtedRest$ = "." + ObtedRest$
        GoTo Add
    Else
        If crryFrmult <> 0 Then ObtedRest$ = _Trim$(Str$(crryFrmult)) + ObtedRest$
    End If

    EAdd$ = ObtedRest$

End Function

Rem ESub("number 1"|str$(num1), "number 2"|str$(num2))
Function ESub$ (mult1num$, mult2num$)

    mult1num$ = _Trim$(mult1num$): mult2num$ = _Trim$(mult2num$)


    'case 2, both rational numbers
    If InStr(mult1num$, ".") <> 0 And InStr(mult2num$, ".") <> 0 Then
        dmult1num$ = Mid$(mult1num$, InStr(mult1num$, ".") + 1): ledmult1num = Len(dmult1num$)
        dmult2num$ = Mid$(mult2num$, InStr(mult2num$, ".") + 1): ledmult2num = Len(dmult2num$)
        If ledmult2num < ledmult1num Then
            dmult2num$ = dmult2num$ + String$(ledmult1num - ledmult2num, "0")
            'dmult1num$ = Left$(dmult1num$, ledmult2num)
        Else
            dmult1num$ = dmult1num$ + String$(ledmult2num - ledmult1num, "0")
            'dmult2num$ = Left$(dmult2num$, ledmult1num)
        End If
        cmult1num$ = Left$(mult1num$, InStr(mult1num$, ".") - 1): cmult2num$ = Left$(mult2num$, InStr(mult2num$, ".") - 1)
        mult1num$ = dmult1num$: mult2num$ = dmult2num$: ContDecipnt = 1


        'case 3, any one is rational
    ElseIf InStr(mult1num$, ".") <> 0 Or InStr(mult2num$, ".") <> 0 Then
        dmult1num$ = Mid$(mult1num$, InStr(mult1num$, ".")): dmult2num$ = Mid$(mult2num$, InStr(mult2num$, "."))
        If dmult1num$ = mult1num$ Then
            dmult1num$ = String$(Len(dmult2num$) - 1, "0"): dmult2num$ = Mid$(dmult2num$, 2)
            cmult2num$ = Left$(mult2num$, InStr(mult2num$, ".") - 1): cmult1num$ = mult1num$
        Else
            dmult2num$ = String$(Len(dmult1num$) - 1, "0"): dmult1num$ = Mid$(dmult1num$, 2)
            cmult1num$ = Left$(mult1num$, InStr(mult1num$, ".") - 1): cmult2num$ = mult2num$
        End If
        mult1num$ = dmult1num$: mult2num$ = dmult2num$: ContDecipnt = 1
    End If


    'case 1, both integer
    Subt:
    If ContDecipnt = 1 Then
        varmult1num$ = cmult1num$: varmult2num$ = cmult2num$
    Else
        varmult1num$ = mult1num$: varmult2num$ = mult2num$
    End If

    If Len(varmult1num$) = Len(varmult2num$) Then
        For i = 1 To Len(varmult1num$)
            mult1num = Val(Mid$(varmult1num$, i, 1)): mult2num = Val(Mid$(varmult2num$, i, 1))
            If mult1num = mult2num Then _Continue
            EkSaman = -1: Exit For
        Next

        If EkSaman = -1 Then
            If mult1num < mult2num GoTo sign
        Else
            If ContDecipnt = 1 Then
                For i = 1 To Len(mult1num$)
                    mult1num = Val(Mid$(mult1num$, i, 1)): mult2num = Val(Mid$(mult2num$, i, 1))
                    If mult1num = mult2num Then _Continue
                    EkSaman = -1: Exit For
                Next

                If EkSaman = -1 Then
                    If mult1num < mult2num GoTo sign
                Else
                    ESub$ = "0": Exit Function
                End If
            Else
                ESub$ = "0": Exit Function
            End If
        End If

    ElseIf Len(varmult1num$) < Len(varmult2num$) Then
        sign: nmult2num$ = mult1num$: mmult1num$ = mult2num$
        mult1num$ = mmult1num$: mult2num$ = nmult2num$
        sigmult2num$ = "-"
    End If


    Do
        lstmult1num = Val(Right$(mult1num$, 1)) + crryFrmult
        lstmult2num = Val(Right$(mult2num$, 1)): crryFrmult = 0
        If lstmult1num < lstmult2num Then
            lstmult1num = lstmult1num + 10: crryFrmult = -1
        End If
        SabkaSub$ = _Trim$(Str$(lstmult1num - lstmult2num))
        ObtedRest$ = SabkaSub$ + ObtedRest$
        mult1num$ = Left$(mult1num$, Len(mult1num$) - 1): mult2num$ = Left$(mult2num$, Len(mult2num$) - 1)
    Loop Until mult1num$ = "" And mult2num$ = ""


    If ContDecipnt = 1 Then
        ContDecipnt = -1: mult1num$ = cmult1num$: mult2num$ = cmult2num$
        ObtedRest$ = "." + ObtedRest$
        GoTo Subt
    Else
        If crryFrmult <> 0 Then ObtedRest$ = _Trim$(Str$(crryFrmult)) + ObtedRest$
    End If

    ESub$ = sigmult2num$ + ObtedRest$

End Function

Rem EFact(number)
Function EFact$ (numfurFact)

    If numfurFact < 0 Then Exit Function

    ObtedRest$ = "1"
    For i = numfurFact To 1 Step -1
        ObtedRest$ = EMult(ObtedRest$, Str$(i))
    Next

    EFact$ = ObtedRest$

End Function
