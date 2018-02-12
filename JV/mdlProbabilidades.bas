Attribute VB_Name = "mdlProbabilidades"
Option Explicit

Public Function ProbabilidadDeBatalla(intAtaqueInicio As Integer, intDefensaInicio As Integer, intAtaqueFin As Integer, intDefensaFin As Integer) As Double
'Calcula la probabilidad de que una Batalla (una sola
'tirada de dados) se inicie y finalice con los valores
'especificados

'A1-D1
If intAtaqueInicio = 1 And intDefensaInicio = 1 Then
    If intAtaqueFin = 0 And intDefensaFin = 1 Then
        ProbabilidadDeBatalla = 21 / 36
        Exit Function
    Else
        ProbabilidadDeBatalla = 15 / 36
        Exit Function
    End If
End If

'A2-D1
If intAtaqueInicio = 2 And intDefensaInicio = 1 Then
    If intAtaqueFin = 1 And intDefensaFin = 1 Then
        ProbabilidadDeBatalla = 91 / 216
        Exit Function
    Else
        ProbabilidadDeBatalla = 125 / 216
        Exit Function
    End If
End If

'A3-D1
If intAtaqueInicio = 3 And intDefensaInicio = 1 Then
    If intAtaqueFin = 2 And intDefensaFin = 1 Then
        ProbabilidadDeBatalla = 441 / 1296
        Exit Function
    Else
        ProbabilidadDeBatalla = 855 / 1296
        Exit Function
    End If
End If


'A1-D2
If intAtaqueInicio = 1 And intDefensaInicio = 2 Then
    If intAtaqueFin = 0 And intDefensaFin = 2 Then
        ProbabilidadDeBatalla = 161 / 216
        Exit Function
    Else
        ProbabilidadDeBatalla = 55 / 216
        Exit Function
    End If
End If


'A1-D3
If intAtaqueInicio = 1 And intDefensaInicio = 3 Then
    If intAtaqueFin = 0 And intDefensaFin = 3 Then
        ProbabilidadDeBatalla = 1071 / 1296
        Exit Function
    Else
        ProbabilidadDeBatalla = 225 / 1296
        Exit Function
    End If
End If


'A2-D2
If intAtaqueInicio = 2 And intDefensaInicio = 2 Then
    If intAtaqueFin = 0 And intDefensaFin = 2 Then
        ProbabilidadDeBatalla = 581 / 1296
        Exit Function
    Else
        If intAtaqueFin = 1 And intDefensaFin = 1 Then
            ProbabilidadDeBatalla = 420 / 1296
            Exit Function
        Else
            ProbabilidadDeBatalla = 295 / 1296
            Exit Function
        End If
    End If
End If


'A3-D2
If intAtaqueInicio = 3 And intDefensaInicio = 2 Then
    If intAtaqueFin = 1 And intDefensaFin = 2 Then
        ProbabilidadDeBatalla = 1729 / 7776
        Exit Function
    Else
        If intAtaqueFin = 2 And intDefensaFin = 1 Then
            ProbabilidadDeBatalla = 3724 / 7776
            Exit Function
        Else 'intAtaqueFin = 3 And intDefensaFin = 0 Then
            ProbabilidadDeBatalla = 2323 / 7776
            Exit Function
        End If
    End If
End If


'A2-D3
If intAtaqueInicio = 2 And intDefensaInicio = 3 Then
    If intAtaqueFin = 0 And intDefensaFin = 3 Then
        ProbabilidadDeBatalla = 4816 / 7776
        Exit Function
    Else
        If intAtaqueFin = 1 And intDefensaFin = 2 Then
            ProbabilidadDeBatalla = 1981 / 7776
            Exit Function
        Else 'intAtaqueFin = 2 And intDefensaFin = 1 Then
            ProbabilidadDeBatalla = 979 / 7776
            Exit Function
        End If
    End If
End If


'A3-D3
If intAtaqueInicio = 3 And intDefensaInicio = 3 Then
    If intAtaqueFin = 0 And intDefensaFin = 3 Then
        ProbabilidadDeBatalla = 17871 / 46656
        Exit Function
    Else
        If intAtaqueFin = 1 And intDefensaFin = 2 Then
            ProbabilidadDeBatalla = 12348 / 46656
            Exit Function
        Else
            If intAtaqueFin = 2 And intDefensaFin = 1 Then
                ProbabilidadDeBatalla = 10017 / 46656
                Exit Function
            Else 'intAtaqueFin = 3 And intDefensaFin = 0 Then
                ProbabilidadDeBatalla = 6420 / 46656
                Exit Function
            End If
        End If
    End If
End If

End Function

Public Function CalcularGuerra(intAtaque As Integer, intDefensa As Integer) As Double
'Calcula la probabilidad de que en una guerra
'(seguidilla de Batallas hasta que uno de los dos ejercitos desaparece)
'se logre la Conquista (defensa=0)

Dim dblProbabilidad As Double
Dim intTropasEnJuego As Integer
'Dim dblCamino(3) As Double
Dim dblCamino1 As Double
Dim dblCamino2 As Double
Dim dblCamino3 As Double
Dim dblCamino4 As Double

'Cantidad de tropas de ataque y defensa no mayor a 3
Dim intTresAtaque As Integer
Dim intTresDefensa As Integer

'Casos especiales
If intAtaque <= 0 Or intDefensa <= 0 Then
    If intAtaque <= 0 Then
    'Si no hay ataque, seguro que no hay Conquista
        dblProbabilidad = 0
    Else
    'Si hay ataque y no hay defensa, seguro que hay conquista
        dblProbabilidad = 1
    End If
    
Else
    
    'Calcula la cantidad de tropas en juego
    intTropasEnJuego = HastaTres(IIf(intAtaque <= intDefensa, intAtaque, intDefensa))
    intTresAtaque = HastaTres(intAtaque)
    intTresDefensa = HastaTres(intDefensa)
    
    'Algoritmo: Se va armando el arbol, en el cual:
    '   - Los caminos se suman
    '   - Los saltos (de padre a hijo) se multiplican
    'Nota: todos los caminos se suman, aquellos que no conducen a
    'la conquista valdran cero.
    Select Case intTropasEnJuego
        Case 1
            'Hay 2 caminos
            
            '1er camino
            dblCamino1 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque, intTresDefensa - intTropasEnJuego)
            dblCamino1 = dblCamino1 * CalcularGuerra(intAtaque, intDefensa - intTropasEnJuego)
            
            '2do camino
            
            dblCamino2 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque - intTropasEnJuego, intTresDefensa)
            dblCamino2 = dblCamino2 * CalcularGuerra(intAtaque - intTropasEnJuego, intDefensa)
                        
            dblProbabilidad = dblCamino1 + dblCamino2
            
        Case 2
            'Hay 3 caminos
            
            '1er camino
            dblCamino1 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque, intTresDefensa - intTropasEnJuego)
            dblCamino1 = dblCamino1 * CalcularGuerra(intAtaque, intDefensa - intTropasEnJuego)
            
            '2do camino
            dblCamino2 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque - intTropasEnJuego, intTresDefensa)
            dblCamino2 = dblCamino2 * CalcularGuerra(intAtaque - intTropasEnJuego, intDefensa)
                        
            '3er camino
            dblCamino3 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque - 1, intTresDefensa - 1)
            dblCamino3 = dblCamino3 * CalcularGuerra(intAtaque - 1, intDefensa - 1)
                        
            dblProbabilidad = dblCamino1 + dblCamino2 + dblCamino3
            
        Case 3
            'Hay 4 caminos
            
            '1er camino
            dblCamino1 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque, intTresDefensa - intTropasEnJuego)
            dblCamino1 = dblCamino1 * CalcularGuerra(intAtaque, intDefensa - intTropasEnJuego)
            
            '2do camino
            dblCamino2 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque - intTropasEnJuego, intTresDefensa)
            dblCamino2 = dblCamino2 * CalcularGuerra(intAtaque - intTropasEnJuego, intDefensa)
                        
            '3er camino
            dblCamino3 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque - 1, intTresDefensa - 2)
            dblCamino3 = dblCamino3 * CalcularGuerra(intAtaque - 1, intDefensa - 2)
                        
            '4to camino
            dblCamino4 = ProbabilidadDeBatalla(intTresAtaque, intTresDefensa, intTresAtaque - 2, intTresDefensa - 1)
            dblCamino4 = dblCamino4 * CalcularGuerra(intAtaque - 2, intDefensa - 1)
                        
            dblProbabilidad = dblCamino1 + dblCamino2 + dblCamino3 + dblCamino4
    
    End Select

End If

CalcularGuerra = dblProbabilidad

End Function

Public Function HastaTres(intNumero As Integer) As Integer
' Recibe un entero cualquiera y si es mayor que tres,
' devuelve el numero tres, sino devuelve el mismo nro.

If intNumero > 3 Then
    HastaTres = 3
Else
    HastaTres = intNumero
End If

End Function

Public Function GuerraCalculable(intAtaque As Integer, intDefensa As Integer) As Boolean
    'Dada una guerra determinada devuelve si es o no posible
    'calcularla con el algoritmo
    Dim intMax As Integer
    Dim intMin As Integer
    
    'Busca el máximo y mínimo
    If intAtaque > intDefensa Then
        intMax = intAtaque
        intMin = intDefensa
    Else
        intMax = intDefensa
        intMin = intAtaque
    End If
    
    'Verifica que se cumplan las condiciones
    If intMax > 3 And intMin > 3 Then
        If intMax <= 40 And (intMin <= 10 Or (intMax + intMin) <= 30) Then
            GuerraCalculable = True
        Else
            GuerraCalculable = False
        End If
    Else
        GuerraCalculable = True
    End If
    
End Function

Public Function ProbabilidadDeGanarGuerra(intAtaque As Integer, intDefensa As Integer) As Double

    Dim sngAtaque As Single
    Dim sngDefensa As Single
    
    sngAtaque = intAtaque
    sngDefensa = intDefensa
    
    'Simplifica la guerra hasta llevarla a valores calculables
    While Not GuerraCalculable(CInt(sngAtaque), CInt(sngDefensa))
        sngAtaque = sngAtaque - 1.9
        sngDefensa = sngDefensa - 1.1
    Wend
    
    ProbabilidadDeGanarGuerra = CalcularGuerra(CInt(sngAtaque), CInt(sngDefensa))

End Function
