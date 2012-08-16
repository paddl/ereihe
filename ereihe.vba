Public Function EReihe(Source As Double)

    Dim rval, last_res, closest, exponent, mantissa
    
    '1% Resistor Value array E24, E48 and E96
    ResValue = Array(1#, 1.02, 1.05, 1.07, 1.1, 1.13, 1.15, 1.18, 1.2, 1.21, 1.24, 1.27, 1.3, 1.33, 1.37, 1.4, 1.43, 1.47, _
                    1.5, 1.54, 1.58, 1.6, 1.62, 1.65, 1.69, 1.74, 1.78, 1.8, 1.82, 1.87, 1.91, 1.96, 2#, 2.05, 2.1, 2.15, _
                    2.2, 2.21, 2.26, 2.32, 2.37, 2.4, 2.43, 2.49, 2.55, 2.61, 2.67, 2.7, 2.74, 2.8, 2.87, 2.94, 3#, 3.01, _
                    3.09, 3.16, 3.24, 3.3, 3.32, 3.4, 3.48, 3.57, 3.6, 3.65, 3.74, 3.83, 3.9, 3.92, 4.02, 4.12, 4.22, 4.3, _
                    4.32, 4.42, 4.53, 4.64, 4.7, 4.75, 4.87, 4.99, 5.1, 5.11, 5.23, 5.36, 5.49, 5.6, 5.62, 5.76, 5.9, 6.04, _
                    6.19, 6.2, 6.34, 6.49, 6.65, 6.8, 6.81, 6.98, 7.15, 7.32, 7.5, 7.68, 7.87, 8.06, 8.2, 8.25, 8.45, 8.66, _
                    8.87, 9.09, 9.1, 9.31, 9.53, 9.76, 10#)
                    
    exponent = WorksheetFunction.RoundDown(WorksheetFunction.Log(Source, 10), 0)
    
    mantissa = Source / 10 ^ exponent
    
    For Each rval In ResValue
    
        closest = rval
        
        If rval < mantissa Then
            last_res = rval
        ElseIf rval = mantissa Then
            Exit For
        Else
            If (rval - mantissa) >= (mantissa - last_res) Then
                closest = last_res
            End If
            Exit For
        End If
        
    Next rval
    
    EReihe = closest * (10 ^ exponent)

End Function