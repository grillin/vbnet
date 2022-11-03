Imports System.IO

Public Class Gtexto

    Public ftexto As StreamWriter

    Sub iniciar_rpt(ByVal ruta As String)
        ftexto = New StreamWriter(ruta, False)
    End Sub

    Sub adicionar_linea(ByVal linea As String)
        ftexto.WriteLine(linea)
    End Sub

    Sub terminar_rpt()
        ftexto.Close()
    End Sub

    Sub ancho_pagina(ByVal ancho As Integer)

    End Sub

    Sub largo_pagina(ByVal largo As Integer)

    End Sub

    Sub cabecera_pagina()

    End Sub

    Sub pie_pagina()

    End Sub

    Sub detalle_pagina()

    End Sub

    Sub imprimir_linea(ByVal ParamArray valores() As String)


        Dim datos(10, 2) As String
        Dim conta As Integer = 1
        Dim item As String
        Dim indice As Integer = 1

        For Each item In valores

            If conta = 1 Then

                datos(indice, 1) = CInt(item)
                conta = conta + 1
            Else
                If conta = 2 Then
                    datos(indice, 2) = item
                    conta = 1
                    indice = indice + 1
                End If
            End If

        Next

        indice = indice - 1

        Dim i As Integer        
        Dim linea As String = ""

        For i = 1 To indice
            While Len(linea) < CInt(datos(i, 1))
                linea = linea + " "
            End While
            linea = linea + datos(i, 2)
        Next

        'MsgBox(linea)
        Me.adicionar_linea(linea)

    End Sub

End Class
