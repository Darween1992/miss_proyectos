Attribute VB_Name = "Módulo1"
Sub IMPRIMIR()

Dim i As Byte
Dim CONTADOR As Byte

CONTADOR = 1


Dim respuesta As Byte
Dim titulo As String

titulo = "Mensaje de Aviso"
respuesta = MsgBox(" DESEA IMPRIMIR LOS FORMULARIOS DE TODOS LOS OPERARIOS?" & Chr(13) & vbNewLine & " ¿Desea Continuar?", vbQuestion + vbYesNo, titulo)

If respuesta = vbYes Then



For i = 1 To Sheets("FORMATO").Range("R1").Value

Sheets("FORMATO").Select

Range("P1") = CONTADOR



 ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False


CONTADOR = CONTADOR + 1


Next i


Range("P1") = 1

Else

MsgBox ("No se Imprimio Ningun Formulario")

End If

End Sub



