Attribute VB_Name = "Cortos"

'Hablar en Excel
Sub hablarExcel()
    Application.Speech.Speak ("Hola amigos, ¿Cómo estan?")
    MsgBox "Hola amigos , soy un mensaje de texto"
End Sub


'Preguntar para Deshabilitar un boton
Sub DeshabilitarBoton()

    answer = MsgBox("¿Deshabilitar el boton?", vbYesNo)

    If answer = vbYes Then
    CommandButton1.Enabled = False
    Else
    'no hacer nada
    End If

End Sub


'Proteger y desproteger hoja

Sub protegerYDesproteger()
    ActiveSheet.Protect ("contrasena")
    ActiveSheet.Unprotect ("contrasena")
End Sub


