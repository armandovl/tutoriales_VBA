Attribute VB_Name = "Cortos"

'****************Hablar en Excel************************************************************
Sub hablarExcel()
    Application.Speech.Speak ("Hola amigos, ¿Cómo estan?")
    MsgBox "Hola amigos , soy un mensaje de texto"
End Sub

'***************Preguntar para Deshabilitar un boton***************************************
Sub DeshabilitarBoton()
    answer = MsgBox("¿Deshabilitar el boton?", vbYesNo)  
    If answer = vbYes Then
    CommandButton1.Enabled = False
    Else
    'no hacer nada
    End If    
End Sub

'**************Proteger y desproteger hoja*************************************************

Sub protegerYDesproteger()
    ActiveSheet.Protect ("contrasena")
    ActiveSheet.Unprotect ("contrasena")
End Sub


' ***********copiar y pegar de un lado a otro siempre que haya espacio en blanco***********
Sub copiaryPegarBlanco()
    Sheets("copiar").Select 'seleccionar hoja copiar
    Range("A1:C2").Select  'seleccionar rango a copiar
    Selection.Copy 'copiar la seleccion
    Sheets("pegar").Select 'seleccionar la hoja donde se va a pegar por nombre
    Range("D3").Select ' seleccionar el rango donde se va a empezar a pegarse

    Do While Not IsEmpty(ActiveCell) 'mientras la celda no este limpia
    ActiveCell.Offset(1, 0).Select 'recorrete una fila hacia abajo y cero columnas
    Loop
    ActiveSheet.Paste 'pega en la celda activa
End Sub

'****************************copiar y pegar de una hoja oculta a otra*****
Sub copiaryPegarBlanco()
    'Se puede copiar con este comando incluso en hojas ocultas
    Worksheets("copiar").Range("A1:C2").Copy

    Sheets("pegar").Select 'seleccionar la hoja donde se va a pegar por nombre
    Range("D3").Select ' seleccionar el rango donde se va a empezar a pegarse

    Do While Not IsEmpty(ActiveCell) 'mientras la celda no este limpia
    ActiveCell.Offset(1, 0).Select 'recorrete una fila hacia abajo y cero columnas
    Loop

    ActiveSheet.Paste 'pega en la celda activa
End Sub

'****************************Habilitar e inhabilitar un boton*****                                
Private Sub CommandButton1_Click()
    MsgBox "Ejecutando botón"
    Command.Button1.Enable = False
End Sub

Private Sub CommandButton2_Click()
    Command.Button1.Enable = True
End Sub
                    



