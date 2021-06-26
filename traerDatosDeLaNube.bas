Attribute VB_Name = "traerDatosDeLaNube"
Sub traerDatos()
Dim Url As String, lastRow As Long
Dim XMLHTTP As Object, html As Object
Dim tbl As Object, obj_tbl As Object, obj_row As Object
Dim TR As Object, TD As Object

'Cambiar este link
'En la spreadsheet Archivo--> Publicar en la web --Pagina web
Url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQvAwN4g8lWXDy3Kp-HStEjdB5835sBv71JUdJoa6qHVAYXyijKWwNEwBfKZu8WSgGKXPzJ8fDYDHYJ/pubhtml"

Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
XMLHTTP.Open "GET", Url, False

XMLHTTP.setRequestHeader "Content-Type", "text/xml"

XMLHTTP.send

Set html = CreateObject("htmlfile")
html.body.innerHTML = XMLHTTP.responseText
Set obj_tbl = html.getElementsByTagName("table")

SW = 0

For Each tbl In obj_tbl
    If tbl.className = "waffle" Then
        Set TR = tbl.getElementsByTagName("TR")
        For Each obj_row In TR
            For Each TD In obj_row.getElementsByTagName("TD")
                
                'Este es el rago que va a servir como ancla C12 de la pagina= item 1 de la sheet
                If obj_row.getElementsByTagName("TD").Item(1).innerText = CStr(Range("C12").Value) Then
                
                'Rango a donde va a traer los datos y el item que va a traer de la sheet
                Range("C14").Value = obj_row.getElementsByTagName("TD").Item(2).innerText
                Range("C16").Value = obj_row.getElementsByTagName("TD").Item(3).innerText
                Range("C18").Value = obj_row.getElementsByTagName("TD").Item(4).innerText
                
                SW = 1
                End If
            Next
        Next
    End If
Next
If SW = 0 Then
 Application.Speech.Speak ("Su búsqueda no arrojo resultado")
 MsgBox "Su búsqueda no arrojó resultado", vbInformation, "Nombre Tecnico"
End If

Set XMLHTTP = Nothing

End Sub

