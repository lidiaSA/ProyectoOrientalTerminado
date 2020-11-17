

Dim username: username = "userGaia"
Dim password: password = "Gaia2020"
Dim dataSource: dataSource = "186.3.247.227,8282"
Dim dataB: dataB = "bdChatBotOriental"

'var para reporte diario
Dim objExcelReporte
Dim excelBaseReporte
Dim workSheetR1
Dim workSheetR2
Dim workSheetR3
Dim workSheetR4
Dim workSheetR5
Dim workSheetR6
Dim workSheetR7
Dim workSheetR8
Dim workSheetR9
Dim nombreReporte

Const initExcelReporte = 2 'fila fija. en esta se debe comenzar a escrir en el excel
Const inicial = 0
Dim excRowLog 'fila en la que se debe de escribir en workSheetLog
excRowLog = initExcelReporte


Dim fechaReporte
fechaReporte = Day(NOW) & "-" & Month(NOW) & "-" & YEAR(NOW) & ".xlsx"
Dim fechaReporte1
fechaReporte1 = Day(NOW) & "-" & Month(NOW) & "-" & YEAR(NOW)
nombreReporte = "C:\GAIA\OrientalFormulario\reportes\Reporte-Diario-ChatBot-" & fechaReporte

Set objExcelReporte = CreateObject("Excel.Application")
objExcelReporte.Visible = True

Set excelBaseReporte = objExcelReporte.Workbooks.Open("C:\GAIA\OrientalFormulario\reportes\ORIENTALREPORTES.xlsx")  'abre excel en 2 plano

Set workSheetR1 = excelBaseReporte.WorkSheets("BASE GENERAL")
Set workSheetR2 = excelBaseReporte.WorkSheets("PRODUCTO MAS VENDIDO")
Set workSheetR3 = excelBaseReporte.WorkSheets("CATEGORIA MAS VENDIDA")
Set workSheetR4 = excelBaseReporte.WorkSheets("PEDIDOS POR CLIENTE")
Set workSheetR5 = excelBaseReporte.WorkSheets("INTENCIONES CLIENTES")
Set workSheetR6 = excelBaseReporte.WorkSheets("COMENTARIOS")
Set workSheetR7 = excelBaseReporte.WorkSheets("FUERA DE LIBRETO")
Set workSheetR8 = excelBaseReporte.WorkSheets("CALIFICACION")
Set workSheetR9 = excelBaseReporte.WorkSheets("CLIENTES QUE NO HICIERON PEDIDO")


excelBaseReporte.SaveAs (nombreReporte) 'guardo instancia del excel con la fecha actual

'Fecha
workSheetR1.Cells(4,5 ) = fechaReporte1
' Inicion seccion para pestaÃ±a Intecion Clientes reporte

Dim idPedido
Dim nombreCliente
Dim apellidosCliente
Dim cedulaCLiente
Dim telefonoCliente
Dim fechaCliente
Dim canalMensajeria
Dim nombreProducto
Dim categoriaProducto
Dim totalPedido
dim sumaTotal
dim totalProducto

 sumaTotal = 0.0
 totalProducto = 0.0
Set objConnection = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB
' exec [dbo].[sp_getclientes] @celular = '${celular}'
rst.Open "DECLARE       @return_value int EXEC @return_value = [dbo].[sp_getclientesByDay] @hoy='"&fechaReporte1&"'  SELECT 'Return Value' = @return_value", objConnection
Do While Not rst.EOF

    idPedido = rst.Fields(0)
    nombreCliente = rst.Fields(1)
    apellidosCliente = rst.Fields(2)
    cedulaCLiente = rst.Fields(3)
    telefonoCliente = rst.Fields(4)
    canalMensajeria = rst.Fields(5)
    fechaCliente = rst.Fields(6)
    nombreProducto= rst.Fields(7)
    totalPedido = rst.Fields(8)
    categoriaProducto = rst.Fields(9)
    
    
    workSheetR1.Cells(excRowLog+5, 1) = nombreCliente &" "&apellidosCliente
    workSheetR1.Cells(excRowLog+5, 2) = idPedido
    workSheetR1.Cells(excRowLog+5, 3) = cedulaCLiente
    
    workSheetR1.Cells(excRowLog+5, 4) = telefonoCliente
    workSheetR1.Cells(excRowLog+5, 5) = canalMensajeria
    workSheetR1.Cells(excRowLog+5, 6) = fechaCliente
    workSheetR1.Cells(excRowLog+5, 7) = nombreProducto
    workSheetR1.Cells(excRowLog+5, 8) = categoriaProducto 
    workSheetR1.Cells(excRowLog+5, 9) = totalPedido
    sumaTotal = sumaTotal + rst.Fields(9)

    rst.MoveNext
    excRowLog = excRowLog + 1
    totalProducto = totalProducto + 1  
Loop
workSheetR1.Cells(7,13) = sumaTotal
workSheetR1.Cells(10,13) = totalProducto

rst.Close
objConnection.Close

' Fin seccion para pestaÃ±a Intecion Clientes reporte

' Inicio seccion para pestaÃ±a INTENCION DE COMPRAR

excRowLog = initExcelReporte
excRowLog = excRowLog - 1


Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB

rst.Open "DECLARE       @return_value int EXEC @return_value = [dbo].[sp_getProductoMasVendido] @hoy='"&fechaReporte1&"' SELECT 'Return Value' = @return_value", objConnection

Do While Not rst.EOF

      workSheetR2.Cells(excRowLog+1, 1) = rst.Fields(0)
      workSheetR2.Cells(excRowLog+1, 2) = rst.Fields(1)
      workSheetR2.Cells(excRowLog+1, 3) = rst.Fields(2)
    

    rst.MoveNext
    excRowLog = excRowLog + 1
Loop
rst.Close
objConnection.Close

' Fin  seccion para pestaÃ±a INTENCION DE COMPRAR


' Inicio seccion para pestaÃ±a SEGUIMIENTO CARRITO

excRowLog = initExcelReporte
excRowLog = excRowLog - 1

Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB

rst.Open "DECLARE       @return_value int EXEC @return_value = [dbo].[sp_getCategoriasMasVendidas] @hoy='"&fechaReporte1&"' SELECT  'Return Value' = @return_value", objConnection
Do While Not rst.EOF

    workSheetR3.Cells(excRowLog+1, 1) = rst.Fields(0)
    workSheetR3.Cells(excRowLog+1, 2) = rst.Fields(1)
    workSheetR3.Cells(excRowLog+1, 3) = rst.Fields(2)
    rst.MoveNext
    excRowLog = excRowLog + 1
Loop
rst.Close
objConnection.Close

' Fin  seccion para pestaÃ±a INTENCION DE COMPRAR



' Inicio seccion para pestaÃ±a SEGUIMIENTO CARRITO

excRowLog = initExcelReporte
excRowLog = excRowLog - 1

Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB
dim cantPedidos
    cantPedidos=0

rst.Open "DECLARE @return_value int EXEC	@return_value = [dbo].[sp_getPedidosID] @hoy = '"&fechaReporte1&"' SELECT  'Return Value' = @return_value", objConnection
Do While Not rst.EOF
    
    set rst2 = CreateObject("ADODB.RecordSet")
    rst2.Open "DECLARE @return_value int EXEC	@return_value = [dbo].[sp_getClienteByCedula] @cedula='"&rst.Fields(1)&"' SELECT  'Return Value' = @return_value", objConnection
    Do While Not rst2.EOF
        workSheetR4.Cells(excRowLog+1, 1) = "Nombres"
        workSheetR4.Cells(excRowLog+1, 2) = rst2.Fields(0)+rst2.Fields(1)
        workSheetR4.Cells(excRowLog+2, 1) = "C"&Chr(233)&"dula"
        workSheetR4.Cells(excRowLog+2, 2) =rst.Fields(1)
        workSheetR4.Cells(excRowLog+3, 1) = "Tel"&Chr(233)&"fono"
        workSheetR4.Cells(excRowLog+3, 2) = rst2.Fields(2)
        workSheetR4.Cells(excRowLog+4, 1) = "Ubicaci"&Chr(243)&"n"
        workSheetR4.Cells(excRowLog+4, 2) = rst2.Fields(4)
         workSheetR4.Cells(excRowLog+5, 1) = "Direcci"&Chr(243)&"n"
        workSheetR4.Cells(excRowLog+5, 2) = rst2.Fields(3)
        workSheetR4.Cells(excRowLog+6, 1) = "Celular Messenger"
        workSheetR4.Cells(excRowLog+6, 2) = rst2.Fields(5)
       
        workSheetR4.Cells(excRowLog+7, 1) = "Id Pedido"
        workSheetR4.Cells(excRowLog+7, 2) = rst.Fields(0)
        rst2.MoveNext
        excRowLog = excRowLog + 1
    Loop
    rst2.Close
    Set rst1 = CreateObject("ADODB.RecordSet")
    
    'objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB
    rst1.Open "DECLARE @return_value int EXEC	@return_value = [dbo].[sp_getDetallePedidosByID] @idPed = "&rst.Fields(0)&"SELECT  'Return Value' = @return_value", objConnection
    Do While Not rst1.EOF
        workSheetR4.Cells(excRowLog+1, 5).NumberFormat = "0.##"

        workSheetR4.Cells(excRowLog+1, 3) = rst1.Fields(0)
        workSheetR4.Cells(excRowLog+1, 4) = rst1.Fields(1)
        workSheetR4.Cells(excRowLog+1, 5) = rst1.Fields(2)
        
        rst1.MoveNext
        excRowLog = excRowLog + 1
        
    Loop
    rst1.Close
     workSheetR4.Cells(excRowLog+1, 5).NumberFormat = "0.##"
    workSheetR4.Cells(excRowLog+1, 3).Interior.ColorIndex = 3
    workSheetR4.Cells(excRowLog+1, 3) = "TOTAL DE VENTA"
    workSheetR4.Cells(excRowLog+1, 5) = rst.Fields(2)
   
    rst.MoveNext
    excRowLog = excRowLog + 1
    cantPedidos = cantPedidos + 1
Loop
rst.Close
workSheetR1.Cells(12,13) = cantPedidos
objConnection.Close

'Para las intencioones

excRowLog = initExcelReporte
excRowLog = excRowLog - 1


Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB

rst.Open "DECLARE @return_value int EXEC @return_value = [dbo].[sp_getIntencionesByDay] @hoy ='"&fechaReporte1&"' SELECT  'Return Value' = @return_value", objConnection

Do While Not rst.EOF

      workSheetR5.Cells(excRowLog+9, 1) = rst.Fields(0)
      workSheetR5.Cells(excRowLog+9, 2) = rst.Fields(1)
     
    rst.MoveNext
    excRowLog = excRowLog + 1
Loop
rst.Close
objConnection.Close



'Para las intencioones de saludo

excRowLog = initExcelReporte
excRowLog = excRowLog - 1


Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB

rst.Open "DECLARE @return_value int EXEC @return_value = [dbo].[sp_IngresaronChatBot] @hoy ='"&fechaReporte1&"' SELECT  'Return Value' = @return_value", objConnection

Do While Not rst.EOF

      workSheetR5.Cells(excRowLog+1, 2) = rst.Fields(1)
     
    rst.MoveNext
    excRowLog = excRowLog + 1
Loop
rst.Close
objConnection.Close



'Para las calificaciones

excRowLog = initExcelReporte
excRowLog = excRowLog - 1
dim totalCalificaciones
totalCalificaciones = 0

Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB

rst.Open "DECLARE @return_value int EXEC @return_value = [dbo].[sp_getCalificacionesByDay] @hoy ='"&fechaReporte1&"' SELECT  'Return Value' = @return_value", objConnection

Do While Not rst.EOF
    workSheetR8.Cells(excRowLog+1, 1) = rst.Fields(0)
     
      workSheetR8.Cells(excRowLog+1, 2) = rst.Fields(1)
        totalCalificaciones = totalCalificaciones + rst.Fields(1) 
    rst.MoveNext
    excRowLog = excRowLog + 1
    
Loop

workSheetR8.Cells(6, 1) = "Total de Calificaciones"
     
workSheetR8.Cells(6, 2) = totalCalificaciones

rst.Close
objConnection.Close




'Para los comentarios

excRowLog = initExcelReporte
excRowLog = excRowLog - 1

dim totalComentarios
totalComentarios=0
Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB

rst.Open "DECLARE @return_value int EXEC @return_value = [dbo].[sp_getComentarioByDay] @hoy ='"&fechaReporte1&"' SELECT  'Return Value' = @return_value", objConnection

Do While Not rst.EOF
    workSheetR6.Cells(excRowLog+1, 1) = rst.Fields(0)
     
      workSheetR6.Cells(excRowLog+1, 2) = rst.Fields(1)
    totalComentarios= totalComentarios + rst.Fields(1)
    rst.MoveNext
    excRowLog = excRowLog + 1
    
Loop
workSheetR6.Cells(1, 3) = "Total de comentarios"
     
workSheetR6.Cells(2, 3) = totalComentarios
     
rst.Close
objConnection.Close



'Para los fuera de libreto

excRowLog = initExcelReporte
excRowLog = excRowLog - 1

dim totalFueraDeLibreto
totalFueraDeLibreto=0
Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB

rst.Open "DECLARE @return_value int EXEC @return_value = [dbo].[sp_getFueraDeLibretoByDay] @hoy ='"&fechaReporte1&"' SELECT  'Return Value' = @return_value", objConnection

Do While Not rst.EOF
    workSheetR7.Cells(excRowLog+1, 1) = rst.Fields(0)
    rst.MoveNext
    excRowLog = excRowLog + 1
   totalFueraDeLibreto= totalFueraDeLibreto + 1
Loop
workSheetR7.Cells(1, 2) = "Total de fuera de libretos"
     
workSheetR7.Cells(2, 2) = totalFueraDeLibreto
     
rst.Close
objConnection.Close




' Para clientes que no hicieron pedidos

excRowLog = initExcelReporte
excRowLog = excRowLog - 1

Set rst = CreateObject("ADODB.RecordSet")
objConnection.open "Provider=SQLOLEDB; Data Source="&dataSource&"; uid="&username&"; pwd="&password&"; DATABASE=" &dataB

rst.Open "DECLARE       @return_value int EXEC @return_value = [dbo].[sp_getclientesNotPedidosIds] @hoy='"&fechaReporte1&"' SELECT  'Return Value' = @return_value", objConnection
Do While Not rst.EOF
    Set rst2 = CreateObject("ADODB.RecordSet")
    rst2.Open "DECLARE       @return_value int EXEC @return_value = [dbo].[sp_getclienteById] @id="&rst.Fields(0)&" SELECT  'Return Value' = @return_value", objConnection
    Do While Not rst2.EOF
      
        workSheetR9.Cells(excRowLog+1, 1) = rst2.Fields(0)
        workSheetR9.Cells(excRowLog+1, 2) = rst2.Fields(1)
        workSheetR9.Cells(excRowLog+1, 3) = rst2.Fields(2)
        workSheetR9.Cells(excRowLog+1, 4) = rst2.Fields(3)
        IF (rst.Fields(1) = 1) Then
            workSheetR9.Cells(excRowLog+1, 6) = "WHATSAAP"
        END IF
        IF (rst.Fields(1) = 2) Then
            workSheetR9.Cells(excRowLog+1, 6) = "FACEBOOK"
        END IF
        rst2.MoveNext
        excRowLog = excRowLog + 1
    Loop
    rst2.Close
    rst.MoveNext
Loop
workSheetR9.Cells(3, 7) = excRowLog-1

rst.Close

objConnection.Close

' Fin  seccion 



excelBaseReporte.Close True
objExcelReporte.Quit

objExcelReporte = Empty
Set excelBaseReporte = Nothing
'televentas@gruporiental.com
Set emailObj = CreateObject("CDO.Message")

emailObj.From = "Chi-li-lee@gruporiental.com"
emailObj.To = "lidia.salcedo@gaiaconsultores.biz"
emailObj.Cc = "televentas@gruporiental.com"

emailObj.Subject = "Reporte Diario ChatBot Chi Li Lee"
emailObj.TextBody = "Envios de Reportes Diario ChatBot"
emailObj.TextBody = "Test envio masasail"
emailObj.HTMLBody = "<p>Saludos Cordiales</p><br><p>En la presente se adjunta el reporte diario del chatbot con la informacion de los pedidos, intenciones, clientes que se han registrado el dia de hoy.</p><br><p>Su asistente virtual Chi Li Lee</p>"
emailObj.AddAttachment (nombreReporte)

Set emailConfig = emailObj.Configuration

emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.comandato.com"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "chatbot1"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "cmdchatbot#20"

emailConfig.Fields.Update

emailObj.Send
Set emailConfig = Nothing





