Imports System.IO.Ports
Imports System.Timers

Module ModTransbankPOS

    Public Structure varPolling
        Public Estado As Boolean
        Public Respuesta As String
    End Structure

    Public Structure varNuevaVenta
        Public Comando As String
        Public Respuesta As String
        Public CodComercio As Long
        Public IdTerminal As String
        Public Ticket As String
        Public CodAutorizacion As String
        Public Monto As Integer
        Public Cuotas As Integer
        Public MontoCuota As Integer
        Public Ultimos4Dig As Integer
        Public NumOperacion As String
        Public TipoTarjeta As String
        Public FechaContable As String
        Public NumCuenta As String
        Public AbreviacionTarjeta As String
        Public FechaTransaccion As String
        Public HoraTransaccion As String
        Public Empleado As Integer
        Public Propina As Integer
    End Structure

    Public Structure varUltimaVenta
        Public Comando As String
        Public Respuesta As String
        Public CodComercio As Long
        Public IdTerminal As String
        Public Ticket As String
        Public CodAutorizacion As String
        Public Monto As Integer
        Public Cuotas As Integer
        Public MontoCuota As Integer
        Public Ultimos4Dig As Integer
        Public NumOperacion As String
        Public TipoTarjeta As String
        Public FechaContable As String
        Public NumCuenta As String
        Public AbreviacionTarjeta As String
        Public FechaTransaccion As String
        Public HoraTransaccion As String
        Public Empleado As Integer
        Public Propina As Integer
    End Structure

    Public Structure varAnulacion
        Public Comando As String
        Public Respuesta As String
        Public CodComercio As Long
        Public IdTerminal As String
        Public CodAutorizacion As String
        Public NumOperacion As String
    End Structure

    Public Structure varCierre
        Public Comando As String
        Public Respuesta As String
        Public CodComercio As Long
        Public IdTerminal As String
    End Structure

    Public Structure varDetalleDeVenta
        Public Estado As Boolean
        Public Respuesta As String
    End Structure

    Public Structure varSolicitudDeTotales
        Public Comando As String
        Public Respuesta As String
        Public NumeroTX As Integer
        Public Totales As Integer
    End Structure

    Public Structure varSolicitudCargaDeLLaves
        Public Comando As String
        Public Respuesta As String
        Public CodComercio As Long
        Public IdTerminal As String
    End Structure

    Private Tiempo As Integer
    Private WithEvents Temporizador As New Timer
    Private WithEvents PuertoCOM As New SerialPort

    Private FMuestreo As Integer = 115200
    Private STX As Byte = &H2
    Private ETX As Byte = &H3
    Private SEP As Byte = &H7C
    Private LRC As Byte
    Private Memoria As String
    Private Mensaje As String
    Public tbkPuerto As String = ""

    Private Sub Temporizador_Tick(sender As Object, e As EventArgs) Handles Temporizador.Elapsed
        Tiempo = Tiempo + 1
    End Sub

    Private Sub PuertoCOM_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles PuertoCOM.DataReceived
        Memoria = Memoria & PuertoCOM.ReadExisting
    End Sub

    Public Function ListaDePuertos() As List(Of String)
        Dim Lista As New List(Of String)

        If My.Computer.Ports.SerialPortNames.Count > 0 Then
            For Each Puerto As String In My.Computer.Ports.SerialPortNames
                Lista.Add(Puerto)
            Next
        Else
            Lista.Add("")
            MessageBox.Show("No Se Encontraron Puestos COM Activos", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        Return Lista
    End Function

    Public Function ConectarDesconectarPos(COM As String) As Boolean
        Dim Estado As Boolean

        If PuertoCOM.IsOpen Then
            PuertoCOM.Close()
            Estado = False
        Else
            PuertoCOM.BaudRate = FMuestreo
            PuertoCOM.PortName = COM
            PuertoCOM.Open()
            Estado = True
        End If

        Return Estado
    End Function

    Private Function ValidarConeccionCOM() As Boolean
        Dim Estado As Boolean

        If Not PuertoCOM Is Nothing Then
            If PuertoCOM.IsOpen Then
                Estado = True
            Else
                Estado = False
                tbkPuerto = ""
            End If
        Else
            Estado = False
            tbkPuerto = ""
        End If

        Return Estado
    End Function

    Public Function tbkPolling() As varPolling
        'STX + Comando + SEP + ETX + LRC
        Dim Estado As Boolean

        Mensaje = "0100|"
        Dim n As Integer = Mensaje.Length + 2
        Dim by(n) As Byte
        Memoria = ""
        Estado = True

        by(0) = STX  'Inicia Mensaje

        For i = 1 To n - 2
            by(i) = GetHexa(Mensaje.Substring(i - 1, 1)) 'Mensaje
        Next

        by(n - 1) = ETX 'Fin Mensaje
        by(n) = GetLRC(by) 'XOR LRC

        Try
            PuertoCOM.Write(by, 0, by.Length)

            Tiempo = 0
            Temporizador.Interval = 1000
            Temporizador.Start()

            While Estado
                If Memoria.IndexOf("") > -1 Then
                    Estado = False
                    Mensaje = Memoria
                ElseIf Tiempo > 5 Then
                    Estado = False
                    Mensaje = ""
                End If
            End While

            Temporizador.Stop()

            If Mensaje.Length > 0 Then
                tbkPolling.Estado = True
                tbkPolling.Respuesta = "POS Conectado Correctamente"
            Else
                tbkPolling.Estado = False
                tbkPolling.Respuesta = "POS Desconectado"
            End If
        Catch ex As Exception
            tbkPuerto = ""
            tbkPolling.Estado = False
            tbkPolling.Respuesta = "POS Desconectado"
        End Try

        Return tbkPolling
    End Function

    Public Function tbkNuevaVenta(Monto As Integer, Tiket As String) As varNuevaVenta
        'STX + Comando + SEP + Monto + SEP + Tiket + SEP + ETX + LRC
        Dim Estado As Boolean

        If ValidarConeccionCOM() Then
            Mensaje = "0200|" & Monto & "|" & Tiket & "|||0"
            Dim n As Integer = Mensaje.Length + 2
            Dim by(n) As Byte
            Memoria = ""
            Estado = True

            by(0) = STX  'Inicia Mensaje

            For i = 1 To n - 2
                by(i) = GetHexa(Mensaje.Substring(i - 1, 1)) 'Mensaje
            Next

            by(n - 1) = ETX 'Fin Mensaje
            by(n) = GetLRC(by) 'XOR LRC

            PuertoCOM.Write(by, 0, by.Length)
            Mensaje = ""

            Tiempo = 0
            Temporizador.Interval = 1000
            Temporizador.Start()

            While Estado
                If ValidarConeccionCOM() Then
                    If Memoria.IndexOf("") > 0 Then
                        Estado = False
                        Mensaje = Memoria.Substring(2, Memoria.Length - 4)
                    ElseIf Tiempo > 150 And Mensaje.Trim = "" Then ' Si No recibe datos en 150 Segundos asumira que el POS Esta Desconectado
                        Estado = False
                    End If
                Else
                    Mensaje = ""
                    Estado = False
                End If
            End While

            Temporizador.Stop()

            'Registro de Mensajes ----------------
            RegistrotbkNuevaVenta(Tiket, Mensaje)
            '-------------------------------------

            If Mensaje.Trim <> "" Then
                PuertoCOM.Write("")

                If Memoria.Substring(Memoria.Length - 1, 1) <> "|" Then
                    Memoria = Memoria & "|"
                End If

                Mensaje = Mensaje.Replace("||", "|0|")
                Mensaje = Mensaje.Replace("||", "|0|")

                tbkNuevaVenta.Comando = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.Respuesta = GetMensaje(Mensaje.Substring(0, Mensaje.IndexOf("|")))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.CodComercio = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.IdTerminal = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.Ticket = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.CodAutorizacion = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.Monto = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.Cuotas = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.MontoCuota = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.Ultimos4Dig = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.NumOperacion = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.TipoTarjeta = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.FechaContable = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.NumCuenta = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.AbreviacionTarjeta = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.FechaTransaccion = FechaSQL(Mensaje.Substring(0, Mensaje.IndexOf("|")))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.HoraTransaccion = HoraSQL(Mensaje.Substring(0, Mensaje.IndexOf("|")))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.Empleado = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkNuevaVenta.Propina = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Else
                tbkNuevaVenta.Comando = ""
                tbkNuevaVenta.Respuesta = "Error con la Conexión del POS o Puerto COM No Encontrado"
                tbkNuevaVenta.CodComercio = 0
                tbkNuevaVenta.IdTerminal = ""
                tbkNuevaVenta.Ticket = ""
                tbkNuevaVenta.CodAutorizacion = ""
                tbkNuevaVenta.Monto = 0
                tbkNuevaVenta.Cuotas = 0
                tbkNuevaVenta.MontoCuota = 0
                tbkNuevaVenta.Ultimos4Dig = 0
                tbkNuevaVenta.NumOperacion = ""
                tbkNuevaVenta.TipoTarjeta = ""
                tbkNuevaVenta.FechaContable = ""
                tbkNuevaVenta.NumCuenta = 0
                tbkNuevaVenta.AbreviacionTarjeta = ""
                tbkNuevaVenta.FechaTransaccion = ""
                tbkNuevaVenta.HoraTransaccion = ""
                tbkNuevaVenta.Empleado = 0
                tbkNuevaVenta.Propina = 0
            End If
        Else
            tbkNuevaVenta.Comando = ""
            tbkNuevaVenta.Respuesta = "Error con la Conexión del POS o Puerto COM No Encontrado"
            tbkNuevaVenta.CodComercio = 0
            tbkNuevaVenta.IdTerminal = ""
            tbkNuevaVenta.Ticket = ""
            tbkNuevaVenta.CodAutorizacion = ""
            tbkNuevaVenta.Monto = 0
            tbkNuevaVenta.Cuotas = 0
            tbkNuevaVenta.MontoCuota = 0
            tbkNuevaVenta.Ultimos4Dig = 0
            tbkNuevaVenta.NumOperacion = ""
            tbkNuevaVenta.TipoTarjeta = ""
            tbkNuevaVenta.FechaContable = ""
            tbkNuevaVenta.NumCuenta = 0
            tbkNuevaVenta.AbreviacionTarjeta = ""
            tbkNuevaVenta.FechaTransaccion = ""
            tbkNuevaVenta.HoraTransaccion = ""
            tbkNuevaVenta.Empleado = 0
            tbkNuevaVenta.Propina = 0
        End If

        Return tbkNuevaVenta
    End Function

    Public Function tbkUltimaVenta() As varUltimaVenta
        'STX + Comando + SEP + ETX + LRC
        Dim Estado As Boolean

        If ValidarConeccionCOM() Then
            Mensaje = "0250|"
            Dim n As Integer = Mensaje.Length + 2
            Dim by(n) As Byte
            Memoria = ""
            Estado = True

            by(0) = STX  'Inicia Mensaje

            For i = 1 To n - 2
                by(i) = GetHexa(Mensaje.Substring(i - 1, 1)) 'Mensaje
            Next

            by(n - 1) = ETX 'Fin Mensaje
            by(n) = GetLRC(by) 'XOR LRC

            PuertoCOM.Write(by, 0, by.Length)

            While Estado
                If Memoria.IndexOf("") > 0 Then
                    Estado = False
                    Mensaje = Memoria.Substring(2, Memoria.Length - 4)
                End If
            End While

            PuertoCOM.Write("")

            Mensaje = Mensaje.Replace("||", "|0|")
            Mensaje = Mensaje.Replace("||", "|0|")

            tbkUltimaVenta.Comando = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.Respuesta = GetMensaje(Mensaje.Substring(0, Mensaje.IndexOf("|")))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.CodComercio = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.IdTerminal = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.Ticket = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.CodAutorizacion = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.Monto = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.Cuotas = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.MontoCuota = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.Ultimos4Dig = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.NumOperacion = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.TipoTarjeta = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.FechaContable = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.NumCuenta = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.AbreviacionTarjeta = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.FechaTransaccion = FechaSQL(Mensaje.Substring(0, Mensaje.IndexOf("|")))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.HoraTransaccion = HoraSQL(Mensaje.Substring(0, Mensaje.IndexOf("|")))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.Empleado = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkUltimaVenta.Propina = Mensaje.Substring(0, Mensaje.IndexOf("|"))
        Else
            tbkUltimaVenta.Comando = ""
            tbkUltimaVenta.Respuesta = "Error con la Conexión del POS o Puerto COM No Encontrado"
            tbkUltimaVenta.CodComercio = 0
            tbkUltimaVenta.IdTerminal = ""
            tbkUltimaVenta.Ticket = ""
            tbkUltimaVenta.CodAutorizacion = ""
            tbkUltimaVenta.Monto = 0
            tbkUltimaVenta.Cuotas = 0
            tbkUltimaVenta.MontoCuota = 0
            tbkUltimaVenta.Ultimos4Dig = 0
            tbkUltimaVenta.NumOperacion = ""
            tbkUltimaVenta.TipoTarjeta = ""
            tbkUltimaVenta.FechaContable = ""
            tbkUltimaVenta.NumCuenta = 0
            tbkUltimaVenta.AbreviacionTarjeta = ""
            tbkUltimaVenta.FechaTransaccion = ""
            tbkUltimaVenta.HoraTransaccion = ""
            tbkUltimaVenta.Empleado = 0
            tbkUltimaVenta.Propina = 0
        End If

        Return tbkUltimaVenta
    End Function

    Public Function tbkAnulacion(Tiket As String) As varAnulacion
        'STX + Comando + SEP + Tiket + ETX + LRC
        Dim Estado As Boolean

        If ValidarConeccionCOM() Then
            Mensaje = "1200|" & Tiket
            Dim n As Integer = Mensaje.Length + 2
            Dim by(n) As Byte
            Memoria = ""
            Estado = True

            by(0) = STX  'Inicia Mensaje

            For i = 1 To n - 2
                by(i) = GetHexa(Mensaje.Substring(i - 1, 1)) 'Mensaje
            Next

            by(n - 1) = ETX 'Fin Mensaje
            by(n) = GetLRC(by) 'XOR LRC

            PuertoCOM.Write(by, 0, by.Length)

            While Estado
                If Memoria.IndexOf("") > 0 Then
                    Estado = False
                    Mensaje = Memoria.Substring(2, Memoria.Length - 4)
                End If
            End While

            PuertoCOM.Write("")

            Mensaje = Mensaje.Replace("||", "|0|")

            tbkAnulacion.Comando = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)

            If Mensaje.Substring(0, Mensaje.IndexOf("|")) = "00" Then
                tbkAnulacion.Respuesta = GetMensaje(Mensaje.Substring(0, Mensaje.IndexOf("|")))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkAnulacion.CodComercio = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkAnulacion.IdTerminal = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkAnulacion.CodAutorizacion = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                tbkAnulacion.NumOperacion = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            Else
                tbkAnulacion.Respuesta = GetMensaje(Mensaje.Substring(0, Mensaje.IndexOf("|")))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkAnulacion.CodComercio = Mensaje.Substring(0, Mensaje.IndexOf("|"))
                Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
                tbkAnulacion.IdTerminal = Mensaje.Substring(0, Mensaje.IndexOf("|"))

                tbkAnulacion.CodAutorizacion = 0
                tbkAnulacion.NumOperacion = 0
            End If
        Else
            tbkAnulacion.Comando = ""
            tbkAnulacion.Respuesta = "Error con la Conexión del POS o Puerto COM No Encontrado"
            tbkAnulacion.CodComercio = 0
            tbkAnulacion.IdTerminal = ""
            tbkAnulacion.CodAutorizacion = ""
            tbkAnulacion.NumOperacion = ""
        End If

        Return tbkAnulacion
    End Function

    Public Function tbkCierre() As varCierre
        'STX + Comando + SEP + SEP + ETX + LRC
        Dim Estado As Boolean

        If ValidarConeccionCOM() Then
            Mensaje = "0500||"
            Dim n As Integer = Mensaje.Length + 2
            Dim by(n) As Byte
            Memoria = ""
            Estado = True

            by(0) = STX  'Inicia Mensaje

            For i = 1 To n - 2
                by(i) = GetHexa(Mensaje.Substring(i - 1, 1)) 'Mensaje
            Next

            by(n - 1) = ETX 'Fin Mensaje
            by(n) = GetLRC(by) 'XOR LRC

            PuertoCOM.Write(by, 0, by.Length)

            While Estado
                If Memoria.IndexOf("") > 0 Then
                    Estado = False
                    Mensaje = Memoria.Substring(2, Memoria.Length - 4)
                End If
            End While

            PuertoCOM.Write("")

            tbkCierre.Comando = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkCierre.Respuesta = GetMensaje(Mensaje.Substring(0, Mensaje.IndexOf("|")))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkCierre.CodComercio = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            tbkCierre.IdTerminal = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
        Else
            tbkCierre.Comando = ""
            tbkCierre.Respuesta = "Error con la Conexión del POS o Puerto COM No Encontrado"
            tbkCierre.CodComercio = 0
            tbkCierre.IdTerminal = ""
        End If

        Return tbkCierre
    End Function

    Public Function tbkDetalleDeVenta(ImprimirEnPOS As Boolean) As varDetalleDeVenta
        'STX + Comando + SEP + SEP + ETX + LRC
        'Opcion = 0: El POS Imprime el Voucher con el Detalle
        'Opcion = 1: El POS Envia a la Caja con el Detalle
        Dim Estado As Boolean

        If ValidarConeccionCOM() Then
            If ImprimirEnPOS Then
                Mensaje = "0260|0|" 'Imprime en POS
                Dim n As Integer = Mensaje.Length + 2
                Dim by(n) As Byte
                Memoria = ""
                Estado = True

                by(0) = STX  'Inicia Mensaje

                For i = 1 To n - 2
                    by(i) = GetHexa(Mensaje.Substring(i - 1, 1)) 'Mensaje
                Next

                by(n - 1) = ETX 'Fin Mensaje
                by(n) = GetLRC(by) 'XOR LRC

                PuertoCOM.Write(by, 0, by.Length)

                While Estado
                    If Memoria = "" Then
                        Estado = False
                    End If
                End While

                tbkDetalleDeVenta.Estado = True
                tbkDetalleDeVenta.Respuesta = "Aprobado"
            Else
                'Esto no esta Desarrollado
                Mensaje = "0260|1|" 'Enviar Datos a la Caja
                Dim n As Integer = Mensaje.Length + 2
                Dim by(n) As Byte
                Memoria = ""

                tbkDetalleDeVenta.Estado = False
                tbkDetalleDeVenta.Respuesta = "Aun No Programado"
            End If
        Else
            tbkDetalleDeVenta.Estado = False
            tbkDetalleDeVenta.Respuesta = "Error con la Conexión del POS o Puerto COM No Encontrado"
        End If

        Return tbkDetalleDeVenta
    End Function

    Public Function tbkSolicitudTotales() As varSolicitudDeTotales
        'STX + Comando + SEP + SEP + ETX + LRC
        Dim Estado As Boolean

        If ValidarConeccionCOM() Then
            Mensaje = "0700||"
            Dim n As Integer = Mensaje.Length + 2
            Dim by(n) As Byte
            Memoria = ""
            Estado = True

            by(0) = STX  'Inicia Mensaje

            For i = 1 To n - 2
                by(i) = GetHexa(Mensaje.Substring(i - 1, 1)) 'Mensaje
            Next

            by(n - 1) = ETX 'Fin Mensaje
            by(n) = GetLRC(by) 'XOR LRC

            PuertoCOM.Write(by, 0, by.Length)

            While Estado
                If Memoria.IndexOf("") > 0 Then
                    Estado = False
                    Mensaje = Memoria.Substring(2, Memoria.Length - 4)
                End If
            End While

            PuertoCOM.Write("")

            tbkSolicitudTotales.Comando = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkSolicitudTotales.Respuesta = GetMensaje(Mensaje.Substring(0, Mensaje.IndexOf("|")))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkSolicitudTotales.NumeroTX = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkSolicitudTotales.Totales = Mensaje.Substring(0, Mensaje.IndexOf("|"))
        Else
            tbkSolicitudTotales.Comando = ""
            tbkSolicitudTotales.Respuesta = "Error con la Conexión del POS o Puerto COM No Encontrado"
            tbkSolicitudTotales.NumeroTX = 0
            tbkSolicitudTotales.Totales = 0
        End If

        Return tbkSolicitudTotales
    End Function

    Public Function tbkSolicitudCargaDeLlaves() As varSolicitudCargaDeLLaves
        'STX + Comando + ETX + LRC
        Dim Estado As Boolean

        If ValidarConeccionCOM() Then
            Mensaje = "0800"
            Dim n As Integer = Mensaje.Length + 2
            Dim by(n) As Byte
            Memoria = ""
            Estado = True

            by(0) = STX  'Inicia Mensaje

            For i = 1 To n - 2
                by(i) = GetHexa(Mensaje.Substring(i - 1, 1)) 'Mensaje
            Next

            by(n - 1) = ETX 'Fin Mensaje
            by(n) = GetLRC(by) 'XOR LRC

            PuertoCOM.Write(by, 0, by.Length)

            While Estado
                If Memoria.IndexOf("") > 0 Then
                    Estado = False
                    Mensaje = Memoria.Substring(2, Memoria.Length - 4)
                End If
            End While

            PuertoCOM.Write("")

            tbkSolicitudCargaDeLlaves.Comando = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkSolicitudCargaDeLlaves.Respuesta = GetMensaje(Mensaje.Substring(0, Mensaje.IndexOf("|")))
            Mensaje = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
            tbkSolicitudCargaDeLlaves.CodComercio = Mensaje.Substring(0, Mensaje.IndexOf("|"))
            tbkSolicitudCargaDeLlaves.IdTerminal = Mensaje.Substring(Mensaje.IndexOf("|") + 1, Mensaje.Length - Mensaje.IndexOf("|") - 1)
        Else
            tbkSolicitudCargaDeLlaves.Comando = ""
            tbkSolicitudCargaDeLlaves.Respuesta = "Error con la Conexión del POS o Puerto COM No Encontrado"
            tbkSolicitudCargaDeLlaves.CodComercio = 0
            tbkSolicitudCargaDeLlaves.IdTerminal = 0
        End If

        Return tbkSolicitudCargaDeLlaves
    End Function

    Private Function GetHexa(Caracter As Char) As Byte
        Dim Resultado As Byte

        If Caracter = "0" Then
            Resultado = &H30
        ElseIf Caracter = "1" Then
            Resultado = &H31
        ElseIf Caracter = "2" Then
            Resultado = &H32
        ElseIf Caracter = "3" Then
            Resultado = &H33
        ElseIf Caracter = "4" Then
            Resultado = &H34
        ElseIf Caracter = "5" Then
            Resultado = &H35
        ElseIf Caracter = "6" Then
            Resultado = &H36
        ElseIf Caracter = "7" Then
            Resultado = &H37
        ElseIf Caracter = "8" Then
            Resultado = &H38
        ElseIf Caracter = "9" Then
            Resultado = &H39
        ElseIf Caracter = "A" Then
            Resultado = &H41
        ElseIf Caracter = "B" Then
            Resultado = &H42
        ElseIf Caracter = "C" Then
            Resultado = &H43
        ElseIf Caracter = "D" Then
            Resultado = &H44
        ElseIf Caracter = "E" Then
            Resultado = &H45
        ElseIf Caracter = "F" Then
            Resultado = &H46
        ElseIf Caracter = "G" Then
            Resultado = &H47
        ElseIf Caracter = "H" Then
            Resultado = &H48
        ElseIf Caracter = "I" Then
            Resultado = &H49
        ElseIf Caracter = "J" Then
            Resultado = &H4A
        ElseIf Caracter = "K" Then
            Resultado = &H4B
        ElseIf Caracter = "L" Then
            Resultado = &H4C
        ElseIf Caracter = "M" Then
            Resultado = &H4D
        ElseIf Caracter = "N" Then
            Resultado = &H4E
        ElseIf Caracter = "O" Then
            Resultado = &H4F
        ElseIf Caracter = "P" Then
            Resultado = &H50
        ElseIf Caracter = "Q" Then
            Resultado = &H51
        ElseIf Caracter = "R" Then
            Resultado = &H52
        ElseIf Caracter = "S" Then
            Resultado = &H53
        ElseIf Caracter = "T" Then
            Resultado = &H54
        ElseIf Caracter = "U" Then
            Resultado = &H55
        ElseIf Caracter = "|" Then
            Resultado = &H7C
        End If

        Return Resultado
    End Function

    Private Function GetLRC(bt As Byte()) As Byte
        Dim LRC As Byte
        Dim n As Integer = bt.Length

        LRC = bt(1) Xor bt(2)
        For i = 3 To n - 2
            LRC = LRC Xor bt(i)
        Next

        Return LRC
    End Function

    Private Function GetMensaje(CodMsge As String) As String
        Dim Resultado As String

        If CodMsge = "00" Then
            Resultado = "Aprobado"
        ElseIf CodMsge = "01" Then
            Resultado = "Rechazado"
        ElseIf CodMsge = "02" Then
            Resultado = "Host no Responde"
        ElseIf CodMsge = "03" Then
            Resultado = "Conexión Fallo"
        ElseIf CodMsge = "04" Then
            Resultado = "Transacción ya Fue Anulada"
        ElseIf CodMsge = "05" Then
            Resultado = "No existe Transacción para Anular"
        ElseIf CodMsge = "06" Then
            Resultado = "Tarjeta no Soportada"
        ElseIf CodMsge = "07" Then
            Resultado = "Transacción Cancelada desde el POS"
        ElseIf CodMsge = "08" Then
            Resultado = "No puede Anular Transacción Debito"
        ElseIf CodMsge = "09" Then
            Resultado = "Error Lectura Tarjeta"
        ElseIf CodMsge = "10" Then
            Resultado = "Monto menor al mínimo permitido"
        ElseIf CodMsge = "11" Then
            Resultado = "No existe venta"
        ElseIf CodMsge = "12" Then
            Resultado = "Transacción No Soportada"
        ElseIf CodMsge = "13" Then
            Resultado = "Debe ejecutar cierre"
        ElseIf CodMsge = "14" Then
            Resultado = "No hay Tono"
        ElseIf CodMsge = "15" Then
            Resultado = "Archivo BITMAP.DAT no encontrado. Favor cargue"
        ElseIf CodMsge = "16" Then
            Resultado = "Error Formato Respuesta del HOST"
        ElseIf CodMsge = "17" Then
            Resultado = "Error en los 4 últimos dígitos"
        ElseIf CodMsge = "18" Then
            Resultado = "Menú invalido"
        ElseIf CodMsge = "19" Then
            Resultado = "ERROR_TARJ_DIST"
        ElseIf CodMsge = "20" Then
            Resultado = "Tarjeta Invalida"
        ElseIf CodMsge = "21" Then
            Resultado = "Anulación. No Permitida"
        ElseIf CodMsge = "22" Then
            Resultado = "TIMEOUT"
        ElseIf CodMsge = "24" Then
            Resultado = "Impresora Sin Papel"
        ElseIf CodMsge = "25" Then
            Resultado = "Fecha Invalida"
        ElseIf CodMsge = "26" Then
            Resultado = "Debe Cargar Llaves"
        ElseIf CodMsge = "27" Then
            Resultado = "Debe Actualizar"
        ElseIf CodMsge = "60" Then
            Resultado = "Error en Número de Cuotas"
        ElseIf CodMsge = "61" Then
            Resultado = "Error en Armado de Solicitud"
        ElseIf CodMsge = "62" Then
            Resultado = "Problema con el Pinpad interno"
        ElseIf CodMsge = "65" Then
            Resultado = "Error al Procesar la Respuesta del Host"
        ElseIf CodMsge = "67" Then
            Resultado = "Superó Número Máximo de Ventas, Debe Ejecutar Cierre"
        ElseIf CodMsge = "68" Then
            Resultado = "Error Genérico, Falla al Ingresar Montos"
        ElseIf CodMsge = "70" Then
            Resultado = "Error de formato Campo de Boleta MAX 6"
        ElseIf CodMsge = "71" Then
            Resultado = "Error de Largo Campo de Impresión"
        ElseIf CodMsge = "72" Then
            Resultado = "Error de Monto Venta, Debe ser Mayor que '0'"
        ElseIf CodMsge = "73" Then
            Resultado = "Terminal ID no configurado"
        ElseIf CodMsge = "74" Then
            Resultado = "Debe Ejecutar CIERRE"
        ElseIf CodMsge = "75" Then
            Resultado = "Comercio no tiene Tarjetas Configuradas"
        ElseIf CodMsge = "76" Then
            Resultado = "Supero Número Máximo de Ventas, Debe Ejecutar CIERRE"
        ElseIf CodMsge = "77" Then
            Resultado = "Debe Ejecutar Cierre"
        ElseIf CodMsge = "78" Then
            Resultado = "Esperando Leer Tarjeta"
        ElseIf CodMsge = "79" Then
            Resultado = "Solicitando Confirmar Monto"
        ElseIf CodMsge = "81" Then
            Resultado = "Solicitando Ingreso de Clave"
        ElseIf CodMsge = "82" Then
            Resultado = "Enviando transacción al Host"
        ElseIf CodMsge = "88" Then
            Resultado = "Error Cantidad Cuotas"
        ElseIf CodMsge = "93" Then
            Resultado = "Declinada"
        ElseIf CodMsge = "94" Then
            Resultado = "Error al Procesar Respuesta"
        ElseIf CodMsge = "95" Then
            Resultado = "Error al Imprimir TASA"
        Else
            Resultado = "Error Desconocido en la Tarjeta y/o POS"
        End If

        Return Resultado
    End Function

    Private Function FechaSQL(ByVal Fecha As String) As String
        Dim FSQL As String

        If Fecha.Length < 8 Then
            FSQL = Nothing
        Else
            FSQL = Fecha.Substring(4, 4)
            FSQL = FSQL & Fecha.Substring(2, 2)
            FSQL = FSQL & Fecha.Substring(0, 2)
        End If

        Return FSQL
    End Function

    Private Function HoraSQL(ByVal Fecha As String) As String
        Dim HSQL As String

        If Fecha.Length < 6 Then
            HSQL = Nothing
        Else
            HSQL = Fecha.Substring(0, 2)
            HSQL = HSQL & ":" & Fecha.Substring(2, 2)
            HSQL = HSQL & ":" & Fecha.Substring(4, 2)
        End If

        Return HSQL
    End Function

End Module
