Public Class Form1
    Public Shared Dir_App As String = Mid(Application.ExecutablePath, 1, Application.ExecutablePath.Length - Diagnostics.Process.GetCurrentProcess.ProcessName.Length - 4)
    Public Shared readersql As OleDb.OleDbDataReader 'Global estatica compartida entre todos los archivos de la solucion.
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Conexion As New OleDb.OleDbConnection
        Conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Dir_App & "app.dll;Jet OLEDB:Database Password=;"
        Dim Comando As New OleDb.OleDbCommand
        Comando.Connection = Conexion
        Conexion.Open()

        Try
            Comando.CommandText = "insert into cliente (id, Nombre, Apellidos) values (1, 'Cristian', 'Diaz Porter');"
            Comando.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message) ' Descripcion Excepcion
            MsgBox(ex.StackTrace) ' Ruta de ejecucion hasta excepcion
            MsgBox(ex.ToString) ' Combinacion de los dos anteriores
        End Try



        Comando.CommandText = "select * from cliente;"
        readersql = Comando.ExecuteReader()
        Dim nombre
        Dim Apellido
        While readersql.Read()
            nombre = readersql(1)
            Apellido = readersql(2)
        End While
        readersql.Close()


    End Sub
End Class
