Imports System.IO
Imports System.Net
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Form1
    Public conn As New SqlConnection
    Public command As New SqlCommand
    Public stateConnection As String
    Public Data As String


    Sub Link()
        Try

            conn.ConnectionString = "Data Source=MPC-JMV2\SQLEXPRESS;Initial Catalog=Okasama2;Integrated Security=True"
            conn.Open()
            stateConnection = "conectado"
        Catch ex As Exception
            stateConnection = "desconectado"
        End Try
    End Sub

    Sub GetData()

        Try
            'Creation initial request

            Dim response As HttpWebResponse = Nothing
            Dim reader As StreamReader
            Dim request As HttpWebRequest = DirectCast(WebRequest.Create("https://pidetu.cl/wp-json/wc/v3/orders"), HttpWebRequest)
            Dim client As String = "ck_5b6f1ff6731fc02c51eb1efbd7713bbbcdf53d00"
            Dim secret As String = "cs_5729f25372e8317708661055d30bc80a54478a73"
            Dim tempcookies As New CookieContainer
            Dim Credentials As String = client + ":" + secret
            Dim data As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(Credentials)
            Dim CredentialsBase64 = System.Convert.ToBase64String(data)
            request.Headers.Add("Authorization", "Basic " + CredentialsBase64)
            request.ContentType = "application/json; charset=utf-8"

            response = DirectCast(request.GetResponse(), HttpWebResponse)
            reader = New StreamReader(response.GetResponseStream())
            Dim OrdersString As String = reader.ReadToEnd
            Dim result = JsonConvert.DeserializeObject(Of Object)(OrdersString)

            reader.Close()
            Dim Orders = JArray.Parse(OrdersString)

            For Each orden As Object In result

                Try
                    Dim query As String = "SELECT * FROM Ordenes WHERE id = " & orden("id")
                    Dim resultQuery = New SqlClient.SqlCommand(query, conn).ExecuteScalar()

                    If resultQuery = Nothing Then

                        Dim QueryDeleteData As String = "BEGIN TRAN;DELETE FROM Envio WHERE id_orden = " & orden("id") & ";DELETE FROM Factura WHERE id_orden = " & orden("id") & ";DELETE FROM MetaData WHERE id_orden = " & orden("id") & "DELETE FROM Ordenes WHERE id = " & orden("id") & "; COMMIT"
                        command = New SqlClient.SqlCommand(QueryDeleteData, conn)
                        Dim Resultado = command.ExecuteNonQuery()

                        Dim QueryInsert = "INSERT INTO [Okasama2].[dbo].[Ordenes] (id,parent_id,status,currency,version,prices_include_tax,date_created,date_modified,discount_total,discount_tax,shipping_total,shipping_tax,cart_tax,total,total_tax,customer_id,order_key,payment_method,payment_method_title,transaction_id,customer_ip_address,customer_user_agent,created_via,customer_note,date_completed,date_paid,cart_hash,number,date_created_gmt,date_modified_gmt,date_completed_gmt,date_paid_gmt,currency_symbol) VALUES ("
                        QueryInsert += orden("id") & ","
                        QueryInsert += orden("parent_id") & ","
                        QueryInsert += "'" & orden("status") & "',"
                        QueryInsert += "'" & orden("currency") & "',"
                        QueryInsert += "'" & orden("version") & "',"
                        QueryInsert += "'" & orden("prices_include_tax") & "',"
                        QueryInsert += "'" & orden("date_created") & "',"
                        QueryInsert += "'" & orden("date_modified") & "',"
                        QueryInsert += "'" & orden("discount_total") & "',"
                        QueryInsert += "'" & orden("discount_tax") & "',"
                        QueryInsert += "'" & orden("shipping_total") & "',"
                        QueryInsert += "'" & orden("shipping_tax") & "',"
                        QueryInsert += "'" & orden("cart_tax") & "',"
                        QueryInsert += "'" & orden("total") & "',"
                        QueryInsert += "'" & orden("total_tax") & "',"
                        QueryInsert += "'" & orden("customer_id") & "',"
                        QueryInsert += "'" & orden("order_key") & "',"
                        QueryInsert += "'" & orden("payment_method") & "',"
                        QueryInsert += "'" & orden("payment_method_title") & "',"
                        QueryInsert += "'" & orden("transaction_id") & "',"
                        QueryInsert += "'" & orden("customer_ip_address") & "',"
                        QueryInsert += "'" & orden("customer_user_agent") & "',"
                        QueryInsert += "'" & orden("created_via") & "',"
                        QueryInsert += "'" & orden("customer_note") & "',"
                        QueryInsert += "'" & orden("date_completed") & "',"
                        QueryInsert += "'" & orden("date_paid") & "',"
                        QueryInsert += "'" & orden("cart_hash") & "',"
                        QueryInsert += "'" & orden("number") & "',"
                        QueryInsert += "'" & orden("date_created_gmt") & "',"
                        QueryInsert += "'" & orden("date_modified_gmt") & "',"
                        QueryInsert += "'" & orden("date_completed_gmt") & "',"
                        QueryInsert += "'" & orden("date_paid_gmt") & "',"
                        QueryInsert += "'" & orden("currency_symbol") & "')"

                        command = New SqlClient.SqlCommand(QueryInsert, conn)
                        Dim resultQueryInsert = command.ExecuteNonQuery()
                    End If
                Catch ex As Exception

                End Try

            Next

        Catch ex As Exception

        End Try

    End Sub

    Sub PrintData()
        Dim ds As New DataSet
        Dim query As String = "SELECT 
        id as ID,
        total as TOTAL,
        payment_method_title as [METODO DE PAGO],
        status as ESTADO,
        currency as [TIPO DE DIVISA],
        date_created as [FECHA DE CREACIÓN]
        FROM Ordenes ORDER BY date_created DESC"

        ds.Tables.Add("tabla")
        Dim da = New System.Data.SqlClient.SqlDataAdapter(query, conn)
        da.Fill(ds.Tables("tabla"))
        DataGridView1.DataSource = ds.Tables("tabla")

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Load()
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 300000
        Me.Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Load()
    End Sub

    Private Sub DataGridOrdes_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        conn.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    Sub Load()
        Link()
        GetData()
        PrintData()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
