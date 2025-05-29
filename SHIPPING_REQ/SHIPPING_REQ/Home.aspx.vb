Imports System.Data.SqlClient

Public Class Home
    Inherits System.Web.UI.Page
    Dim connStr As String = "Data Source=192.168.1.7;Initial Catalog=CS_DATA;User Id=sa;Password=p@ssw0rd;"
    Dim conn As New SqlConnection(connStr)
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub btnLogin_Click(ByVal sender As Object, ByVal e As EventArgs)

        Dim cmd As New SqlCommand("SELECT * FROM [TEST_DATA].[dbo].[Users_Ship] WHERE Username = @Username AND Password = @Password", conn)
        cmd.Parameters.AddWithValue("@Username", txtUsername.Text)
        cmd.Parameters.AddWithValue("@Password", txtPassword.Text) ' *อย่าใช้ plaintext password จริง ให้แฮชในโปรเจกต์จริง*

        conn.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader()

        If reader.HasRows Then
            ' ล็อกอินสำเร็จ
            Session("Username") = txtUsername.Text
            Response.Redirect("ShipReq.aspx")
        Else
            lblMessage.Text = "Invalid username or password"
        End If

        conn.Close()
    End Sub

End Class