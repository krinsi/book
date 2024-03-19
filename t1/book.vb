Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class book
    Dim connection As OleDbConnection
    Private Sub book_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connection = New OleDbConnection("Provider=oraOLEDB.oracle;Data Source=localhost;User Id=system;Password=int1;")
        FillDataGrid()
    End Sub

    Private Sub FillDataGrid()
        connection.Open()
        Dim adapter As New OleDbDataAdapter("Select *From record", connection)
        Dim ds As New DataSet
        adapter.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        connection.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FillDataGrid()
        Try
            connection.Open()
            'Dim connection As New OleDbConnection("Data source=localhost,user Id=system,password=int1,provider=oraOLEDB.oracle")
            Dim sql As String = "insert into record(r_id,b_id,u_id,b_date) values(?,?,?,?)"
            Dim cmd As New OleDbCommand(sql, connection)
            cmd.Parameters.AddWithValue("?", CInt(TextBox1.Text))
            cmd.Parameters.AddWithValue("?", CInt(TextBox2.Text))
            cmd.Parameters.AddWithValue("?", CInt(TextBox3.Text))

            Dim startDate As DateTime = DateTimePicker1.Value
            ' cmd.Parameters.AddWithValue("@b_date", startDate)

            cmd.Parameters.AddWithValue("?", startDate)

            Dim affectedrow = cmd.ExecuteNonQuery()
            connection.Close()

            If (affectedrow >= 1) Then
                MessageBox.Show("Inserted")
                FillDataGrid()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            connection.Close()

        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            connection.Open()
            Dim adapter As New OleDbDataAdapter("SELECT * FROM record WHERE r_id=12", connection)
            Dim ds As New DataSet
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            connection.Close()
        End Try
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            connection.Open()
            Dim adapter As New OleDbDataAdapter("SELECT * FROM book WHERE b_id>=5", connection)
            Dim ds As New DataSet
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            connection.Close()
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            connection.Open()
            Dim adapter As New OleDbDataAdapter("SELECT * FROM book WHERE b_id in(select b_id from record where p_year=2019)", connection)
            Dim ds As New DataSet
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            connection.Close()
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            connection.Open()
            Dim adapter As New OleDbDataAdapter("SELECT b.b_id,b.title,b.author,b.p_year,b.qty,COUNT(br.r_id) AS Borrow_Count
FROM book b LEFT JOIN record br on b.b_id = br.b_id group by  b.b_id, b.title, b.author, b.gender, b.p_year, b.qty
HAVING COUNT(br.r_id) >= 1", connection)
            Dim ds As New DataSet
            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            connection.Close()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FillDataGrid()
        Try
            connection.Open()
            Dim sql As String = "delete from book where r_id=?"
            Dim cmd As New OleDbCommand(sql, connection)
            cmd.Parameters.AddWithValue("?", CInt(TextBox1.Text))

            Dim affectedrow = cmd.ExecuteNonQuery()
            connection.Close()

            If (affectedrow >= 1) Then
                MessageBox.Show("Deleted")
                FillDataGrid()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            connection.Close()
        End Try
    End Sub
End Class