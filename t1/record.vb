Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class record

    Dim connection As OleDbConnection
    Private Sub record_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connection = New OleDbConnection("Provider=oraOLEDB.oracle;Data Source=localhost;User Id=system;Password=int1;")
        FillDataGrid()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FillDataGrid()
        Try
            connection.Open()

            Dim gender As String = ""
            If RadioButton1.Checked Then
                gender = "Male"
            ElseIf RadioButton2.Checked Then
                gender = "Female"
            End If

            'Dim quantity As Integer
            'If Integer.TryParse(ComboBox1.SelectedItem.ToString(), quantity) Then

            Dim sql As String = "insert into book(b_id,title,author,p_year,qty,gender) values(?,?,?,?,?,?)"

                Dim cmd As New OleDbCommand(sql, connection)
            cmd.Parameters.AddWithValue("?", CInt(TextBox1.Text))
            cmd.Parameters.AddWithValue("?", ComboBox1.SelectedItem)
            cmd.Parameters.AddWithValue("?", TextBox3.Text)
                cmd.Parameters.AddWithValue("?", CInt(TextBox4.Text))
            cmd.Parameters.AddWithValue("?", CInt(TextBox5.Text))
            'cmd.Parameters.AddWithValue("?", quantity)
            'cmd.Parameters.AddWithValue("?", ComboBox1.SelectedValue)
            cmd.Parameters.AddWithValue("?", gender)

                Dim affectedrow = cmd.ExecuteNonQuery()
                connection.Close()

                If (affectedrow >= 1) Then
                    MessageBox.Show("Inserted")
                    FillDataGrid()
                End If
            'Else
            'MessageBox.Show("Invalid quantity value selected.")
            'End If
        Catch ex As Exception
            MsgBox(ex.Message)
            connection.Close()

        End Try
    End Sub

    Private Sub FillDataGrid()
        connection.Open()
        Dim adapter As New OleDbDataAdapter("Select *From book", connection)
        Dim ds As New DataSet
        adapter.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        connection.Close()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FillDataGrid()
        Try
            connection.Open()
            Dim gender As String = ""
            If RadioButton1.Checked Then
                gender = "Male"
            ElseIf RadioButton2.Checked Then
                gender = "Female"
            End If
            Dim sql As String = "update book set title=?,author=?,p_year=?,qty=?,gender=? where b_id=?"
            Dim cmd As New OleDbCommand(sql, connection)
            cmd.Parameters.AddWithValue("?", ComboBox1.SelectedItem)
            cmd.Parameters.AddWithValue("?", TextBox3.Text)
            cmd.Parameters.AddWithValue("?", CInt(TextBox4.Text))
            cmd.Parameters.AddWithValue("?", CInt(TextBox5.Text))
            'cmd.Parameters.AddWithValue("?", ComboBox2.SelectedIndex)
            cmd.Parameters.AddWithValue("?", gender)
            cmd.Parameters.AddWithValue("?", CInt(TextBox1.Text))

            Dim affectedrow = cmd.ExecuteNonQuery()
            connection.Close()

            If (affectedrow >= 1) Then
                MessageBox.Show("Updated")
                FillDataGrid()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            connection.Close()
        End Try
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try
            TextBox1.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString()
            ComboBox1.Text = DataGridView1.SelectedRows(0).Cells(1).Value.ToString()
            TextBox3.Text = DataGridView1.SelectedRows(0).Cells(2).Value.ToString()
            TextBox4.Text = DataGridView1.SelectedRows(0).Cells(3).Value.ToString()
            TextBox5.Text = DataGridView1.SelectedRows(0).Cells(4).Value.ToString()

            'ComboBox2.Text = DataGridView1.SelectedRows(0).Cells(4).Value.ToString()
            Dim gender As String = DataGridView1.SelectedRows(0).Cells(5).Value.ToString()
            If gender = "Male" Then
                RadioButton1.Checked = True
            ElseIf gender = "Female" Then
                RadioButton2.Checked = True
            End If
        Catch ex As Exception

        End Try

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FillDataGrid()
        Try
            connection.Open()
            Dim sql As String = "delete from book where b_id=?"
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