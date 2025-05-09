Imports System.Data.OleDb
Imports System.Threading.Tasks

Public Class Form4
    Dim connectionstring As String
    Dim dbconnection As OleDbConnection
    Dim dbadapter As New OleDbDataAdapter
    Dim dbdataset As New DataSet
    Private placeholderText As String = "Enter Product Name"

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Product List"
        DisplayNorm()
        ListBox1.Visible = False
    End Sub

    Private Sub DisplayNorm()
        Label1.Text = "Product ID:"
        Label2.Text = "Product Name:"
        Label3.Text = "Retail Price:"
        Label4.Text = "Stocks:"
        Button1.Text = "Add"
        Button2.Text = "Edit"
        Button3.Text = "Delete"
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox1.Text = vbNullString
        TextBox2.Text = vbNullString
        TextBox3.Text = vbNullString
        TextBox4.Text = vbNullString
        TextBox5.Text = vbNullString
        DataGridView1.AllowUserToAddRows = False
        connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= POS.mdb"
        dbconnection = New OleDbConnection(connectionstring)
        Try
            dbconnection.Open()
            dbdataset.Clear()
            dbadapter = New OleDbDataAdapter("Select * from Product", connectionstring)
            dbadapter.Fill(dbdataset, "Product")
            DataGridView1.DataSource = dbdataset.Tables("Product").DefaultView
            dbconnection.Close()
            Label5.Text = "• Connected"
            Label5.ForeColor = Color.LimeGreen
        Catch ex As Exception
            Label5.Text = "Disconnected"
            Label5.ForeColor = Color.Red
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "Add" Then
            Button1.Enabled = False
            Button2.Text = "Save"
            Button3.Text = "Cancel"
            TextBox2.Text = vbNullString
            TextBox3.Text = vbNullString
            TextBox4.Text = vbNullString
            TextBox1.Text = GetNextProductID().ToString()
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            DataGridView1.AllowUserToAddRows = True
        ElseIf Button1.Text = "Save" Then
            dbconnection.Open()
            Dim dbcommand As New OleDbCommand("UPDATE Product SET PRODUCT_NAME = '" & TextBox2.Text.Trim & "', PRICE = '" & TextBox3.Text.Trim & "',STOCKS = '" & TextBox4.Text.Trim & "',PROVIDER = '" & TextBox5.Text.Trim & "' WHERE Product_ID = " & TextBox1.Text & " ", dbconnection)
            dbcommand.ExecuteNonQuery()
            dbcommand.Dispose()
            dbconnection.Close()
            TextBox5.Clear()
            DisplayNorm()
            MsgBox("Record Save", vbInformation, "Update")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "Edit" Then
            Button2.Enabled = False
            Button1.Text = "Save"
            Button3.Text = "Cancel"
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
        ElseIf Button2.Text = "Save" Then
            If (TextBox2.Text <> vbNullString And TextBox3.Text <> vbNullString And TextBox4.Text <> vbNullString) Then
                dbconnection.Open()
                Dim dbcommand As New OleDbCommand("INSERT INTO Product(Product_ID, PRODUCT_NAME, PRICE, STOCKS, PROVIDER) VALUES (" & TextBox1.Text.Trim & ",'" & TextBox2.Text.Trim & "','" & TextBox3.Text.Trim & "','" & TextBox4.Text.Trim & "','" & TextBox5.Text.Trim & "')", dbconnection)
                dbcommand.ExecuteNonQuery()
                dbcommand.Dispose()
                dbconnection.Close()
                DisplayNorm()
                MsgBox("Successful", vbInformation, "Add Record")
            ElseIf (TextBox2.Text = vbNullString) Then
                MsgBox("Enter Product Name", vbCritical, "Missing")
            ElseIf (TextBox3.Text = vbNullString) Then
                MsgBox("Enter Price", vbCritical, "Missing")
            ElseIf (TextBox4.Text = vbNullString) Then
                MsgBox("Enter how many stocks", vbCritical, "Missing")
            End If
        ElseIf Button2.Text = "Cancel" Then
            TextBox5.Clear()
            DisplayNorm()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Button3.Text = "Delete" Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Dim Response = MsgBox("Are you sure to delete the record?", vbYesNo, "Confirmation")
            If Response = vbYes Then
                dbconnection.Open()
                Dim dbcommand As New OleDbCommand("DELETE FROM Product WHERE PRODUCT_ID = " & TextBox1.Text & " ", dbconnection)
                dbcommand.ExecuteNonQuery()
                dbcommand.Dispose()
                dbconnection.Close()
                DisplayNorm()
                MsgBox("Record was permanently deleted", vbInformation, "Successful")
            Else
                DisplayNorm()
                TextBox5.Clear()
            End If
        Else
            DisplayNorm()
            TextBox5.Clear()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Text = DataGridView1.CurrentRow.Cells(0).Value
        TextBox2.Text = DataGridView1.CurrentRow.Cells(1).Value
        TextBox3.Text = DataGridView1.CurrentRow.Cells(2).Value
        TextBox4.Text = DataGridView1.CurrentRow.Cells(3).Value
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim searchValue As String = TextBox6.Text.Trim()

        If searchValue = "" OrElse searchValue = placeholderText Then
            MsgBox("Please enter Product Name.", MsgBoxStyle.Exclamation, "Empty Search")
            Exit Sub
        End If

        Dim view As DataView = dbdataset.Tables("Product").DefaultView
        view.RowFilter = "PRODUCT_NAME = '" & searchValue.Replace("'", "''") & "'"

        If view.Count = 0 Then
            MsgBox("No record found.", MsgBoxStyle.Information, "Search Result")
            TextBox6.Clear()
            view.RowFilter = ""
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If Not dbdataset.Tables.Contains("Product") Then Exit Sub

        Dim filterText As String = TextBox6.Text.Trim()
        Dim view As DataView = dbdataset.Tables("Product").DefaultView
        ListBox1.Items.Clear()

        If filterText <> "" AndAlso filterText <> placeholderText Then
            view.RowFilter = "PRODUCT_NAME LIKE '%" & filterText.Replace("'", "''") & "%'"
            For Each row As DataRowView In view
                ListBox1.Items.Add(row("PRODUCT_NAME").ToString())
            Next
            ListBox1.Visible = ListBox1.Items.Count > 0
        Else
            view.RowFilter = ""
            ListBox1.Visible = False
        End If
    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        If ListBox1.SelectedItem IsNot Nothing Then
            TextBox6.Text = ListBox1.SelectedItem.ToString()
            ListBox1.Visible = False
        End If
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        If TextBox6.Text = placeholderText Then
            TextBox6.Text = ""
            TextBox6.ForeColor = Color.Black
        End If
    End Sub

    Private Async Sub TextBox6_LostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        If String.IsNullOrWhiteSpace(TextBox6.Text) Then
            TextBox6.Text = placeholderText
            TextBox6.ForeColor = Color.Gray
        End If
        Await Task.Delay(200)
        ListBox1.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form8.Show()
        Me.Hide()
        DisplayNorm()
        Me.Close()
    End Sub

    Private Function GetNextProductID() As Integer
        Dim nextID As Integer = 1
        Try
            dbconnection.Open()
            Dim cmd As New OleDbCommand("SELECT MAX(Product_ID) FROM Product", dbconnection)
            Dim result = cmd.ExecuteScalar()
            If Not IsDBNull(result) Then
                nextID = Convert.ToInt32(result) + 1
            End If
            dbconnection.Close()
        Catch ex As Exception
            MsgBox("Error getting next Product ID: " & ex.Message, MsgBoxStyle.Critical)
            dbconnection.Close()
        End Try
        Return nextID
    End Function

End Class
