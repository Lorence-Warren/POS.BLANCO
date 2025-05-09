Imports System.Data.OleDb
Public Class Form5
    Dim connectionstring As String
    Dim dbconnection As OleDbConnection
    Dim dbadapter As New OleDbDataAdapter
    Dim dbdataset As New DataSet
    Private placeholderText As String = "Enter Lastname"
    Dim lastNameList As New List(Of String)

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Customer Registration"
        DisplayNorm()
        ListBox1.Visible = False
        LoadLastNames()
    End Sub

    Private Sub DisplayNorm()
        Label1.Text = "Customer ID:"
        Label2.Text = "Firstname:"
        Label3.Text = "Lastname:"
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
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        DataGridView1.AllowUserToAddRows = False
        connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= POS.mdb"
        dbconnection = New OleDbConnection(connectionstring)
        Try
            dbconnection.Open()
            dbdataset.Clear()
            dbadapter = New OleDbDataAdapter("Select * from Customer", connectionstring)
            dbadapter.Fill(dbdataset, "Customer")
            DataGridView1.DataSource = dbdataset.Tables("Customer").DefaultView
            dbconnection.Close()
            Label5.Text = "• Connected"
            Label5.ForeColor = Color.LimeGreen
        Catch ex As Exception
            Label5.Text = "Disconnected"
            Label5.ForeColor = Color.Red
        End Try
    End Sub

    Private Sub LoadLastNames()
        lastNameList.Clear()
        dbconnection = New OleDbConnection(connectionstring)
        Try
            dbconnection.Open()
            Dim cmd As New OleDbCommand("SELECT DISTINCT LASTNAME FROM Customer", dbconnection)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            While reader.Read()
                lastNameList.Add(reader("LASTNAME").ToString())
            End While
            reader.Close()
            dbconnection.Close()
        Catch ex As Exception
            MsgBox("Error loading last names: " & ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "Add" Then
            Button1.Enabled = False
            Button2.Text = "Save"
            Button3.Text = "Cancel"
            TextBox2.Text = vbNullString
            TextBox3.Text = vbNullString
            TextBox4.Text = vbNullString
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            DataGridView1.AllowUserToAddRows = True
        ElseIf Button1.Text = "Save" Then
            dbconnection.Open()
            Dim dbcommand As New OleDbCommand("UPDATE Customer SET FIRSTNAME = '" & TextBox2.Text.Trim & "', LASTNAME = '" & TextBox3.Text.Trim & "', EMAIL = '" & TextBox4.Text.Trim & "', ADDRESS = '" & TextBox5.Text.Trim & "' WHERE Customer_ID = " & TextBox1.Text & " ", dbconnection)
            dbcommand.ExecuteNonQuery()
            dbcommand.Dispose()
            dbconnection.Close()
            DisplayNorm()
            MsgBox("Record Save", vbInformation, "Update")
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
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
                Dim dbcommand As New OleDbCommand("INSERT INTO Customer(FIRSTNAME, LASTNAME, EMAIL, ADDRESS) VALUES ('" + TextBox2.Text.Trim + "','" + TextBox3.Text.Trim + "','" + TextBox4.Text.Trim + "','" + TextBox5.Text.Trim + "')", dbconnection)
                dbcommand.ExecuteNonQuery()
                dbcommand.Dispose()
                dbconnection.Close()
                DisplayNorm()
                MsgBox("Successful", vbInformation, "Add Record")
            ElseIf (TextBox2.Text = vbNullString) Then
                MsgBox("Enter Firstname", vbCritical, "Missing")
            ElseIf (TextBox3.Text = vbNullString) Then
                MsgBox("Enter Lastname", vbCritical, "Missing")
            ElseIf (TextBox4.Text = vbNullString) Then
                MsgBox("Enter Email", vbCritical, "Missing")
            End If
        ElseIf Button2.Text = "Cancel" Then
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
                Dim dbcommand As New OleDbCommand("DELETE FROM Customer WHERE Customer_ID = " & TextBox1.Text & " ", dbconnection)
                dbcommand.ExecuteNonQuery()
                dbcommand.Dispose()
                dbconnection.Close()
                DisplayNorm()
                MsgBox("Record was permanently deleted", vbInformation, "Successful")
            Else
                DisplayNorm()
            End If
        Else
            DisplayNorm()
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim searchValue As String = TextBox6.Text.Trim()

        If searchValue = "" OrElse searchValue = placeholderText Then
            MsgBox("Please enter Lastname.", MsgBoxStyle.Exclamation, "Empty Search")
            Exit Sub
        End If

        Dim view As DataView = dbdataset.Tables("Customer").DefaultView
        view.RowFilter = "LASTNAME = '" & searchValue.Replace("'", "''") & "'"

        If view.Count = 0 Then
            MsgBox("No record found.", MsgBoxStyle.Information, "Search Result")
            TextBox6.Clear()
            view.RowFilter = "" ' Reset filter
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        TextBox1.Text = DataGridView1.CurrentRow.Cells(0).Value
        TextBox2.Text = DataGridView1.CurrentRow.Cells(1).Value
        TextBox3.Text = DataGridView1.CurrentRow.Cells(2).Value
        TextBox4.Text = DataGridView1.CurrentRow.Cells(3).Value
    End Sub

    ' AUTOCOMPLETE FEATURE

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Dim view As DataView = dbdataset.Tables("Customer").DefaultView

        If TextBox6.Text.Trim() = "" Or TextBox6.Text = placeholderText Then
            ListBox1.Visible = False
            view.RowFilter = "" ' Show all records when cleared
            Exit Sub
        End If

        ' Filter autocomplete suggestions
        Dim filtered = lastNameList.Where(Function(name) name.ToLower().Contains(TextBox6.Text.Trim.ToLower())).ToList()
        If filtered.Count > 0 Then
            ListBox1.Items.Clear()
            For Each name As String In filtered
                ListBox1.Items.Add(name)
            Next
            ListBox1.Visible = True
            ListBox1.Top = TextBox6.Bottom
            ListBox1.Left = TextBox6.Left
            ListBox1.Width = TextBox6.Width
            ListBox1.BringToFront()
        Else
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

    Private Sub TextBox6_LostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        If String.IsNullOrWhiteSpace(TextBox6.Text) Then
            TextBox6.Text = placeholderText
            TextBox6.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form8.Show()
        Me.Hide()
    End Sub

End Class
