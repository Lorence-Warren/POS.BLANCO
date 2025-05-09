Imports System.Data.OleDb

Public Class Form6
    Dim connectionstring As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=POS.mdb"
    Dim dbconnection As New OleDbConnection(connectionstring)
    Dim dbadapter As New OleDbDataAdapter
    Dim dbdataset As New DataSet
    Private placeholderText As String = "Enter ID No."

    Private previousDate As Date = Date.MinValue
    Private previousIncome As Decimal = 0

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Dashboard"
        TextBox1.Text = placeholderText
        TextBox1.ForeColor = Color.Gray

        Try
            dbconnection.Open()

            ' Load all transactions into DataGridView
            dbdataset.Clear()
            dbadapter = New OleDbDataAdapter("SELECT * FROM [Transaction]", dbconnection)
            dbadapter.Fill(dbdataset, "Transaction")
            Dim view As DataView = dbdataset.Tables("Transaction").DefaultView
            view.Sort = "ORDER_ID ASC"
            DataGridView1.DataSource = view

            Dim currentDate As Date = Date.Today
            Dim yesterdayDate As Date = currentDate.AddDays(-1)
            Dim monthStartDate As New Date(currentDate.Year, currentDate.Month, 1)
            Dim todayTotal As Decimal = 0
            Dim yesterdayTotal As Decimal = 0
            Dim monthlyTotal As Decimal = 0

            Dim seenToday As New HashSet(Of String)
            Dim seenYesterday As New HashSet(Of String)
            Dim seenMonth As New HashSet(Of String)

            ' Check if yesterday's income is already saved
            Dim checkCmd As New OleDbCommand("SELECT COUNT(*) FROM Income WHERE [DATE] = ?", dbconnection)
            checkCmd.Parameters.AddWithValue("?", yesterdayDate)
            Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

            ' Read all transactions only once for efficiency
            Dim cmdAll As New OleDbCommand("SELECT [TRANSACTION_ID], [TOTAL], [ORDER_DATE] FROM [Transaction]", dbconnection)
            Dim reader As OleDbDataReader = cmdAll.ExecuteReader()

            While reader.Read()
                If Not reader.IsDBNull(0) AndAlso Not reader.IsDBNull(1) AndAlso Not reader.IsDBNull(2) Then
                    Dim transactionId As String = reader("TRANSACTION_ID").ToString()
                    Dim totalAmount As Decimal = Convert.ToDecimal(reader("TOTAL"))
                    Dim orderDate As Date = Convert.ToDateTime(reader("ORDER_DATE"))

                    If orderDate.Date = currentDate AndAlso Not seenToday.Contains(transactionId) Then
                        todayTotal += totalAmount
                        seenToday.Add(transactionId)
                    End If

                    If orderDate.Date = yesterdayDate AndAlso Not seenYesterday.Contains(transactionId) Then
                        yesterdayTotal += totalAmount
                        seenYesterday.Add(transactionId)
                    End If

                    If orderDate.Date >= monthStartDate AndAlso orderDate.Date <= currentDate AndAlso Not seenMonth.Contains(transactionId) Then
                        monthlyTotal += totalAmount
                        seenMonth.Add(transactionId)
                    End If
                End If
            End While
            reader.Close()

            ' Save yesterday’s income if not already in Income table
            ' Save yesterday’s income (even if zero) if not already in Income table
            If count = 0 Then
                SaveIncomeToDatabase(yesterdayDate, yesterdayTotal)
            End If


            ' Set totals in textboxes
            TextBox2.Text = todayTotal.ToString("N2")
            TextBox3.Text = monthlyTotal.ToString("N2")

            dbconnection.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
            If dbconnection.State = ConnectionState.Open Then dbconnection.Close()
        End Try
    End Sub

    Private Sub SaveIncomeToDatabase(incomeDate As Date, incomeAmount As Decimal)
        Try
            Dim cmd As New OleDbCommand("INSERT INTO Income ([DATE], [INCOME]) VALUES (?, ?)", dbconnection)
            cmd.Parameters.AddWithValue("?", incomeDate)
            cmd.Parameters.AddWithValue("?", incomeAmount)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Error saving income: " & ex.Message)
        End Try
    End Sub



    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim searchValue As String = TextBox1.Text.Trim()

        If searchValue = "" OrElse searchValue = placeholderText Then
            MsgBox("Please enter ID No.", MsgBoxStyle.Exclamation, "Empty Search")
            Exit Sub
        End If

        Dim view As DataView = dbdataset.Tables("Transaction").DefaultView
        view.RowFilter = "TRANSACTION_ID = '" & searchValue.Replace("'", "''") & "'"

        If view.Count = 0 Then
            MsgBox("No record found.", MsgBoxStyle.Information, "Search Result")
            TextBox1.Clear()
            view.RowFilter = "" ' Reset filter
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If dbdataset.Tables.Contains("Transaction") Then
            Dim view As DataView = dbdataset.Tables("Transaction").DefaultView
            If TextBox1.Text.Trim() = "" OrElse TextBox1.Text = placeholderText Then
                view.RowFilter = "" ' Clear filter to show all records
            End If
        End If
    End Sub



    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        If TextBox1.Text = placeholderText Then
            TextBox1.Text = ""
            TextBox1.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            TextBox1.Text = placeholderText
            TextBox1.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form8.Show()
        Me.Close()
        Me.Hide()
    End Sub
End Class
