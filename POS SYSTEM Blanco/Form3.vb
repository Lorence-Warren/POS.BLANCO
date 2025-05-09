Imports System.Data.OleDb

Public Class Form3
    Dim connectionstring As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=POS.mdb"
    Dim dbconnection As New OleDbConnection(connectionstring)
    Dim dbadapter As New OleDbDataAdapter
    Dim dbdataset As New DataSet
    Dim dbdatasetCustomer As New DataSet
    Dim dbdatasetProduct As New DataSet



    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Point of Sale System"
        SetupCartGrid()
        DisplayNorm()
        DateTimePicker1.Value = Date.Today
        DateTimePicker1.Enabled = False
        ComboBox1.SelectedItem = "Cash"
        TextBox6.Text = "Walk-In"
        If String.IsNullOrWhiteSpace(TextBox6.Text) Then
            TextBox6.Text = "Walk-In"
            TextBox6.ForeColor = Color.Black ' Placeholder color
        End If

    End Sub

    Private Sub SetupCartGrid()
        With DataGridView2
            .Columns.Clear()
            .Columns.Add("ProductName", "Product Name")
            .Columns.Add("Price", "Price")
            .Columns.Add("Qty", "Quantity")
            .Columns.Add("Subtotal", "Subtotal") ' Add Subtotal column

        End With
    End Sub
    Private suppressTextChanged As Boolean = False


    Private Sub DisplayNorm()
        ' Reset UI and labels
        Label1.Text = "Product Info."
        Label2.Text = "Product ID:"
        Label3.Text = "Product Name:"
        Label4.Text = "Price:"
        Label5.Text = "Qty."
        Label6.Text = "Customer's Name:"
        Label7.Text = "Customer ID:"
        Label8.Text = "Transaction ID:"
        Label10.Text = "Transaction"
        Label11.Text = "Date:"
        Label16.Text = "Stocks:"
        Label19.Visible = False
        TextBox10.ReadOnly = False

        ' Disable and clear input fields
        For Each ctrl In {TextBox1, TextBox2, TextBox3, NumericUpDown1, TextBox5, TextBox6, TextBox8, TextBox12}
            ctrl.Enabled = False
        Next
        For Each tb In {TextBox2, TextBox3, TextBox4, TextBox5, TextBox6, TextBox8, TextBox9, TextBox10, TextBox11, TextBox12}
            tb.Clear()
        Next
        NumericUpDown1.Value = 1

        ' Set button states
        Button1.Text = "New"
        Button2.Text = "Remove"
        Button3.Text = "Cancel"
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        ListBox1.Enabled = False
        ListBox1.Visible = False
        ComboBox1.Enabled = False
        TextBox4.Visible = False
        TextBox10.Enabled = False

        ' Load transactions
        Try
            dbconnection.Open()
            dbdataset.Clear()
            dbadapter = New OleDbDataAdapter("SELECT * FROM [Transaction]", dbconnection)
            dbadapter.Fill(dbdataset, "Transaction")
            Label15.Text = "• Connected"
            Label15.ForeColor = Color.LimeGreen
        Catch ex As Exception
            Label15.Text = "Disconnected"
            Label15.ForeColor = Color.Red
        Finally
            dbconnection.Close()
        End Try
    End Sub
    Private Function GenerateTransactionID() As String
        Dim newID As Integer = 1
        Try
            dbconnection.Open()
            Dim cmd As New OleDbCommand("SELECT MAX(VAL(MID(TRANSACTION_ID, 5))) FROM [Transaction] WHERE TRANSACTION_ID LIKE 'VVA-%'", dbconnection)
            Dim result = cmd.ExecuteScalar()

            If Not IsDBNull(result) AndAlso result IsNot Nothing Then
                newID = Convert.ToInt32(result) + 1
            End If
        Catch ex As Exception
            MsgBox("Error generating transaction ID: " & ex.Message)
        Finally
            dbconnection.Close()
        End Try

        Return "VVA-" & newID.ToString()
    End Function
    ' TextBox1 (PRODUCT_ID) - Ensure no ListBox1 shows here
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If suppressTextChanged Then Exit Sub

        ' Hide ListBox1 when typing in TextBox1 (PRODUCT_ID)
        ListBox1.Visible = False
        NumericUpDown1.Enabled = True


        ' Ensure TextBox1 only interacts with the database for PRODUCT_ID
        Try
            If dbconnection.State <> ConnectionState.Closed Then dbconnection.Close()
            dbconnection.Open()

            dbdatasetProduct.Clear()
            dbadapter = New OleDbDataAdapter("SELECT * FROM PRODUCT WHERE PRODUCT_ID = ?", dbconnection)
            dbadapter.SelectCommand.Parameters.AddWithValue("?", TextBox1.Text.Trim())
            dbadapter.Fill(dbdatasetProduct, "PRODUCT")

            If dbdatasetProduct.Tables("PRODUCT").Rows.Count > 0 Then
                Dim row As DataRow = dbdatasetProduct.Tables("PRODUCT").Rows(0)
                ' Update TextBoxes based on the PRODUCT_ID
                TextBox2.Text = row("PRODUCT_NAME").ToString()
                TextBox12.Text = row("STOCKS").ToString()
                TextBox3.Text = row("PRICE").ToString()
            Else
                ' Clear fields if no matching PRODUCT_ID is found
                TextBox2.Clear()
                TextBox12.Clear()
                TextBox3.Clear()
            End If

        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        Finally
            If dbconnection.State <> ConnectionState.Closed Then dbconnection.Close()
        End Try
    End Sub

    ' TextBox2 (PRODUCT_NAME) - Show ListBox1 with suggestions here
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If suppressTextChanged Then Exit Sub

        ' If TextBox1 is not empty, hide ListBox1 and exit
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            ListBox1.Visible = False
            Exit Sub
        End If

        ' Show ListBox1 when typing in TextBox2 (PRODUCT_NAME)
        If Not String.IsNullOrWhiteSpace(TextBox2.Text) Then
            Try
                If dbconnection.State <> ConnectionState.Closed Then dbconnection.Close()
                dbconnection.Open()

                dbdatasetProduct.Clear()
                dbadapter = New OleDbDataAdapter("SELECT * FROM PRODUCT WHERE PRODUCT_NAME LIKE ?", dbconnection)
                dbadapter.SelectCommand.Parameters.AddWithValue("?", "%" & TextBox2.Text.Trim() & "%")
                dbadapter.Fill(dbdatasetProduct, "PRODUCT")

                ' Clear ListBox1 and load new suggestions
                ListBox1.Items.Clear()
                For Each row As DataRow In dbdatasetProduct.Tables("PRODUCT").Rows
                    ListBox1.Items.Add(row("PRODUCT_NAME").ToString())
                Next

                ' Show ListBox1 only if it has suggestions
                ListBox1.Visible = ListBox1.Items.Count > 0

            Catch ex As Exception
                MsgBox("Error: " & ex.Message)
            Finally
                If dbconnection.State <> ConnectionState.Closed Then dbconnection.Close()
            End Try
        Else
            ' If TextBox2 is cleared, hide the ListBox
            ListBox1.Visible = False
        End If
    End Sub

    ' Handle selection from ListBox1
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedIndex = -1 Then Exit Sub

        Dim selectedProductName As String = ListBox1.SelectedItem.ToString()

        Try
            suppressTextChanged = True ' Prevent TextBox2 event during DB load

            If dbconnection.State <> ConnectionState.Closed Then dbconnection.Close()
            dbconnection.Open()

            dbadapter = New OleDbDataAdapter("SELECT * FROM PRODUCT WHERE PRODUCT_NAME = ?", dbconnection)
            dbadapter.SelectCommand.Parameters.AddWithValue("?", selectedProductName)
            dbdatasetProduct.Clear()
            dbadapter.Fill(dbdatasetProduct, "PRODUCT")

            If dbdatasetProduct.Tables("PRODUCT").Rows.Count > 0 Then
                Dim row = dbdatasetProduct.Tables("PRODUCT").Rows(0)

                ' Update TextBoxes with selected product data
                TextBox1.Text = row("PRODUCT_ID").ToString()
                TextBox2.Text = row("PRODUCT_NAME").ToString()
                TextBox3.Text = row("PRICE").ToString()
                TextBox12.Text = row("STOCKS").ToString()

                Dim stock As Integer = Convert.ToInt32(row("STOCKS"))
                If stock > 0 Then
                    NumericUpDown1.Maximum = stock
                    NumericUpDown1.Minimum = 1
                    NumericUpDown1.Value = 1
                    NumericUpDown1.Enabled = True
                Else
                    NumericUpDown1.Maximum = 1
                    NumericUpDown1.Minimum = 1
                    NumericUpDown1.Value = 1
                    NumericUpDown1.Enabled = False
                End If
            End If

            ' Hide the ListBox after selection
            ListBox1.Visible = False

        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        Finally
            suppressTextChanged = False ' Allow TextBox2 event to be triggered again
            If dbconnection.State <> ConnectionState.Closed Then dbconnection.Close()
        End Try
    End Sub







    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Try
            dbconnection.Open()
            dbdatasetCustomer.Clear()
            dbadapter = New OleDbDataAdapter("SELECT * FROM CUSTOMER WHERE CUSTOMER_ID = ?", dbconnection)
            dbadapter.SelectCommand.Parameters.AddWithValue("?", TextBox5.Text.Trim())
            dbadapter.Fill(dbdatasetCustomer, "CUSTOMER")

            If dbdatasetCustomer.Tables("CUSTOMER").Rows.Count > 0 Then
                Dim row = dbdatasetCustomer.Tables("CUSTOMER").Rows(0)
                TextBox6.Text = row("FIRSTNAME").ToString() & " " & row("LASTNAME").ToString()
            Else
                TextBox6.Clear()
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        Finally
            dbconnection.Close()
        End Try
    End Sub
    Private Sub TextBox6_Click(sender As Object, e As EventArgs) Handles TextBox6.Click
        ' When the user clicks on TextBox6 and it's the placeholder text ("Walk-In"), clear the text for input
        If TextBox6.Text = "Walk-In" Then
            TextBox6.Text = ""
            TextBox6.ForeColor = Color.Black ' Change text color to black for typing
        End If
    End Sub

    Private Sub TextBox6_Leave(sender As Object, e As EventArgs) Handles TextBox6.Leave
        ' If TextBox6 is empty after the user leaves the field, set it to "Walk-In"
        If String.IsNullOrWhiteSpace(TextBox6.Text) Then
            TextBox6.Text = "Walk-In"
            TextBox6.ForeColor = Color.Black ' Change text color back to gray for placeholder effect
        End If
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox2.Text <> "" AndAlso TextBox3.Text <> "" AndAlso NumericUpDown1.Value > 0 Then
            Dim productName As String = TextBox2.Text
            Dim price As Decimal = Decimal.Parse(TextBox3.Text)
            Dim qty As Integer = CInt(NumericUpDown1.Value)
            Dim stock As Integer = CInt(TextBox12.Text)

            ' Generate and display the Transaction ID in TextBox8
            TextBox8.Text = GenerateTransactionID() ' Automatically display the generated Transaction ID in TextBox8

            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = False

            If stock <= 0 Then
                MsgBox("Insufficient stock.", vbExclamation, "Stock Error")
                Return
            End If

            ' Check if quantity equals current stock exactly
            If qty = stock Then
                MsgBox("Insufficient stock. Quantity matches current stock, leaving no remaining items.", vbExclamation, "Stock Warning")
                Return
            End If

            Dim itemExists As Boolean = False
            For Each row As DataGridViewRow In DataGridView2.Rows
                If Not row.IsNewRow AndAlso row.Cells("ProductName").Value.ToString() = productName Then
                    Dim existingQty As Integer = Convert.ToInt32(row.Cells("Qty").Value)
                    ' Check if adding new quantity exceeds stock
                    If existingQty + qty > stock Then
                        MsgBox("Insufficient stock available.", vbExclamation, "Stock Error")
                        Return
                    End If
                    ' Update quantity and recalculate subtotal
                    row.Cells("Qty").Value = existingQty + qty
                    row.Cells("Subtotal").Value = price * (existingQty + qty) ' Update subtotal
                    itemExists = True
                    Exit For
                End If
            Next

            If Not itemExists Then
                If qty > stock Then
                    MsgBox("Insufficient stock available.", vbExclamation, "Stock Error")
                    Return
                End If
                ' Add new row to DataGridView
                DataGridView2.Rows.Add(productName, price, qty, price * qty) ' Add subtotal as price * qty
            End If

            ' Update stock after adding item to the cart
            stock -= qty
            TextBox12.Text = stock.ToString()
            NumericUpDown1.Maximum = stock
            NumericUpDown1.Minimum = If(stock > 0, 1, 0)
            NumericUpDown1.Value = If(stock > 0, 1, 0)

            ' Calculate the total amount
            CalculateTotal()

            ' Clear and reset the fields for the next item
            TextBox1.Clear() ' Clear product ID
            TextBox2.Clear() ' Clear product name
            TextBox3.Clear() ' Clear price
            TextBox12.Clear() ' Clear stock
            NumericUpDown1.Value = 1 ' Reset NumericUpDown value
            NumericUpDown1.Enabled = False ' Disable NumericUpDown until a product is selected
            TextBox1.Focus() ' Focus on product ID for next entry

        Else
            MsgBox("Please enter or select a valid product first.", vbExclamation, "Missing Info")
        End If
    End Sub






    Private Sub CalculateTotal()
        Dim total As Decimal = 0
        For Each row As DataGridViewRow In DataGridView2.Rows
            If Not row.IsNewRow Then
                ' Get the subtotal from column 4
                Dim subtotal As Decimal = Convert.ToDecimal(row.Cells("Subtotal").Value)
                total += subtotal
            End If
        Next
        TextBox9.Text = total.ToString("F2")
    End Sub


    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        Dim totalAmount As Decimal
        Dim amountPaid As Decimal

        ' Ensure both values are valid
        If Decimal.TryParse(TextBox9.Text, totalAmount) AndAlso Decimal.TryParse(TextBox10.Text, amountPaid) Then
            TextBox11.Text = (amountPaid - totalAmount).ToString("F2")

            ' Check if the tendered amount is sufficient
            If amountPaid < totalAmount Then
                TextBox10.ForeColor = Color.Red
                TextBox11.ForeColor = Color.Red
            Else
                TextBox10.ForeColor = Color.Lime
                TextBox11.ForeColor = Color.Lime
            End If
        Else
            TextBox11.Clear()
        End If
    End Sub
    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        ' Allow control keys (e.g., backspace)
        If Char.IsControl(e.KeyChar) Then
            Return
        End If

        ' Allow digits
        If Char.IsDigit(e.KeyChar) Then
            Return
        End If

        ' Allow only one decimal point
        If e.KeyChar = "."c AndAlso Not TextBox10.Text.Contains(".") Then
            Return
        End If

        ' Block everything else
        e.Handled = True
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView2.Rows.Count = 0 Then
            MsgBox("Your cart is empty. Add items to proceed.", vbExclamation, "No Items")
            Return
        End If

        If ComboBox1.Text = "" Then
            MsgBox("Choose payment method.", vbExclamation, "Missing Info")
            Return
        End If
        If String.IsNullOrWhiteSpace(TextBox6.Text) Then
            MsgBox("Please enter a customer name before proceeding.", vbExclamation, "Customer Name Required")
            Return
        End If
        ' Require card number for non-cash payments
        If ComboBox1.Text.Trim().ToUpper() <> "CASH" Then
            If String.IsNullOrWhiteSpace(TextBox4.Text) Then
                MsgBox("Please enter a card number for non-cash payments.", vbExclamation, "Missing Card Info")
                TextBox4.Focus()
                Return
            End If
        End If

        ' Ensure the user has input the exact tendered amount
        If Not Decimal.TryParse(TextBox9.Text, Nothing) OrElse Not Decimal.TryParse(TextBox10.Text, Nothing) Then
            MsgBox("Invalid amount or tendered input.", vbCritical, "Invalid Data")
            Return
        End If

        Dim totalAmount As Decimal = Decimal.Parse(TextBox9.Text)
        Dim tenderedAmount As Decimal = Decimal.Parse(TextBox10.Text)

        If ComboBox1.Text.Trim().ToUpper() <> "CASH" Then
            ' For non-cash payments, tendered amount must be EXACT
            If tenderedAmount <> totalAmount Then
                MsgBox("For non-cash payments, the amount must be exact.", vbExclamation, "Payment Error")
                Return
            End If
        Else
            ' For cash payments, allow overpayment but not underpayment
            If tenderedAmount < totalAmount Then
                MsgBox("Insufficient amount.", vbExclamation, "Payment Error")
                Return
            End If
        End If

        ' Create a new connection locally so it's fully managed
        Using conn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=POS.mdb")
            Try
                conn.Open()

                ' Loop through the cart items
                Dim isFirstItem As Boolean = True
                For Each row As DataGridViewRow In DataGridView2.Rows
                    If row.IsNewRow Then Continue For

                    ' Insert transaction record for each item
                    Using cmd As New OleDbCommand("INSERT INTO [Transaction] (PRODUCT_NAME, PRICE, QTY, CUSTOMER_ID, CUSTOMER_NAME, TRANSACTION_ID, ORDER_DATE, TOTAL, PAYMENT) VALUES (?,?,?,?,?,?,?,?,?)", conn)
                        cmd.Parameters.AddWithValue("?", row.Cells("ProductName").Value.ToString())
                        cmd.Parameters.AddWithValue("?", Convert.ToDecimal(row.Cells("Price").Value))
                        cmd.Parameters.AddWithValue("?", Convert.ToInt32(row.Cells("Qty").Value))

                        If String.IsNullOrWhiteSpace(TextBox5.Text) Then
                            cmd.Parameters.AddWithValue("?", DBNull.Value)
                        Else
                            cmd.Parameters.AddWithValue("?", TextBox5.Text.Trim())
                        End If

                        If String.IsNullOrWhiteSpace(TextBox6.Text) Then
                            cmd.Parameters.AddWithValue("?", DBNull.Value)
                        Else
                            cmd.Parameters.AddWithValue("?", TextBox6.Text.Trim())
                        End If

                        cmd.Parameters.AddWithValue("?", TextBox8.Text.Trim()) ' TRANSACTION_ID
                        cmd.Parameters.AddWithValue("?", DateTimePicker1.Value) ' ORDER_DATE

                        ' Only add the TOTAL and PAYMENT for the first item
                        If isFirstItem Then
                            cmd.Parameters.AddWithValue("?", Convert.ToDecimal(TextBox9.Text.Trim())) ' TOTAL
                            cmd.Parameters.AddWithValue("?", ComboBox1.Text.Trim()) ' PAYMENT
                            isFirstItem = False
                        Else
                            cmd.Parameters.AddWithValue("?", DBNull.Value)
                            cmd.Parameters.AddWithValue("?", DBNull.Value)
                        End If

                        cmd.ExecuteNonQuery()
                    End Using

                    ' Update stock for the product
                    Using stockCmd As New OleDbCommand("UPDATE PRODUCT SET STOCKS = STOCKS - ? WHERE PRODUCT_NAME = ?", conn)
                        stockCmd.Parameters.AddWithValue("?", Convert.ToInt32(row.Cells("Qty").Value))
                        stockCmd.Parameters.AddWithValue("?", row.Cells("ProductName").Value.ToString())
                        stockCmd.ExecuteNonQuery()
                    End Using
                Next

                ' Insert a blank row for visual separation (empty row with null values)
                Using separatorCmd As New OleDbCommand("INSERT INTO [Transaction] (PRODUCT_NAME, PRICE, QTY, CUSTOMER_ID, CUSTOMER_NAME, TRANSACTION_ID, ORDER_DATE, TOTAL, PAYMENT) VALUES (?,?,?,?,?,?,?,?,?)", conn)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.Parameters.AddWithValue("?", DBNull.Value)
                    separatorCmd.ExecuteNonQuery()
                End Using

                MsgBox("Payment processed successfully!", vbInformation, "Success")
                MessageBox.Show("Print the receipt")
                Button5.Enabled = True
                Button4.Enabled = False
                Button2.Enabled = False
                Button6.Enabled = False
                TextBox2.Enabled = False
                TextBox4.Enabled = False
                TextBox10.ReadOnly = True
                NumericUpDown1.Enabled = False
                TextBox5.Enabled = False
                TextBox6.Enabled = False
                TextBox6.Text = "Walk-In"
                ComboBox1.Enabled = False
                ListBox1.Visible = False
                Button3.Text = "Close"

            Catch ex As Exception
                MsgBox("Error during payment: " & ex.Message, vbCritical, "Database Error")
            Finally
                If conn.State <> ConnectionState.Closed Then conn.Close()
            End Try
        End Using
    End Sub









    ' Adding the requested button functionality
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "New" Then
            Button1.Enabled = False
            Button2.Text = "Remove"
            Button3.Text = "Cancel"
            Button3.Enabled = True
            Button6.Enabled = True
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox10.Enabled = True
            ListBox1.Enabled = True
            ComboBox1.Enabled = True
            ComboBox1.SelectedItem = "Cash"
            NumericUpDown1.Enabled = True
            TextBox1.Clear()
            TextBox5.Clear()
            TextBox8.Clear()
            DataGridView2.Rows.Clear()
            Button4.Enabled = False
            Button5.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "Remove" Then
            ' Allow removal of the selected item from the cart
            If DataGridView2.SelectedRows.Count > 0 Then
                For Each row As DataGridViewRow In DataGridView2.SelectedRows
                    If Not row.IsNewRow Then
                        DataGridView2.Rows.Remove(row) ' Remove selected row from cart
                    End If
                Next
                CalculateTotal() ' Recalculate total after removal
            End If
        ElseIf Button2.Text = "Cancel" Then
            ' Cancel any other changes and reset UI
            DisplayNorm()
        End If
    End Sub


    Private Sub EnableRemoveMode(enable As Boolean)
        ' Enable or disable the DataGridView selection for removal
        DataGridView2.MultiSelect = enable
        For Each row As DataGridViewRow In DataGridView2.Rows
            row.Selected = False ' Deselect all rows when entering remove mode
        Next
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Reset the form to its default state


        ' Clear the DataGridView (cart) content
        DataGridView2.Rows.Clear()

        ' Reset any other form fields or values that were modified
        TextBox9.Clear() ' Clear the total amount field
        TextBox10.Clear() ' Clear the amount paid field
        TextBox11.Clear() ' Clear the change field

        ' Reset the NumericUpDown and TextBox values
        NumericUpDown1.Value = 1
        NumericUpDown1.Enabled = False ' Disable until a product is selected

        ' Clear the product and customer fields
        TextBox1.Clear() ' Clear product ID
        TextBox2.Clear() ' Clear product name
        TextBox3.Clear() ' Clear price
        TextBox12.Clear() ' Clear stock information
        TextBox5.Clear() ' Clear customer ID
        TextBox6.Clear() '
        TextBox6.Text = "Walk-In"
        TextBox8.Clear()
       

        ' Clear order ID

        ' Focus on the Product ID field to start over
        TextBox1.Focus()
        Button2.Enabled = False
        Button3.Enabled = False
        Button1.Enabled = True
        DisplayNorm()


    End Sub
   


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Form6.Show()
        Me.Hide()

    End Sub
   Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim items As New List(Of String())()

        For Each row As DataGridViewRow In DataGridView2.Rows
            If Not row.IsNewRow Then
                Dim productName As String = row.Cells("ProductName").Value.ToString()
                Dim price As String = row.Cells("Price").Value.ToString()
                Dim qty As String = row.Cells("Qty").Value.ToString()
                Dim subtotal As String = row.Cells("Subtotal").Value.ToString()

                items.Add(New String() {productName, price, qty, subtotal})
            End If
        Next

        ' Replace with actual values from your form
        Dim transactionId As String = TextBox8.Text ' Or wherever the transaction ID is stored
        Dim customerName As String = TextBox6.Text ' Or the appropriate control
        Dim tendered As String = TextBox10.Text
        Dim payment As String = ComboBox1.Text

        ' Create and show Form7
        Dim frm7 As New Form7()
        frm7.DisplayItemsFromForm3(items, transactionId, customerName, tendered, payment)
        frm7.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Form5.Show()

    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        Dim ans As Integer
        ans = MsgBox("All unfinished transactions won't be saved.", vbYesNo + vbQuestion, "Attention!")
        If ans = vbYes Then
            Form8.Show()
            Me.Close()
            Me.Hide()
        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged


        If ComboBox1.Text <> "Cash" Then
            Label17.Text = "Enter exact amount"
            Label19.Visible = True
            TextBox4.Visible = True
            TextBox4.Enabled = True
            TextBox10.ReadOnly = False

        Else
            Label17.Text = "Enter amount here"
            Label19.Visible = False
            TextBox4.Visible = False
            TextBox10.ReadOnly = False ' or your original label text
        End If
    End Sub



    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Form8.Show()
        Me.Hide()
        Me.Close()

    End Sub
End Class
