Imports System.Drawing.Printing

Public Class Form7
    Dim bmp As Bitmap

    ' Button click event to print the form
    Private Sub ButtonPrint_Click(sender As Object, e As EventArgs) Handles ButtonPrint.Click
        ' Hide the print button only for the printout
        ButtonPrint.Visible = False
        Me.Refresh() ' Ensure visibility change is rendered

        ' Capture the form into a bitmap
        bmp = New Bitmap(Me.Width, Me.Height)
        Me.DrawToBitmap(bmp, New Rectangle(0, 0, Me.Width, Me.Height))

        ' Show print preview or send to printer
        Dim printDoc As New PrintDocument()
        AddHandler printDoc.PrintPage, AddressOf Me.PrintPageHandler

        Dim dlg As New PrintPreviewDialog()
        dlg.Document = printDoc
        dlg.ShowDialog()

        ' Restore the button for the user
        ButtonPrint.Visible = True
    End Sub


    ' PrintPage event handler with scaling to fit paper
    Private Sub PrintPageHandler(sender As Object, e As PrintPageEventArgs)
        Dim pageBounds As Rectangle = e.MarginBounds

        ' Calculate scale to fit bitmap into printable area
        Dim scaleX As Single = pageBounds.Width / bmp.Width
        Dim scaleY As Single = pageBounds.Height / bmp.Height
        Dim scale As Single = Math.Min(scaleX, scaleY)

        ' New width and height with scale
        Dim newWidth As Integer = CInt(bmp.Width * scale)
        Dim newHeight As Integer = CInt(bmp.Height * scale)

        ' Center image on page
        Dim posX As Integer = pageBounds.Left + (pageBounds.Width - newWidth) \ 2
        Dim posY As Integer = pageBounds.Top + (pageBounds.Height - newHeight) \ 2

        ' Draw the scaled bitmap
        e.Graphics.DrawImage(bmp, posX, posY, newWidth, newHeight)
    End Sub

    ' Existing method to populate form
    Public Sub DisplayItemsFromForm3(items As List(Of String()), transactionId As String, customerName As String, tendered As String, payment As String)
        Label4.Text = transactionId
        Label5.Text = customerName
        Label14.Text = payment
        TextBox2.Text = tendered

        If DataGridView1.Columns.Count = 0 Then
            DataGridView1.Columns.Add("ProductName", "Product Name")
            DataGridView1.Columns.Add("Price", "Price")
            DataGridView1.Columns.Add("Qty", "Quantity")
            DataGridView1.Columns.Add("Subtotal", "Subtotal")
        End If

        DataGridView1.Rows.Clear()
        Dim total As Decimal = 0D

        For Each item As String() In items
            DataGridView1.Rows.Add(item)
            Dim subtotal As Decimal
            If Decimal.TryParse(item(3), subtotal) Then
                total += subtotal
            End If
        Next

        TextBox1.Text = total.ToString("F2")

        Dim tenderedDecimal As Decimal
        If Decimal.TryParse(tendered, tenderedDecimal) Then
            TextBox3.Text = (tenderedDecimal - total).ToString("F2")
        Else
            TextBox3.Text = "0.00"
        End If

        Label9.Text = Date.Now.ToString("MM/dd/yyyy")
    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Sales Invoice"
    End Sub
End Class
