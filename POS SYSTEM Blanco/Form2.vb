Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Public Class Form2
    Dim connectionstring As String
    Dim dbconnection As OleDbConnection
    Dim dbadapter As New OleDbDataAdapter
    Dim dbdataset As New DataSet

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Sign Up"
        DisplayNorm()
    End Sub

    Private Sub DisplayNorm()

        connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= POS.mdb"
        dbconnection = New OleDbConnection(connectionstring)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox2.Text.Trim() <> vbNullString And TextBox3.Text.Trim() <> vbNullString Then

                Dim dob As Date
                Dim dobInput As String = TextBox3.Text.Trim()
                If Not DateTime.TryParseExact(dobInput, "MM/dd/yyyy", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dob) Then
                    MsgBox("Invalid date format! Please use mm/dd/yyyy", vbExclamation, "Invalid Date")
                    Exit Sub
                End If

                If dbconnection.State <> ConnectionState.Open Then
                    dbconnection.Open()
                End If

                Dim checkCommand As New OleDbCommand("SELECT COUNT(*) FROM Account WHERE [USERNAME] = @username", dbconnection)
                checkCommand.Parameters.AddWithValue("@username", TextBox1.Text.Trim())

                Dim userCount As Integer = Convert.ToInt32(checkCommand.ExecuteScalar())

                If userCount > 0 Then
                    MsgBox("Username already taken. Please choose a different one.", vbExclamation, "Duplicate Username")
                    dbconnection.Close()
                    Exit Sub
                End If

                Dim dbcommand As New OleDbCommand("INSERT INTO Account ([USERNAME], [PASSWORD], [DOB]) VALUES (@username, @password, @dob)", dbconnection)

                dbcommand.Parameters.AddWithValue("@username", TextBox1.Text.Trim())
                dbcommand.Parameters.AddWithValue("@password", TextBox2.Text.Trim())
                dbcommand.Parameters.AddWithValue("@dob", dob)

                Dim rowsAffected As Integer = dbcommand.ExecuteNonQuery()
                dbcommand.Dispose()

                dbconnection.Close()

                If rowsAffected > 0 Then
                    DisplayNorm()
                    MsgBox("Credentials Recorded Successfully!", vbInformation, "Add Record")
                    TextBox1.Clear()
                    TextBox2.Clear()
                    TextBox3.Clear()
                    MsgBox("Back to Log in")
                    Form1.Show()
                    Me.Hide()

                Else
                    MsgBox("Failed to record credentials. Please try again.", vbExclamation, "Error")
                End If

            ElseIf TextBox1.Text.Trim() = vbNullString Then
                MsgBox("Enter Username!", vbCritical, "Missing")
            ElseIf TextBox2.Text.Trim() = vbNullString Then
                MsgBox("Enter Password!", vbCritical, "Missing")
            ElseIf TextBox3.Text.Trim() = vbNullString Then
                MsgBox("Enter Birthdate!", vbCritical, "Missing")
            End If

        Catch ex As OleDbException
            MsgBox("Database Error: " & ex.Message, vbCritical, "Error")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbCritical, "Error")
        Finally
            If dbconnection.State = ConnectionState.Open Then
                dbconnection.Close()
            End If
        End Try
    End Sub


    


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then

            TextBox2.UseSystemPasswordChar = False
        Else
            TextBox2.UseSystemPasswordChar = True

        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.Show()
        Me.Hide()
        Me.Close()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
    End Sub
End Class


