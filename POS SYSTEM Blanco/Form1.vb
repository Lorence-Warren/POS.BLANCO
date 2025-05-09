Imports System.Data.OleDb

Public Class Form1
    Dim conn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= POS.mdb")
    Dim dr As OleDbDataReader

    ' Login attempt logic
    Dim attemptCount As Integer = 0
    Const MaxAttempts As Integer = 3
    Dim lockoutSeconds As Integer = 10
    Dim timeLeft As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Log In"
        lblCountdown.Visible = False
        txtPass.UseSystemPasswordChar = True
        Timer1.Interval = 1000  ' Set interval to 1 second (1000 milliseconds)
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Try
            conn.Open()

            Dim cmdUser As New OleDbCommand("SELECT * FROM Account WHERE USERNAME= '" & txtUsername.Text & "'", conn)
            dr = cmdUser.ExecuteReader()

            If dr.HasRows = True Then
                dr.Close()

                Dim cmdPass As New OleDbCommand("SELECT * FROM Account WHERE USERNAME= '" & txtUsername.Text & "' AND PASSWORD= '" & txtPass.Text & "'", conn)
                dr = cmdPass.ExecuteReader()

                If dr.HasRows = True Then
                    Form8.Show()
                    Me.Hide()
                    txtUsername.Clear()
                    txtPass.Clear()
                    attemptCount = 0 ' reset attempts on success
                Else
                    attemptCount += 1

                    If attemptCount >= MaxAttempts Then
                        Dim result As DialogResult = MessageBox.Show(
                            "Too many failed attempts. Try agin later or reset your password now?",
                            "Login Locked",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning)
                        txtPass.Clear()


                        If result = DialogResult.Yes Then
                            txtUsername.Clear()
                            Form3.Show()
                            Me.Hide()
                        Else
                            MsgBox("Please wait " & lockoutSeconds & " seconds.", MsgBoxStyle.Critical, "Access Denied")
                            btnLogin.Enabled = False
                            timeLeft = lockoutSeconds
                            lblCountdown.Text = "" & timeLeft & "s"
                            lblCountdown.Visible = True
                            Timer1.Start()
                        End If
                    Else
                        MsgBox("Incorrect password! Attempt " & attemptCount & " of " & MaxAttempts, MsgBoxStyle.Critical, "Login Failed")
                    End If
                End If
            Else
                MsgBox("No record found! Sign up now.", MsgBoxStyle.Exclamation, "User Not Found")
                txtUsername.Clear()
                txtPass.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If Not IsNothing(dr) Then dr.Close()
            conn.Close()
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Countdown logic
        timeLeft -= 1
        lblCountdown.Text = "" & timeLeft & "s"

        If timeLeft <= 0 Then
            ' Stop the timer when countdown finishes
            Timer1.Stop()
            btnLogin.Enabled = True  ' Re-enable login button
            lblCountdown.Visible = False  ' Hide countdown label
            attemptCount = 0  ' Reset attempt counter
        End If
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        Form2.Show()
        Me.Hide()
        txtUsername.Clear()
        txtPass.Clear()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        txtPass.UseSystemPasswordChar = Not CheckBox1.Checked
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        txtUsername.Clear()
        txtPass.Clear()
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
End Class
