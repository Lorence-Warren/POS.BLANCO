Public Class Form8



    Private Sub POSSYSTEMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles POSSYSTEMToolStripMenuItem.Click
        Form3.Show()
        Me.Hide()

    End Sub

    Private Sub ProductInventoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductInventoryToolStripMenuItem.Click
        Form4.Show()
        Me.Hide()

    End Sub

    Private Sub CustomerListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CustomerListToolStripMenuItem.Click
        Form5.Show()
        Me.Hide()

    End Sub

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DASHBOARDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DASHBOARDToolStripMenuItem.Click
        Form6.Show()
            Me.Close()
            Me.Hide()

    End Sub

    Private Sub SIGNUPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SIGNUPToolStripMenuItem.Click
        Dim ans As Integer
        ans = MsgBox("You are about to be logged out?", vbYesNo + vbQuestion, "Attention!")
        If ans = vbYes Then
            Form2.Show()
            Me.Hide()
            Me.Close()

        End If

    End Sub

    Private Sub LOGOUTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LOGOUTToolStripMenuItem.Click
        Dim ans As Integer
        ans = MsgBox("You are about to be logged out?", vbYesNo + vbQuestion, "Attention!")
        If ans = vbYes Then
            Form1.Show()
            Me.Close()
            Me.Hide()


        End If
    End Sub
End Class