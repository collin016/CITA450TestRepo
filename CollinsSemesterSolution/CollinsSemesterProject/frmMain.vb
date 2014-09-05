' Project Name:     frmMain
' Purpose:          To refresh the skills of programming
' Change Log:       Taylor Collins, 11/28/12

Option Explicit On
Option Strict On
Option Infer Off

Public Class frmMain
    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnRent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRent.Click
        Dim aRentForm As New frmRent
        aRentForm.Show()
    End Sub

    Private Sub btnSell_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSell.Click
        Dim aSellForm As New frmSell
        aSellForm.Show()
    End Sub
End Class
