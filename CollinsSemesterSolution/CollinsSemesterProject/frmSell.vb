﻿' Project Name:     frmRent
' Purpose:          To refresh the skills of programming
' Change Log:       Taylor Collins, 11/28/2012

Option Explicit On
Option Strict On
Option Infer Off

Public Class frmSell
    ' declare a global variable to hold the data from database
    Private dsSell As DataSet

    ' declare a global variable to hold search mode
    Private blnSearchMode As Boolean = False

    Private Sub frmSell_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'load combo box
        Dim dsBrand As DataSet
        dsBrand = clsBrand.GetRecords()

        ' populate combo box
        With cboBrand
            .DataSource = dsBrand.Tables("tblBrand")
            .DisplayMember = "Brand"
            .ValueMember = "BrandID"
        End With

        ' load with data in the database
        dsSell = clsSell.GetRecords()

        ' call displayRecord
        DisplayRecord(0)

    End Sub

    ' Method name:      DisplayRecord
    ' Purpose:          To display all the records retrieved from the database for this class
    ' Parameters:       Index - Integer
    ' Return:           None
    ' Change log:       Taylor Collins 10/16/12

    Public Sub DisplayRecord(ByVal aIndex As Integer)
        Try
            ' check the bound to make sure there is any record
            If dsSell.Tables("tblSell").Rows.Count = 0 Then
                'there is no record in this table
                lblCurrent.Text = "0"
                MessageBox.Show("There is no record in this table.")
                Exit Sub
            Else
                ' display the record number
                lblCurrent.Text = (aIndex + 1).ToString
                lblTotal.Text = dsSell.Tables("tblSell").Rows.Count.ToString
                ' display the data onto the form
                txtBoardID.Text = dsSell.Tables("tblSell").Rows(aIndex).Item(0).ToString
                txtSize.Text = dsSell.Tables("tblSell").Rows(aIndex).Item(2).ToString
                txtYear.Text = dsSell.Tables("tblSell").Rows(aIndex).Item(3).ToString
                txtPrice.Text = dsSell.Tables("tblSell").Rows(aIndex).Item(4).ToString

                ' set the combo box item
                cboBrand.SelectedValue = dsSell.Tables("tblSell").Rows(aIndex).Item(1).ToString
            End If
        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: DisplayRecord(Integer). Error: " & ex.Message)
        End Try
    End Sub

    ' Method name:      btnNext_Click
    ' Purpose:          To move to the next record and display it
    ' Parameters:       Sender, event argument
    ' Return:           None
    ' Change log:       Taylor Collins 11/28/12

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Try
            If CInt(lblCurrent.Text) = dsSell.Tables("tblSell").Rows.Count Then
                Exit Sub
            End If
            DisplayRecord(CInt(lblCurrent.Text))

        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: btnNext_Click. Error: " & ex.Message)
        End Try
    End Sub
    ' Method name:      btnFirst_Click
    ' Purpose:          To move to the first record and display it
    ' Parameters:       Sender, event argument
    ' Return:           None
    ' Change log:       Taylor Collins 11/28/12

    Private Sub btnFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFirst.Click
        Try
            DisplayRecord(0)
        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: btnFirst_Click. Error: " & ex.Message)
        End Try
    End Sub
    ' Method name:      btnLast_Click
    ' Purpose:          To move to the last record and display it
    ' Parameters:       Sender, event argument
    ' Return:           None
    ' Change log:       Taylor Collins 11/28/12

    Private Sub btnLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLast.Click
        Try
            DisplayRecord(dsSell.Tables("tblSell").Rows.Count - 1)
        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: btnLast_Click. Error: " & ex.Message)
        End Try
    End Sub
    ' Method name:      btnPrev_Click
    ' Purpose:          To move to the previous record and display it
    ' Parameters:       Sender, event argument
    ' Return:           None
    ' Change log:       Taylor Collins 11/28/12

    Private Sub btnPrev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrev.Click
        Try
            If CInt(lblCurrent.Text) < 2 Then
                Exit Sub
            End If
            DisplayRecord(CInt(lblCurrent.Text) - 2)

        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: btnPrev_Click. Error: " & ex.Message)
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        ' declare a pet object
        Dim aSell As clsSell
        ' get and validate user input
        Try
            If IsValidNumber(txtBoardID, "Board ID") Then
                If IsValidNumber(txtSize, "Size") Then
                    If IsValidNumber(txtYear, "Year") Then
                        If IsValidNumber(txtPrice, "Price") Then
                                ' we verified every user input and are ready to create the object
                            aSell = New clsSell(CInt(txtBoardID.Text), txtSize.Text.Trim, CInt(txtYear.Text.Trim), CInt(cboBrand.SelectedValue), CDbl(txtPrice.Text))

                                ' declare variable to hold the result from a function call
                                Dim intResult As Integer
                                ' call the function
                            intResult = clsSell.UpdateRecord(aSell)

                                'evaluate result
                                If intResult = 0 Then
                                    MessageBox.Show("Update failed!")
                                ElseIf intResult = 1 Then
                                    MessageBox.Show("Record updated!")
                                Else
                                    MessageBox.Show("Something is wrong, start debugging!")
                            End If
                        End If
                    End If
                End If
            End If
            ' refresh display
            dsSell = clsSell.GetRecords()
            DisplayRecord(0)
        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: btnUpdate_Click. Error: " & ex.Message)
        End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        ' declare a pet object
        Dim aSell As clsSell
        ' get and validate user input
        Try
            If IsValidNumber(txtBoardID, "Board ID") Then
                If IsValidNumber(txtSize, "Size") Then
                    If IsValidNumber(txtYear, "Year") Then
                        If IsValidNumber(txtPrice, "Price") Then
                                ' we verified every user input and are ready to create the object
                            aSell = New clsSell(0, txtSize.Text, CInt(txtYear.Text), CInt(cboBrand.SelectedValue), CDbl(txtPrice.Text))

                                ' declare variable to hold the result from a function call
                                Dim intResult As Integer
                                ' call the function
                            intResult = clsSell.AddRecord(aSell)

                                'evaluate result
                                If intResult = 0 Then
                                    MessageBox.Show("Add failed!")
                                ElseIf intResult = 1 Then
                                    MessageBox.Show("Record added!")
                                Else
                                    MessageBox.Show("Something is wrong, start debugging!")
                            End If
                        End If
                    End If
                End If
            End If
            ' refresh display
            dsSell = clsSell.GetRecords()
            DisplayRecord(0)
        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: btnDelete_Click. Error: " & ex.Message)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim intResponse As Integer
        Try
            intResponse = MessageBox.Show("Are you sure you want to delete this?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, _
                                          MessageBoxDefaultButton.Button2)
            If intResponse = vbNo Then
                Exit Sub
            Else
                ' prepare for delete
                ' declare a variable to hold the result
                Dim intResult As Integer

                intResult = clsSell.DeleteRecord(txtBoardID.Text.Trim)

                'evaluate result
                If intResult = 0 Then
                    MessageBox.Show("Delete failed!")
                ElseIf intResult = 1 Then
                    MessageBox.Show("Record deleted!")
                Else
                    MessageBox.Show("Something is wrong, start debugging!")
                End If
            End If

            ' refresh display
            dsSell = clsSell.GetRecords()
            DisplayRecord(0)

        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: btnDelete_Click. Error: " & ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim intID As Integer
        Dim aSell As clsSell

        Try
            ' check the search status
            If Not blnSearchMode Then
                ' change search mode status to true
                blnSearchMode = True
                ' clear fields
                Call ClearFields()
                ' change txtID to allow edit
                txtBoardID.ReadOnly = False
                'modify the text of Search button
                btnSearch.Text = "Start Search"
            Else
                ' you are already in search mode
                ' validate data
                If IsNumeric(txtBoardID.Text.Trim) Then
                    intID = CInt(txtBoardID.Text.Trim)
                End If

                'create an object
                aSell = New clsSell(intID, CStr(0), 0, 0, 0.0)

                ' call the search method
                dsSell = clsSell.SearchRecords(aSell)

                ' display the results
                DisplayRecord(0)

                ' put the form back to normal
                blnSearchMode = False
                txtBoardID.ReadOnly = True
                btnSearch.Text = "&Search"

                ' tell user there is no record found
                If dsSell.Tables("tblSell").Rows.Count = 0 Then
                    MessageBox.Show("No record has been found matching search criteria.", "Search Record")
                    DisplayRecord(0)
                End If

            End If
        Catch ex As Exception
            MessageBox.Show("Error occured in frmSell. Method: btnSearch_Click. Error: " & ex.Message)
        End Try
    End Sub

    ' Method name:      ClearFields
    ' Purpose:          To clear text boxes and set combo boxes to default
    ' Parameters:       none
    ' Return:           none
    ' Change log:       Taylor Collins 11/1/12
    Public Sub ClearFields()
        txtBoardID.Clear()
        txtSize.Clear()
        txtYear.Clear()
        cboBrand.SelectedIndex = 0
        txtPrice.Clear()
        txtBoardID.Focus()
    End Sub

    Private Sub btnHome_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnHome.Click
        Me.Close()
    End Sub
End Class