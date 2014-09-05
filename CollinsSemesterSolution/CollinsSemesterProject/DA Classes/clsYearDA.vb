' Class Name:           clsYearDA
' Purpose:              A container to hold attributes and methods for a class
' Change Log:           Taylor Collins 11/28/12

Option Strict On
Option Explicit On
Option Infer Off

Imports System.Data.OleDb

Public Class clsYearDA
    ' Method name:      GetRecords
    ' Purpose:          To get all records from database
    ' Parameters:       None
    ' Return:           All the records for this class - DataSet
    ' Change log:       Taylor Collins 11/28/12
    Public Shared Function GetRecords() As DataSet
        Try
            ' connect to database
            Dim dbConnection As OleDbConnection
            dbConnection = ConnectToDb()

            ' check if connection is succesfull
            If dbConnection Is Nothing Then
                Return Nothing
            End If

            ' declare a SQL string to manipulate the database
            Dim strQuery As String
            strQuery = "SELECT * FROM tblYear"

            ' set up ADO components
            Dim dbDataAdapter As New OleDbDataAdapter
            ' create command object for data adapter
            Dim dbCommand As New OleDbCommand

            ' configure the components
            dbCommand.CommandText = strQuery
            dbCommand.Connection = dbConnection
            dbDataAdapter.SelectCommand = dbCommand

            ' declare a dataset
            Dim ds As New DataSet

            ' fill the dataset with proper table
            dbDataAdapter.Fill(ds, "tblYear")

            ' return your dataset
            Return ds

        Catch ex As Exception
            MessageBox.Show("Error occured in clsYear. Method: GetRecords(). Error: " & ex.Message)
            Return Nothing
        End Try
    End Function
End Class
