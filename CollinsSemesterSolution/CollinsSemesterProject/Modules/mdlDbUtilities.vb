' Module name:      mdlDbUtilities
' Purpose:          A container to hold all database manipulation methods
' Change log:       Taylor Collins 10/22/2012

Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data.OleDb

Module mdlDbUtilities
    ' Method name:      ConnectToDb
    ' Purpose:          To make connection to database
    ' Parameters:       None
    ' Return:           Database connection - OleDbConnection Object (OleDb: Object Linking Embedded Database)
    ' Change log:       Taylor Collins 9/27/2012

    Public Function ConnectToDb() As OleDbConnection
        ' error handling
        Try
            ' declare and create a database connection object
            Dim dbConnection As New OleDbConnection

            ' try to connect to database
            ' first configure the connection object
            Dim strConnection As String
            strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=../../Database/Project.accdb"
            ' assign new connection string to object property
            dbConnection.ConnectionString = strConnection

            ' open database connection
            dbConnection.Open()

            ' return the connection
            Return dbConnection

        Catch ex As Exception
            MessageBox.Show("Error occured in mdlDbUtilities. Method: ConnectToDb. Error: " & ex.Message)
            Return Nothing
        End Try
    End Function

    ' Method name:      CloseDb
    ' Purpose:          To disconnect from database
    ' Parameters:       Database connection - OleDbConnection Object (OleDb: Object Linking Embedded Database)
    ' Return:           None
    ' Change log:       Taylor Collins 9/27/2012

    Public Sub CloseDb(ByVal aConnection As OleDbConnection)
        ' error handling
        Try
            aConnection.Close()
        Catch ex As Exception
            MessageBox.Show("Error occured in mdlDbUtilities. Method: CloseDb(object). Error: " & ex.Message)
        End Try
    End Sub

End Module
