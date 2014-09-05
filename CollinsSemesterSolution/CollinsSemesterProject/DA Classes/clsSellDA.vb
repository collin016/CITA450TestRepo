' Class Name:           clsRentDA
' Purpose:              A class to provide all access to the database
' Change Log:           Taylor Collins 11/28/12

Option Strict On
Option Explicit On
Option Infer Off

Imports System.Data.OleDb
Imports System.Data.Common

Public Class clsSellDA

    ' Method name:      GetRecords
    ' Purpose:          To get all records from database
    ' Parameters:       None
    ' Return:           All the records for this class - DataSet
    ' Change log:       Taylor Collins 11/28/12

    Public Shared Function GetRecords() As DataSet
        ' error handling
        Try
            ' connect to the database
            Dim dbConnection As OleDbConnection
            dbConnection = ConnectToDb()

            ' check connection
            If dbConnection Is Nothing Then
                Return Nothing
            End If

            ' otherwise, the code follows will be executed
            ' declare a SQL query string
            Dim strQuery As String
            strQuery = "SELECT * FROM tblSell"

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

            ' fill the dataset with data retrieved from the database
            dbDataAdapter.Fill(ds, "tblSell")

            Return ds
        Catch ex As Exception
            MessageBox.Show("Error occured in clsSellDA. Method: GetRecords(). Error: " & ex.Message)
            Return Nothing
        End Try
    End Function
    ' Method name:      DeleteRecord
    ' Purpose:          To delete a record from database
    ' Parameters:       primary key - string
    ' Return:           Result (number of rows affected) - integer
    '                   0 - nothing gets deleted
    '                   1 - delete is successful
    '                   else - prepare to debug
    ' Change log:       Taylor Collins 11/28/12

    Public Shared Function DeleteRecord(ByVal aPrimaryKey As String) As Integer
        Try
            ' connect to the database
            Dim dbConnection As OleDbConnection
            dbConnection = ConnectToDb()

            ' check connection
            If dbConnection Is Nothing Then
                Return Nothing
            End If

            ' otherwise, the code follows will be executed
            ' declare a SQL query string
            Dim strQuery As String
            strQuery = "DELETE FROM tblSell WHERE BoardID = ?"

            ' set up ADO components
            'Dim dbDataAdapter As New OleDbDataAdapter
            ' create command object for data adapter
            Dim dbCommand As New OleDbCommand

            ' add a parameter
            dbCommand.Parameters.Add(New OleDbParameter("@BoardID", aPrimaryKey))

            ' configure the components
            dbCommand.CommandText = strQuery
            dbCommand.Connection = dbConnection

            ' declare an integer variable to hold result
            Dim intResult As Integer
            intResult = dbCommand.ExecuteNonQuery

            Return intResult

        Catch ex As Exception
            MessageBox.Show("Error occured in clsSellDA. Method: DeleteRecords(). Error: " & ex.Message)
            Return -9
        End Try
    End Function
    ' Method name:      AddRecord
    ' Purpose:          To add a record from database
    ' Parameters:       object of entity - class object
    ' Return:           Result (number of rows affected) - integer
    '                   0 - nothing gets added
    '                   1 - add is successful
    '                   else - prepare to debug
    ' Change log:       Taylor Collins 11/28/12

    Public Shared Function AddRecord(ByVal aSell As clsSell) As Integer
        Try
            ' connect to the database
            Dim dbConnection As OleDbConnection
            dbConnection = ConnectToDb()

            ' check connection
            If dbConnection Is Nothing Then
                Return Nothing
            End If

            ' otherwise, the code follows will be executed
            ' declare a SQL query string
            Dim strQuery As String
            ' primary key is an auto-number so we can ignore it
            strQuery = "INSERT INTO tblSell (BrandID, SizeID, YearID, Price) VALUES (?,?,?,?)"

            ' set up ADO components
            'Dim dbDataAdapter As New OleDbDataAdapter
            ' create command object for data adapter
            Dim dbCommand As New OleDbCommand

            ' add a parameter, BrandID
            dbCommand.Parameters.AddWithValue("@BrandID", aSell.Brand)

            ' add a parameter, SizeID
            dbCommand.Parameters.AddWithValue("@SizeID", aSell.Size)

            ' add a parameter, YearID
            dbCommand.Parameters.AddWithValue("@YearID", aSell.Year)

            ' add a parameter, Price
            dbCommand.Parameters.AddWithValue("@Price", aSell.Price)

            ' configure the components
            dbCommand.CommandText = strQuery
            dbCommand.Connection = dbConnection

            ' declare an integer variable to hold result
            Dim intResult As Integer
            intResult = dbCommand.ExecuteNonQuery

            Return intResult

        Catch ex As Exception
            MessageBox.Show("Error occured in clsSellDA. Method: AddRecords(). Error: " & ex.Message)
            Return -9
        End Try
    End Function
    ' Method name:      UpdateRecord
    ' Purpose:          To update a record in database
    ' Parameters:       object of entity - class object
    ' Return:           Result (number of rows affected) - integer
    '                   0 - nothing gets added
    '                   1 - add is successful
    '                   else - prepare to debug
    ' Change log:       Taylor Collins 10/30/12

    Public Shared Function UpdateRecord(ByVal aSell As clsSell) As Integer
        Try
            ' connect to the database
            Dim dbConnection As OleDbConnection
            dbConnection = ConnectToDb()

            ' check connection
            If dbConnection Is Nothing Then
                Return Nothing
            End If

            ' otherwise, the code follows will be executed
            ' declare a SQL query string
            Dim strQuery As String
            ' primary key is an auto-number so we can ignore it
            strQuery = "UPDATE tblSell SET BrandID = ?, SizeID = ?, YearID = ?, Price = ? WHERE BoardID = ?"

            ' set up ADO components
            'Dim dbDataAdapter As New OleDbDataAdapter
            ' create command object for data adapter
            Dim dbCommand As New OleDbCommand

            ' add parameters
            dbCommand.Parameters.AddWithValue("@BrandID", aSell.Brand)
            dbCommand.Parameters.AddWithValue("@SizeID", aSell.Size)
            dbCommand.Parameters.AddWithValue("@YearID", aSell.Year)
            dbCommand.Parameters.AddWithValue("@Price", aSell.Price)
            dbCommand.Parameters.AddWithValue("@BoardID", aSell.BoardID)

            ' configure the components
            dbCommand.CommandText = strQuery
            dbCommand.Connection = dbConnection

            ' declare an integer variable to hold result
            Dim intResult As Integer
            intResult = dbCommand.ExecuteNonQuery

            Return intResult

        Catch ex As Exception
            MessageBox.Show("Error occured in clsSellDA. Method: UpdateRecords(). Error: " & ex.Message)
            Return -9
        End Try
    End Function
    ' Method name:      SearchRecords
    ' Purpose:          To allow the user to search a record from the database by using certain attributes
    ' Parameters:       an object - clsRent
    ' Return:           A set of records that matches the certain criteria - DataSet
    ' Change log:       Taylor Collins 11/28/12

    Public Shared Function SearchRecords(ByVal aSell As clsSell) As DataSet
        Try
            ' connect to the database
            Dim dbConnection As OleDbConnection
            dbConnection = ConnectToDb()

            ' check connection
            If dbConnection Is Nothing Then
                Return Nothing
            End If

            ' otherwise, the code follows will be executed
            ' declare a SQL query string
            Dim strQuery As String
            strQuery = "SELECT * FROM tblSell WHERE "

            ' declare a variable to determine whether or not there is already a field added
            Dim blnFieldAdded As Boolean = False

            ' declare a variable to hold the number of parameters 
            Dim intNumOfParam As Integer = 0

            ' declare a database command
            Dim dbCommand As New OleDbCommand

            '*******************************************************************************
            ' find out which fields need to be included
            ' check the ID field 1st
            If aSell.BoardID <> 0 Then
                ' keep building the strQuery
                strQuery &= "BoardID = ?"
                intNumOfParam += 1
                blnFieldAdded = True
                ' add parameter to stack
                dbCommand.Parameters.AddWithValue("@BoardID", aSell.BoardID)
            End If

            ' check if there is no parameter added
            If intNumOfParam = 0 Then
                strQuery = "SELECT * FROM tblSell"
            End If

            ' set up ADO components
            Dim dbDataAdapter As New OleDbDataAdapter

            ' configure the components
            dbCommand.CommandText = strQuery
            dbCommand.Connection = dbConnection
            dbDataAdapter.SelectCommand = dbCommand

            ' declare a dataset
            Dim ds As New DataSet

            ' fill the dataset with data retrieved from the database
            dbDataAdapter.Fill(ds, "tblSell")

            ' close the db connection
            CloseDb(dbConnection)

            Return ds

        Catch ex As Exception
            MessageBox.Show("Error occured in clsSellDA. Method: SearchRecords(Object). Error: " & ex.Message)
            Return Nothing
        End Try
    End Function
End Class
