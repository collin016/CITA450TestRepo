' Class Name:           clsRentDA
' Purpose:              A class to provide all access to the database
' Change Log:           Taylor Collins 11/28/12

Option Strict On
Option Explicit On
Option Infer Off

Imports System.Data.OleDb
Imports System.Data.Common

Public Class clsRentDA

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
            strQuery = "SELECT * FROM tblRent"

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
            dbDataAdapter.Fill(ds, "tblRent")

            Return ds
        Catch ex As Exception
            MessageBox.Show("Error occured in clsRentDA. Method: GetRecords(). Error: " & ex.Message)
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
            strQuery = "DELETE FROM tblRent WHERE BoardID = ?"

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
            MessageBox.Show("Error occured in clsPetDA. Method: DeleteRecords(). Error: " & ex.Message)
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

    Public Shared Function AddRecord(ByVal aRent As clsRent) As Integer
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
            strQuery = "INSERT INTO tblRent (BrandID, SizeID, YearID, Price, CustomerName) VALUES (?,?,?,?,?)"

            ' set up ADO components
            'Dim dbDataAdapter As New OleDbDataAdapter
            ' create command object for data adapter
            Dim dbCommand As New OleDbCommand

            ' add a parameter, BrandID
            dbCommand.Parameters.AddWithValue("@BrandID", aRent.Brand)

            ' add a parameter, SizeID
            dbCommand.Parameters.AddWithValue("@SizeID", aRent.Size)

            ' add a parameter, YearID
            dbCommand.Parameters.AddWithValue("@YearID", aRent.Year)

            ' add a parameter, Price
            dbCommand.Parameters.AddWithValue("@Price", aRent.Price)

            ' add a parameter, CustomerName
            dbCommand.Parameters.AddWithValue("@CustomerName", aRent.Customer)

            ' configure the components
            dbCommand.CommandText = strQuery
            dbCommand.Connection = dbConnection

            ' declare an integer variable to hold result
            Dim intResult As Integer
            intResult = dbCommand.ExecuteNonQuery

            Return intResult

        Catch ex As Exception
            MessageBox.Show("Error occured in clsRentDA. Method: AddRecords(). Error: " & ex.Message)
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

    Public Shared Function UpdateRecord(ByVal aRent As clsRent) As Integer
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
            strQuery = "UPDATE tblRent SET BrandID = ?, SizeID = ?, YearID = ?, Price = ?, CustomerName = ? WHERE BoardID = ?"

            ' set up ADO components
            'Dim dbDataAdapter As New OleDbDataAdapter
            ' create command object for data adapter
            Dim dbCommand As New OleDbCommand

            ' add parameters
            dbCommand.Parameters.AddWithValue("@BrandID", aRent.Brand)
            dbCommand.Parameters.AddWithValue("@SizeID", aRent.Size)
            dbCommand.Parameters.AddWithValue("@YearID", aRent.Year)
            dbCommand.Parameters.AddWithValue("@Price", aRent.Price)
            dbCommand.Parameters.AddWithValue("@CustomerName", aRent.Customer)
            dbCommand.Parameters.AddWithValue("@BoardID", aRent.BoardID)

            ' configure the components
            dbCommand.CommandText = strQuery
            dbCommand.Connection = dbConnection

            ' declare an integer variable to hold result
            Dim intResult As Integer
            intResult = dbCommand.ExecuteNonQuery

            Return intResult

        Catch ex As Exception
            MessageBox.Show("Error occured in clsRentDA. Method: UpdateRecords(). Error: " & ex.Message)
            Return -9
        End Try
    End Function
    ' Method name:      SearchRecords
    ' Purpose:          To allow the user to search a record from the database by using certain attributes
    ' Parameters:       an object - clsRent
    ' Return:           A set of records that matches the certain criteria - DataSet
    ' Change log:       Taylor Collins 11/28/12

    Public Shared Function SearchRecords(ByVal aRent As clsRent) As DataSet
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
            strQuery = "SELECT * FROM tblRent WHERE "

            ' declare a variable to determine whether or not there is already a field added
            Dim blnFieldAdded As Boolean = False

            ' declare a variable to hold the number of parameters 
            Dim intNumOfParam As Integer = 0

            ' declare a database command
            Dim dbCommand As New OleDbCommand

            '*******************************************************************************
            ' find out which fields need to be included
            ' check the ID field 1st
            If aRent.BoardID <> 0 Then
                ' keep building the strQuery
                strQuery &= "BoardID = ?"
                intNumOfParam += 1
                blnFieldAdded = True
                ' add parameter to stack
                dbCommand.Parameters.AddWithValue("@BoardID", aRent.BoardID)
            End If

            ' check the name field
            If aRent.Customer.Length <> 0 Then
                intNumOfParam += 1
                If blnFieldAdded = True Then
                    ' indicates there is a parameter added, ao put "OR" in front of this
                    strQuery &= " OR CustomerName LIKE + '%' + ? + '%'"
                Else
                    ' indicates this is the first one
                    strQuery &= " CustomerName LIKE +'%' + ? + '%'"
                End If
                ' set blnFieldAdded
                blnFieldAdded = True
                dbCommand.Parameters.AddWithValue("@CustomerName", aRent.Customer)
            End If

            ' check if there is no parameter added
            If intNumOfParam = 0 Then
                strQuery = "SELECT * FROM tblRent"
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
            dbDataAdapter.Fill(ds, "tblRent")

            ' close the db connection
            CloseDb(dbConnection)

            Return ds

        Catch ex As Exception
            MessageBox.Show("Error occured in clsRentDA. Method: SearchRecords(Object). Error: " & ex.Message)
            Return Nothing
        End Try
    End Function
End Class
