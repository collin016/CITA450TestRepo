' Class Name:           clsSell
' Purpose:              A container to hold attributes and methods for a class
' Change Log:           Taylor Collins 11/19/12

Option Strict On
Option Explicit On
Option Infer Off
Public Class clsSell
#Region "Class Definition"
    ' private attributes
    Private intBoardID As Integer
    Private strBrand As String
    Private intSize As Integer
    Private intYear As Integer
    Private dblPrice As Double

    ' property method for each attribute
    Public Property BoardID() As Integer
        Get
            Return intBoardID
        End Get
        Set(ByVal value As Integer)
            If value <> 0 Then
                intBoardID = value
            Else
                intBoardID = 0
            End If
        End Set
    End Property

    Public Property Brand() As String
        Get
            Return strBrand
        End Get
        Set(ByVal value As String)
            If value <> "" Then
                strBrand = value
            Else
                strBrand = "Null String"
            End If
        End Set
    End Property

    Public Property Size() As Integer
        Get
            Return intSize
        End Get
        Set(ByVal value As Integer)
            If value <> 0 Then
                intSize = value
            Else
                intSize = 0
            End If
        End Set
    End Property

    Public Property Year() As Integer
        Get
            Return intYear
        End Get
        Set(ByVal value As Integer)
            If value <> 0 Then
                intYear = value
            Else
                intYear = 0
            End If
        End Set
    End Property

    Public Property Price() As Double
        Get
            Return dblPrice
        End Get
        Set(ByVal value As Double)
            If value > 0 Then
                dblPrice = value
            Else
                dblPrice = 0
            End If
        End Set
    End Property

    ' default constructor
    Public Sub New()
        ' setting default values for all attributes
        intBoardID = 0
        strBrand = ""
        intSize = 0
        intYear = 0
        dblPrice = 0.0
    End Sub

    ' paramaterized constructor
    Public Sub New(ByVal aID As Integer, ByVal aName As String, ByVal aSize As Integer, ByVal aYear As Integer, ByVal aPrice As Double)
        ' assign passed in values to attributes
        intBoardID = aID
        strBrand = aName
        intSize = aSize
        intYear = aYear
        dblPrice = aPrice
    End Sub

    ' Method name:      ClassInfo
    ' Purpose:          To display all information about a class
    ' Parameters:       None
    ' Return:           All attributes - string
    ' Change log:       Taylor Collins 11/19/2012

    Public Function ClassInfo() As String
        Dim strInfo As String
        'error handling
        Try
            strInfo = "Board ID: " & BoardID() & vbCr _
                & "Brand: " & Brand() & vbCr _
                & "Size: " & Size() & vbCr _
                & "Year: " & Year() & vbCr _
                & "Price: " & Price() & vbCr
            Return strInfo
        Catch ex As Exception
            MessageBox.Show("Error occured in clsSell. Method: ClassInfo(). Error: " & ex.Message)
            Return ""
        End Try
    End Function
#End Region

    ' Method name:      GetRecords
    ' Purpose:          To get all records from database
    ' Parameters:       None
    ' Return:           All the records for this class - DataSet
    ' Change log:       Taylor Collins 11/19/12

    Public Shared Function GetRecords() As DataSet
        Try
            Return clsSellDA.GetRecords
        Catch ex As Exception
            MessageBox.Show("Error occured in clsSell. Method: GetRecords(). Error: " & ex.Message)
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
    ' Change log:       Taylor Collins 11/19/12

    Public Shared Function DeleteRecord(ByVal aPrimaryKey As String) As Integer
        Try
            Return clsSellDA.DeleteRecord(aPrimaryKey)
        Catch ex As Exception
            MessageBox.Show("Error occured in clsSell. Method: DeleteRecords(). Error: " & ex.Message)
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
    ' Change log:       Taylor Collins 11/19/12

    Public Shared Function AddRecord(ByVal aSell As clsSell) As Integer
        Try
            Return clsSellDA.AddRecord(aSell)
        Catch ex As Exception
            MessageBox.Show("Error occured in clsSell. Method: AddRecords(). Error: " & ex.Message)
            Return -9
        End Try
    End Function
    ' Method name:      UpdateRecord
    ' Purpose:          To update a record in database
    ' Parameters:       object of entity - class object
    ' Return:           Result (number of rows affected) - integer
    '                   0 - nothing gets updated
    '                   1 - update is successful
    '                   else - prepare to debug
    ' Change log:       Taylor Collins 11/19/12

    Public Shared Function UpdateRecord(ByVal aSell As clsSell) As Integer
        Try
            Return clsSellDA.UpdateRecord(aSell)
        Catch ex As Exception
            MessageBox.Show("Error occured in clsSell. Method: UpdateRecord(). Error: " & ex.Message)
            Return -9
        End Try
    End Function
    ' Method name:      SearchRecords
    ' Purpose:          To allow the user to search a record from the database by using certain attributes
    ' Parameters:       an object - clsPet
    ' Return:           A set of records that matches the certain criteria - DataSet
    ' Change log:       Taylor Collins 11/19/12

    Public Shared Function SearchRecords(ByVal aSell As clsSell) As DataSet
        Try
            Return clsSellDA.SearchRecords(aSell)
        Catch ex As Exception
            MessageBox.Show("Error occured in clsSell. Method: SearchRecords(). Error: " & ex.Message)
            Return Nothing
        End Try
    End Function
End Class
