' Class Name:           clsBrand
' Purpose:              A container to hold attributes and methods for a class
' Change Log:           Taylor Collins 11/28/2012

Option Strict On
Option Explicit On
Option Infer Off

Public Class clsSize
    ' private attributes
    Private intSizeID As Integer
    Private intSize As Integer

    ' property methods
    Public Property SizeID() As Integer
        Get
            Return intSizeID
        End Get
        Set(ByVal value As Integer)
            If value <> 0 Then
                intSizeID = value
            Else
                intSizeID = 0
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

    ' default constructor
    Public Sub New()
        ' setting default values for all attributes
        intSizeID = 0
        intSize = 0
    End Sub

    ' paramaterized constructor
    Public Sub New(ByVal aID As Integer, ByVal aName As Integer)
        ' assign passed in values to attributes
        intSizeID = aID
        intSize = aName
    End Sub

    ' Method name:      ClassInfo
    ' Purpose:          To display all information about a class
    ' Parameters:       None
    ' Return:           All attributes - string
    ' Change log:       Taylor Collins 11/28/12

    Public Function ClassInfo() As String
        Dim strInfo As String
        'error handling
        Try
            strInfo = "Size ID: " & SizeID() & vbCr & "Size: " & Size()
        Catch ex As Exception
            MessageBox.Show("Error occured in clsSize. Method: ClassInfo(). Error: " & ex.Message)
            Return ""
        End Try
    End Function

    ' Method name:      GetRecords
    ' Purpose:          To get all records from database
    ' Parameters:       None
    ' Return:           All the records for this class - DataSet
    ' Change log:       Taylor Collins 11/28/12

    Public Shared Function GetRecords() As DataSet
        Try
            Return clsSizeDA.GetRecords
        Catch ex As Exception
            MessageBox.Show("Error occured in clsYear. Method: GetRecords(). Error: " & ex.Message)
            Return Nothing
        End Try
    End Function
End Class
