' Class Name:           clsBrand
' Purpose:              A container to hold attributes and methods for a class
' Change Log:           Taylor Collins 11/28/2012

Option Strict On
Option Explicit On
Option Infer Off

Public Class clsBrand
    ' private attributes
    Private intBrandID As Integer
    Private strBrand As String

    ' property methods
    Public Property BrandID() As Integer
        Get
            Return intBrandID
        End Get
        Set(ByVal value As Integer)
            If value <> 0 Then
                intBrandID = value
            Else
                intBrandID = 0
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
            End If
        End Set
    End Property

    ' default constructor
    Public Sub New()
        ' setting default values for all attributes
        intBrandID = 0
        strBrand = ""
    End Sub

    ' paramaterized constructor
    Public Sub New(ByVal aID As Integer, ByVal aName As String)
        ' assign passed in values to attributes
        intBrandID = aID
        strBrand = aName
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
            strInfo = "Brand ID: " & BrandID() & vbCr & "Brand Name: " & Brand()
        Catch ex As Exception
            MessageBox.Show("Error occured in clsBrand. Method: ClassInfo(). Error: " & ex.Message)
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
            Return clsBrandDA.GetRecords
        Catch ex As Exception
            MessageBox.Show("Error occured in clsBrand. Method: GetRecords(). Error: " & ex.Message)
            Return Nothing
        End Try
    End Function
End Class
