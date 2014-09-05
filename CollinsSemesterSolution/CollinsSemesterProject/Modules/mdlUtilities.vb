' Module name:      mdlUtilities
' Purpose:          A container to hold all comonly used methods
' Change log:       Taylor Collins 10/22/2012

Option Explicit On
Option Strict On
Option Infer Off

Module mdlUtilities
    ' Method name:      IsPresent
    ' Purpose:          To check if user input is there
    ' Parameters:       control - textbox
    '                   string - name of the control
    ' Return:           Boolean - T/F
    ' Change log:       Taylor Collins 9/4/2012

    Public Function IsPresent(ByVal textBox As Control, ByVal name As String) As Boolean
        If (textBox.Text = "") Then
            MessageBox.Show(name & " is a require field.", "Null Input")
            Return False
        Else
            Return True
        End If
    End Function

    ' Method name:      IsValidNumber
    ' Purpose:          To check if user input is numeric
    ' Parameters:       control - textbox
    '                   string - name of the control
    ' Return:           Boolean - T/F
    ' Change log:       Taylor Collins 9/4/2012

    Public Function IsValidNumber(ByVal textBox As Control, ByVal name As String) As Boolean
        If IsPresent(textBox, name) Then
            If IsNumeric(textBox.Text.Trim) Then
                Return True
            Else
                MessageBox.Show("Please enter numbers only for " & name, "Incorrect Input")
                Return False
            End If
        End If
        Return False
    End Function

    ' Method name:      IsValidText
    ' Purpose:          To check if user input is there
    ' Parameters:       control - textbox
    '                   string - name of the control
    ' Return:           Boolean - T/F
    ' Change log:       Taylor Collins 9/4/2012

    Public Function IsValidText(ByVal textBox As Control, ByVal name As String) As Boolean
        If IsPresent(textBox, name) Then
            If Not IsNumeric(textBox.Text.Trim) Then
                Return True
            Else
                MessageBox.Show("Please enter no numbers for " & name, "Incorrect Input")
                Return False
            End If
        End If
        Return False
    End Function

End Module

