
Public NotInheritable Class DateConvType2
    Inherits StringEnumeration(Of DateConvType2)

    Public Shared ReadOnly Persian_Islamic As New DateConvType2("PI")
    Public Shared ReadOnly Persian_Julian As New DateConvType2("PJ")
    Public Shared ReadOnly Islamic_Persian As New DateConvType2("IP")
    Public Shared ReadOnly Islamic_Julian As New DateConvType2("IJ")
    Public Shared ReadOnly Julian_Persian As New DateConvType2("JP")
    Public Shared ReadOnly Julian_Islamic As New DateConvType2("JI")

    Private Sub New(ByVal StringConstant As String)
        MyBase.New(StringConstant)
    End Sub
End Class

Public NotInheritable Class DateConvType1
    Inherits StringEnumeration(Of DateConvType1)

    Public Shared ReadOnly Persian_J As New DateConvType1("P")
    Public Shared ReadOnly Islamic_J As New DateConvType1("I")
    Public Shared ReadOnly Julian_J As New DateConvType1("J")

    Private Sub New(ByVal StringConstant As String)
        MyBase.New(StringConstant)
    End Sub
End Class



' '' <summary>
' '' Base Class for String Enumerations. Fully Represents
' '' an Enumeration using Strings. Automatically casts & compares to Strings.
' '' </summary>
' '' <remarks></remarks>
Public MustInherit Class StringEnumeration(Of TStringEnumeration _
       As StringEnumeration(Of TStringEnumeration))
    Implements IStringEnumeration

    Private myString As String
    Sub New(ByVal StringConstant As String)
        myString = StringConstant
    End Sub

#Region "Properties"
    Public Class [Enum]
        Public Shared Function GetValues() As String()
            Dim myValues As New List(Of String)
            For Each myFieldInfo As System.Reflection.FieldInfo _
                                 In GetSharedFieldsInfo()
                Dim myValue As StringEnumeration(Of TStringEnumeration) = _
                  CType(myFieldInfo.GetValue(Nothing),  _
                  StringEnumeration(Of TStringEnumeration))
                'Shared Fields use a Null object
                myValues.Add(myValue)
            Next
            Return myValues.ToArray
        End Function

        Public Shared Function GetNames() As String()
            Dim myNames As New List(Of String)
            For Each myFieldInfo As System.Reflection.FieldInfo _
                     In GetSharedFieldsInfo()
                myNames.Add(myFieldInfo.Name)
            Next
            Return myNames.ToArray
        End Function

        Public Shared Function GetName(ByVal myName As  _
               StringEnumeration(Of TStringEnumeration)) As String
            Return myName
        End Function

        Public Shared Function isDefined(ByVal myName As String) As Boolean
            If GetName(myName) Is Nothing Then Return False
            Return True
        End Function

        Public Shared Function GetUnderlyingType() As Type
            Return GetType(String)
        End Function

        Friend Shared Function GetSharedFieldsInfo() _
                      As System.Reflection.FieldInfo()
            Return GetType(TStringEnumeration).GetFields
        End Function

        Friend Shared Function GetSharedFields() As  _
                      StringEnumeration(Of TStringEnumeration)()
            Dim myFields As New List(Of  _
                         StringEnumeration(Of TStringEnumeration))
            For Each myFieldInfo As System.Reflection.FieldInfo _
                                 In GetSharedFieldsInfo()
                Dim myField As StringEnumeration(Of TStringEnumeration) = _
                    CType(myFieldInfo.GetValue(Nothing),  _
                    StringEnumeration(Of TStringEnumeration))
                'Shared Fields use a Null object
                myFields.Add(myField)
            Next
            Return myFields.ToArray
        End Function
    End Class
#End Region

#Region "Cast Operators"
    'Downcast to String
    Public Shared Widening Operator CType(ByVal myStringEnumeration _
           As StringEnumeration(Of TStringEnumeration)) As String
        If myStringEnumeration Is Nothing Then Return Nothing
        Return myStringEnumeration.ToString
    End Operator

    'Upcast to StringEnumeration
    Public Shared Widening Operator CType(ByVal myString As String) As  _
                           StringEnumeration(Of TStringEnumeration)
        For Each myElement As StringEnumeration(Of TStringEnumeration) In _
                 StringEnumeration(Of TStringEnumeration).Enum.GetSharedFields
            'Found a Matching StringEnumeration - Return it
            If myElement.ToString = myString Then Return myElement
        Next
        'Did not find a Match - return NOTHING
        Return Nothing
    End Operator

    Overrides Function ToString() As String Implements IStringEnumeration.ToString
        Return myString
    End Function
#End Region

#Region "Concatenation Operators"
    Public Shared Operator &(ByVal left As StringEnumeration(Of  _
           TStringEnumeration), ByVal right As StringEnumeration(Of  _
           TStringEnumeration)) As String
        If left Is Nothing And right Is Nothing Then Return Nothing
        If left Is Nothing Then Return right.ToString
        If right Is Nothing Then Return left.ToString
        Return left.ToString & right.ToString
    End Operator

    Public Shared Operator &(ByVal left As StringEnumeration(Of  _
           TStringEnumeration), ByVal right As IStringEnumeration) As String
        If left Is Nothing And right Is Nothing Then Return Nothing
        If left Is Nothing Then Return right.ToString
        If right Is Nothing Then Return left.ToString
        Return left.ToString & right.ToString
    End Operator
#End Region

#Region "Operator Equals"

    Public Shared Operator =(ByVal left As StringEnumeration(Of  _
           TStringEnumeration), ByVal right As  _
           StringEnumeration(Of TStringEnumeration)) As Boolean
        If left Is Nothing Or right Is Nothing Then Return False
        Return left.ToString.Equals(right.ToString)
    End Operator

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf (obj) Is StringEnumeration(Of TStringEnumeration) Then
            Return CType(obj, StringEnumeration(Of  _
                   TStringEnumeration)).ToString = myString
        ElseIf TypeOf (obj) Is String Then
            Return CType(obj, String) = myString
        End If
        Return False
    End Function
#End Region

#Region "Operator Not Equals"
    Public Shared Operator <>(ByVal left As StringEnumeration(Of  _
           TStringEnumeration), ByVal right As StringEnumeration(Of  _
           TStringEnumeration)) As Boolean
        Return Not left = right
    End Operator

#End Region

End Class

'Base Interface without any Generics for StringEnumerations
Public Interface IStringEnumeration
    Function ToString() As String
End Interface