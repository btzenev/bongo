Namespace Data

    Public Module Common

#Region " Enums "

        Public Enum QueryType
            [Select]
            [Insert]
            [Update]
            [Delete]
            [Drop]
            [Create]
            [Count]
        End Enum

        Public Enum Comparison
            [Equals]
            [NotEqual]
            [Like]
            [LessThan]
            [LessThanEqual]
            [GreaterThan]
            [GreaterThankEqual]
        End Enum

        Public Enum Logic
            [And]
            [Or]
        End Enum

        Public Enum Sort
            [Asc]
            [Desc]
        End Enum

        Public Enum DataType
            [Identity]
            [Integer]
            [Long]
            [Byte]
            [String]
            [DateTime]
        End Enum

#End Region

        Public ReadOnly Property Current() As Data.Engine
            Get
                On Error Resume Next
                Dim obj_data As Data.Engine = ToObj(Of Data.Engine)(HttpContext("cms.data"))
                If obj_data Is Nothing Then
                    obj_data = New Data.Engine("CMS")
                    HttpContext("cms.data") = obj_data
                End If
                Return obj_data
            End Get
        End Property

    End Module

End Namespace
