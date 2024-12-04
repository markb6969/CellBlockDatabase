Imports Microsoft.VisualBasic

Public Class FamilyData
    Public Property pdl_id
    Public Property father_full_name
    Public Property mother_full_name
    Public Property father_full_address
    Public Property mother_full_address
    Public Property number

    Public Function getDictionary() As Dictionary(Of String, Object)
        Dim familyData As New Dictionary(Of String, Object) From {
        {"pdl_id", pdl_id},
        {"father_full_address", father_full_address},
        {"father_full_name", father_full_name},
        {"mother_full_address", mother_full_address},
        {"mother_full_name", mother_full_name}
        }
        Return familyData
    End Function
End Class
