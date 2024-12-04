Imports Microsoft.VisualBasic

Public Class EmergencyContactData
    Public Property pdl_id
    Public Property contact_full_name
    Public Property contact_full_address
    Public Property relationship_with_pdl
    Public Property number

    Public Function getDictionary() As Dictionary(Of String, Object)
        Dim contactData As New Dictionary(Of String, Object) From {
            {"pdl_id", pdl_id},
            {"contact_full_address", contact_full_address},
            {"contact_full_name", contact_full_name},
            {"relationship_with_pdl", relationship_with_pdl},
            {"contact_number", number}
        }
        Return contactData
    End Function
End Class
