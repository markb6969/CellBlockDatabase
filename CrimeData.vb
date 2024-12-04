Imports Microsoft.VisualBasic

Public Class CrimeData
    Public Property pdl_id
    Public Property crime_commited
    Public Function getDictionary() As Dictionary(Of String, Object)
        Dim crimeData As New Dictionary(Of String, Object) From {
          {"pdl_id", pdl_id},
          {"crime_committed", crime_commited}
        }
        Return crimeData
    End Function
End Class
