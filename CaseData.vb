Imports System.Net.Http
Imports Microsoft.VisualBasic

Public Class CaseData
    Public Property pdl_id
    Public Property offence_charge
    Public Property modus_operandi
    Public Property place_of_arrest
    Public Property date_of_arrest
    Public Property arresting_officer
    Public Property lawyer
    Public Property multipleCase As Boolean
    Public Property mugshot
    Public Property cellblock
    Public Property sentence

    Public Function getDictionary() As Dictionary(Of String, Object)
        Dim caseData As New Dictionary(Of String, Object) From {
            {"pdl_id", pdl_id},
            {"offence_charge", offence_charge},
            {"modus_operandi", modus_operandi},
            {"place_of_arrest", place_of_arrest},
            {"date_of_arrest", date_of_arrest},
            {"arresting_officer", arresting_officer},
            {"lawyer", lawyer},
            {"mugshot", mugshot},
            {"multipleCase", multipleCase},
            {"cellblock", cellblock},
            {"sentence", sentence}
            }
        Return caseData
    End Function
End Class
