Imports Microsoft.VisualBasic

Public Class PDLData
    Public Property first_name
    Public Property last_name
    Public Property middle_name
    Public Property date_of_birth
    Public Property suffix
    Public Property sex
    Public Property civil_status
    Public Property municipality
    Public Property street
    Public Property district
    Public Property region
    Public Property zip_code
    Public Property country
    Public Property phone_num
    Public Property height
    Public Property weight
    Public Property hair_color
    Public Property eye_color
    Public Property identifying_mark
    Public Property location_of_identifying_marks
    Public Property deformaties
    Public Property cell

    Public Function GetDictionary() As Dictionary(Of String, Object)
        Dim pdlData As New Dictionary(Of String, Object) From {
            {"first_name", first_name},
            {"middle_name", middle_name},
            {"last_name", last_name},
            {"date_of_birth", date_of_birth},
            {"suffix", suffix},
            {"sex", sex},
            {"civil_status", civil_status},
            {"street", street},
            {"municipality", municipality},
            {"district", district},
            {"region", region},
            {"zip_code", zip_code},
            {"country", country},
            {"phone_num", phone_num},
            {"height", height},
            {"weight", weight},
            {"hair_color", hair_color},
            {"eye_color", eye_color},
            {"identifying_marks", identifying_mark},
            {"location_of_identifying_marks", location_of_identifying_marks},
            {"deformaties", deformaties},
            {"cell", cell}
        }
        Return pdlData
    End Function
End Class
