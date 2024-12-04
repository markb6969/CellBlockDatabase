Public Class UserData
    ' Properties to store data
    Public Property Id As Integer?
    Public Property First_Name As String
    Public Property Middle_Name As String
    Public Property Last_Name As String
    Public Property Date_Of_Birth As Date?
    Public Property Suffix As String
    Public Property Sex As String
    Public Property Civil_Status As String
    Public Property Street As String
    Public Property Municipality As String
    Public Property phone_num As String
    Public Property zip As String
    Public Property region As String
    Public Property city As String
    Public Property country As String
    Public Property Created_At As Date?
    Public Property Updated_At As Date?
    Public Property Email As String
    Public Property Username As String
    Public Property Password_Hash As String
    Public Property user_level As String
    Public Property Image As Byte()
    Public Property IdImage As Byte()

    ' Function to return data as a dictionary
    Public Function GetDictionary() As Dictionary(Of String, Object)
        Dim userDataDict As New Dictionary(Of String, Object) From {
            {"id", Id},
            {"first_name", First_Name},
            {"middle_name", Middle_Name},
            {"last_name", Last_Name},
            {"date_of_birth", Date_Of_Birth},
            {"suffix", Suffix},
            {"sex", Sex},
            {"civil_status", Civil_Status},
            {"street", Street},
            {"municipality", Municipality},
            {"phone_num", phone_num},
            {"zip", zip},
            {"region", region},
            {"city", city},
            {"country", country},
            {"created_at", Created_At},
            {"updated_at", Updated_At},
            {"email", Email},
            {"username", Username},
            {"password_hash", Password_Hash},
            {"user_level", user_level},
            {"Image", Image},
            {"id_Image", IdImage}
        }
        Return userDataDict
    End Function

End Class
