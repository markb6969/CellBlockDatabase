Imports Microsoft.VisualBasic

Public Class MedicalData
    Public Property pdl_id
    Public Property done_at
    Public Property doctor
    Public Property date_of_medical_examination
    Public Property chronic_illnesses
    Public Property allergies
    Public Property mental_health_status
    Public Property risk_of_self_harm
    Public Property pychiatric_treatment_required

    Public Function GetDictionary() As Dictionary(Of String, Object)
        Dim medicalData As New Dictionary(Of String, Object) From {
         {"pdl_id", pdl_id},
        {"done_at", done_at},
        {"doctor", doctor},
        {"date_of_medical_examination", date_of_medical_examination},
        {"chronic_illnesses", chronic_illnesses},
        {"allergies", allergies},
        {"mental_health_status", mental_health_status},
        {"risk_of_self_harm", risk_of_self_harm},
        {"psychiatric_treatment_required", pychiatric_treatment_required}
        }
        Return medicalData
    End Function
End Class
