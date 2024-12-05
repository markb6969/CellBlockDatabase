Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mime.MediaTypeNames
Imports System.Reflection
Imports System.Text
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Org.BouncyCastle.Asn1

Imports iText.Kernel.Pdf
Imports iText.Layout
Imports iText.Layout.Element


Public Class ReportHomeControl
    Private Sub btnGenerateReport_Click(sender As Object, e As EventArgs) Handles btnGenerateReport.Click

        ' Assuming ComboBox is named "cmbReports"
        Dim selectedReport As String = cmbReports.SelectedItem?.ToString()

        If String.IsNullOrEmpty(selectedReport) Then
            MessageBox.Show("Please select a report to generate.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim reportContent As String = GetReportContent(selectedReport)
        If String.IsNullOrEmpty(reportContent) Then
            MessageBox.Show("Failed to generate the report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim savePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{selectedReport}.txt")

        Try
            File.WriteAllText(savePath, reportContent)
            MessageBox.Show($"Report generated successfully at {savePath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"Error saving the report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try




    End Sub

    Private Function GetReportContent(reportType As String) As String
        Select Case reportType
            Case "PDL Population Report"
                Return GeneratePDLPopulationReport()
            Case "Staff Population Report"
                Return GenerateStaffPopulationReport()
            Case "Crime Report"
                Return GenerateCrimeReport()
            Case "Criminal Case Report"
                Return GenerateCriminalCaseReport()
            Case "Medical Report"
                Return GenerateMedicalReport()
            Case "PDL Release Report"
                Return GeneratePDLReleaseReport()
            Case Else
                Return String.Empty
        End Select
    End Function






    Private Function GeneratePDLPopulationReport() As String
        Dim dt As DataTable = modDB.GetTableData("pdl") ' Assuming "pdl" is the correct table name
        Dim report As New StringBuilder("PDL Population Report (Summary)" & Environment.NewLine & Environment.NewLine)

        ' Grouping and counting occurrences of common data for each field

        ' Sex
        Dim sexCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("sex")) _
        .Select(Function(g) New With {.Sex = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(s) s.Count).ToList()

        ' Civil Status
        Dim civilStatusCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("civil_status")) _
        .Select(Function(g) New With {.Status = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(c) c.Count).ToList()

        ' Country
        Dim countryCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("country")) _
        .Select(Function(g) New With {.Country = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(c) c.Count).ToList()

        ' Municipality
        Dim municipalityCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("municipality")) _
        .Select(Function(g) New With {.Municipality = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(m) m.Count).ToList()

        ' City
        Dim districtCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("district")) _
        .Select(Function(g) New With {.District = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(c) c.Count).ToList()

        ' Region
        Dim regionCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("region")) _
        .Select(Function(g) New With {.Region = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(r) r.Count).ToList()



        ' Hair Color
        Dim hairColorCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("hair_color")) _
        .Select(Function(g) New With {.HairColor = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(h) h.Count).ToList()

        ' Eye Color
        Dim eyeColorCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("eye_color")) _
        .Select(Function(g) New With {.EyeColor = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(e) e.Count).ToList()

        ' Deformaties
        Dim deformatiesCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("deformaties")) _
        .Select(Function(g) New With {.Deformity = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(d) d.Count).ToList()

        ' Cell
        Dim cellCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("cell")) _
        .Select(Function(g) New With {.Cell = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(c) c.Count).ToList()

        ' Status
        Dim statusCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("status")) _
        .Select(Function(g) New With {.Status = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(s) s.Count).ToList()

        ' Most common data for each field
        report.AppendLine("Most Common Sex:")
        For Each sex In sexCounts
            report.AppendLine($"Sex: {sex.Sex}, Occurrences: {sex.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Civil Status:")
        For Each status In civilStatusCounts
            report.AppendLine($"Civil Status: {status.Status}, Occurrences: {status.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Country:")
        For Each country In countryCounts
            report.AppendLine($"Country: {country.Country}, Occurrences: {country.Count}")
        Next

        ' Most common data for each field
        report.AppendLine("Most Common Municipality:")
        For Each municipality In municipalityCounts
            report.AppendLine($"Municipality: {municipality.Municipality}, Occurrences: {municipality.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common District:")
        For Each district In districtCounts
            report.AppendLine($"District: {district.District}, Occurrences: {district.District}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Region:")
        For Each regionItem In regionCounts
            report.AppendLine($"Region: {regionItem.Region}, Occurrences: {regionItem.Count}")
        Next






        report.AppendLine(Environment.NewLine & "Most Common Hair Color:")
        For Each hair In hairColorCounts
            report.AppendLine($"Hair Color: {hair.HairColor}, Occurrences: {hair.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Eye Color:")
        For Each eye In eyeColorCounts
            report.AppendLine($"Eye Color: {eye.EyeColor}, Occurrences: {eye.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Deformaties:")
        For Each deformity In deformatiesCounts
            report.AppendLine($"Deformity: {deformity.Deformity}, Occurrences: {deformity.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Cell:")
        For Each cell In cellCounts
            report.AppendLine($"Cell: {cell.Cell}, Occurrences: {cell.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Status:")
        For Each status In statusCounts
            report.AppendLine($"Status: {status.Status}, Occurrences: {status.Count}")
        Next

        ' Separator for readability
        report.AppendLine(New String("-"c, 50))

        Return report.ToString()
    End Function


    Private Function GenerateStaffPopulationReport() As String
        Dim dt As DataTable = modDB.GetTableData("staffdetails") ' Assuming "staffdetails" is the correct table name
        Dim report As New StringBuilder("Staff Population Report (Summary)" & Environment.NewLine & Environment.NewLine)

        ' Grouping and counting occurrences of common data for each field

        ' Municipality
        Dim municipalityCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("municipality")) _
        .Select(Function(g) New With {.Municipality = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(m) m.Count).ToList()

        ' City
        Dim cityCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("city")) _
        .Select(Function(g) New With {.City = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(c) c.Count).ToList()

        ' Region
        Dim regionCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("Region")) _
        .Select(Function(g) New With {.Region = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(r) r.Count).ToList()

        ' Gender
        Dim genderCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("gender")) _
        .Select(Function(g) New With {.Gender = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(g) g.Count).ToList()

        ' Status
        Dim statusCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("status")) _
        .Select(Function(g) New With {.Status = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(s) s.Count).ToList()

        ' Position
        Dim positionCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("position")) _
        .Select(Function(g) New With {.Position = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(p) p.Count).ToList()

        ' Authority
        Dim authorityCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("authority")) _
        .Select(Function(g) New With {.Authority = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(a) a.Count).ToList()

        ' Most common data for each field
        report.AppendLine("Most Common Municipality:")
        For Each municipality In municipalityCounts
            report.AppendLine($"Municipality: {municipality.Municipality}, Occurrences: {municipality.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common City:")
        For Each city In cityCounts
            report.AppendLine($"City: {city.City}, Occurrences: {city.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Region:")
        For Each regionItem In regionCounts
            report.AppendLine($"Region: {regionItem.Region}, Occurrences: {regionItem.Count}")
        Next


        report.AppendLine(Environment.NewLine & "Most Common Gender:")
        For Each gender In genderCounts
            report.AppendLine($"Gender: {gender.Gender}, Occurrences: {gender.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Status:")
        For Each status In statusCounts
            report.AppendLine($"Status: {status.Status}, Occurrences: {status.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Position:")
        For Each position In positionCounts
            report.AppendLine($"Position: {position.Position}, Occurrences: {position.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Authority:")
        For Each authority In authorityCounts
            report.AppendLine($"Authority: {authority.Authority}, Occurrences: {authority.Count}")
        Next

        ' Separator for readability
        report.AppendLine(New String("-"c, 50))

        Return report.ToString()
    End Function





    Private Function GenerateCrimeReport() As String
        Dim dt As DataTable = modDB.GetTableData("crimes")
        Dim report As New StringBuilder("Crime Report" & Environment.NewLine & Environment.NewLine)

        For Each row As DataRow In dt.Rows
            report.AppendLine($"Crime ID: {row("crime_id")}")
            report.AppendLine($"PDL ID: {row("pdl_id")}")
            report.AppendLine($"Crime Committed: {row("crime_committed")}")
            report.AppendLine(New String("-"c, 50)) ' Separator for readability
        Next

        Return report.ToString()
    End Function

    Private Function GenerateCriminalCaseReport() As String
        Dim dt As DataTable = modDB.GetTableData("criminal_case")
        Dim report As New StringBuilder("Criminal Case Report (Summary)" & Environment.NewLine & Environment.NewLine)

        ' Group by cellblock and count cases
        Dim cellblockCounts = dt.AsEnumerable().
        GroupBy(Function(row) row.Field(Of String)("cellblock")).
        Select(Function(g) New With {
            .Cellblock = g.Key,
            .Count = g.Count()
        }).OrderByDescending(Function(c) c.Count)

        report.AppendLine("Cases by Cellblock:")
        For Each cellblock In cellblockCounts
            report.AppendLine($"Cellblock: {cellblock.Cellblock}, Cases: {cellblock.Count}")
        Next

        ' Most common offenses
        Dim offenseCounts = dt.AsEnumerable().
        GroupBy(Function(row) row.Field(Of String)("offence_charge")).
        Select(Function(g) New With {
            .Offense = g.Key,
            .Count = g.Count()
        }).OrderByDescending(Function(o) o.Count)

        report.AppendLine(Environment.NewLine & "Most Common Offenses:")
        For Each offense In offenseCounts
            report.AppendLine($"Offense: {offense.Offense}, Occurrences: {offense.Count}")
        Next

        ' Average sentence length (assuming sentence is numeric)
        Dim avgSentence = dt.AsEnumerable().
        Where(Function(row) Not IsDBNull(row.Field(Of Object)("sentence"))).
        Average(Function(row) Convert.ToDouble(row.Field(Of Object)("sentence")))

        report.AppendLine(Environment.NewLine & $"Average Sentence Length: {avgSentence:F2} years")

        Return report.ToString()
    End Function


    Private Function GenerateMedicalReport() As String
        Dim dt As DataTable = modDB.GetTableData("medical")
        Dim report As New StringBuilder("Medical Report (Summary)" & Environment.NewLine & Environment.NewLine)

        ' Grouping and counting occurrences of common data for each field

        ' Blood Type
        Dim bloodTypeCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("blood_type")) _
        .Select(Function(g) New With {.BloodType = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(b) b.Count).ToList()

        ' Chronic Illnesses
        Dim chronicIllnessCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("chronic_illnesses")) _
        .Select(Function(g) New With {.Illness = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(c) c.Count).ToList()

        ' Allergies
        Dim allergyCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("allergies")) _
        .Select(Function(g) New With {.Allergy = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(a) a.Count).ToList()

        ' Mental Health Status
        Dim mentalHealthCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("mental_health_status")) _
        .Select(Function(g) New With {.Status = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(m) m.Count).ToList()

        ' Psychiatric Treatment Required
        Dim psychiatricTreatmentCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("psychiatric_treatment_required")) _
        .Select(Function(g) New With {.Treatment = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(p) p.Count).ToList()

        ' Insurance Provider
        Dim insuranceProviderCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("insurance_provider")) _
        .Select(Function(g) New With {.Provider = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(i) i.Count).ToList()

        ' Inborn Conditions
        Dim inbornConditionCounts = dt.AsEnumerable() _
        .GroupBy(Function(row) row.Field(Of String)("inborn_conditions")) _
        .Select(Function(g) New With {.Condition = g.Key, .Count = g.Count()}) _
        .OrderByDescending(Function(i) i.Count).ToList()

        ' Most common data for each field
        report.AppendLine("Most Common Blood Types:")
        For Each blood In bloodTypeCounts
            report.AppendLine($"Blood Type: {blood.BloodType}, Occurrences: {blood.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Chronic Illnesses:")
        For Each illness In chronicIllnessCounts
            report.AppendLine($"Illness: {illness.Illness}, Occurrences: {illness.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Allergies:")
        For Each allergy In allergyCounts
            report.AppendLine($"Allergy: {allergy.Allergy}, Occurrences: {allergy.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Mental Health Status:")
        For Each status In mentalHealthCounts
            report.AppendLine($"Mental Health Status: {status.Status}, Occurrences: {status.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Psychiatric Treatment Required:")
        For Each treatment In psychiatricTreatmentCounts
            report.AppendLine($"Treatment Required: {treatment.Treatment}, Occurrences: {treatment.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Insurance Providers:")
        For Each provider In insuranceProviderCounts
            report.AppendLine($"Insurance Provider: {provider.Provider}, Occurrences: {provider.Count}")
        Next

        report.AppendLine(Environment.NewLine & "Most Common Inborn Conditions:")
        For Each condition In inbornConditionCounts
            report.AppendLine($"Inborn Condition: {condition.Condition}, Occurrences: {condition.Count}")
        Next

        ' Separator for readability
        report.AppendLine(New String("-"c, 50))

        Return report.ToString()
    End Function



    Private Function GeneratePDLReleaseReport() As String
        Dim dt As DataTable = modDB.GetTableData("inmatereleasedetails")
        Dim report As New StringBuilder("PDL Release Report" & Environment.NewLine & Environment.NewLine)

        For Each row As DataRow In dt.Rows
            report.AppendLine($"Release ID: {row("release_id")}")
            report.AppendLine($"PDL ID: {row("pdl_id")}")
            report.AppendLine($"Name: {row("first_name")} {row("middle_name")} {row("last_name")}")
            report.AppendLine($"Sex: {row("sex")}")
            report.AppendLine($"Release Date: {row("release_date")}")
            report.AppendLine($"Type of Release: {row("type_of_release")}")
            report.AppendLine($"Reason for Release: {row("reason_release")}")
            report.AppendLine($"Officer Name: {row("officer_name")}")
            report.AppendLine($"Officer Position: {row("officer_position")}")
            report.AppendLine(New String("-"c, 50)) ' Separator for readability
        Next

        Return report.ToString()
    End Function






    Private Sub SaveReportToFile(reportName As String, content As String)
        Dim filePath As String = $"{reportName.Replace(" ", "_")}.txt"
        File.WriteAllText(filePath, content)
    End Sub


End Class
