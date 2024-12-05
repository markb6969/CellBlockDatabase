Imports System.IO
Imports System.Text
Imports System.Security
Imports System.Security.Cryptography
Imports ST = System.Runtime.InteropServices
Imports MySql.Data.MySqlClient
Imports System.Data.SqlClient
Imports System.Drawing.Imaging
Module modDB
    Public accData As New UserData()
    Public inmateData As New PDLData()
    Public medData As New MedicalData()
    Public family As New FamilyData()
    Public contacts As New EmergencyContactData()
    Public crimeCase As New CaseData()
    Public crime As New CrimeData()
    Public capturedImage As Image

    'table names
    Public accounts = "accounts"
    Public cellblocks = "cellblocks"
    Public staff = "staff"
    Public crimes = "crimes"
    Public criminal_case = "criminal_case"
    Public emergency_contact = "emergency_contact"
    Public family_background = "family_background"
    Public pdl = "pdl"
    Public medical = "medical"
    Public inmatereleasedetails = "inmatereleasedetails"

    'used for updates
    Public matchedPDLId As Integer
    Public currentPDL As DataRow
    Public currentMedical As DataRow
    Public currentCase As DataRow
    Public currentFamily As DataRow
    Public currentContacts As DataRow
    Public currentStaff As DataRow
    Public currentAdmin As DataRow
    Public currentPending As DataRow
    Public currentUser As DataRow

    Public myadocon, conn As New MySqlConnection
    Public cmd As New MySqlCommand
    Public cmdRead As MySqlDataReader
    Public db_server As String = "'localhost'"
    Public db_uid As String = "'root'"
    Public db_pwd As String = "''"
    Public db_name As String = "'finaldatabase'"
    Public strConnection As String = "server=" & db_server & ";uid=" & db_uid & ";password=" & db_pwd & ";database=" & db_name & ";" & "allowuservariables='True';"
    Public selectedItem As Integer
    Public Structure LoggedUser
        Dim id As Integer
        Dim name As String
        Dim position As String
        Dim username As String
        Dim password As String
    End Structure

    Public Sub UpdateConnectionString()
        Try
            Dim config As String = System.IO.Directory.GetCurrentDirectory & "\config.txt"
            'MsgBox(config)
            Dim text As String = Nothing
            If System.IO.File.Exists(config) Then
                Using reader As System.IO.StreamReader = New System.IO.StreamReader(config)

                    text = reader.ReadToEnd
                End Using
                Dim arr_text() As String = Split(text, vbCrLf)

                strConnection = "server=" & Split(arr_text(0), "=")(1) & ";uid=" & Split(arr_text(1), "=")(1) & ";password=" & Split(arr_text(2), "=")(1) & ";database=" & Split(arr_text(3), "=")(1) & ";" & "allowuservariables='True';port=8012;"
            Else
                MsgBox("Do not exist")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public CurrentLoggedUser As LoggedUser = Nothing
    Public Sub openConn(ByVal db_name As String)
        Try
            With conn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = strConnection
                .Open()
            End With
        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub readQuery(ByVal sql As String)
        Try
            openConn(db_name)
            With cmd
                .Connection = conn
                .CommandText = sql
                cmdRead = .ExecuteReader
            End With
        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Public Function isConnectedToLocalServer() As Boolean
        Dim result As Boolean = False
        Try
            myadocon = New MySqlConnection
            myadocon.ConnectionString = strConnection
            Try
                myadocon.Open()
                If myadocon.State = ConnectionState.Open Then
                    MsgBox("is connected")
                    result = True
                Else
                    result = False
                End If
            Catch ex As Exception
                Return False
            End Try
            If myadocon.State = ConnectionState.Open Then
                myadocon.Close()
            End If
        Catch
            Return False
        End Try
        Return result
    End Function

    Function LoadToDGV(ByVal query As String, ByVal dgv As DataGridView) As Integer
        Try
            readQuery(query)
            Dim dt As DataTable = New DataTable
            dt.Load(cmdRead)
            dgv.DataSource = dt
            dgv.Refresh()
            Return dgv.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
        Return 0
    End Function

    Function LoadToDGVForDisplay(ByVal query As String, ByVal dgv As DataGridView) As Integer
        Try
            readQuery(query)
            Dim dt As DataTable = New DataTable
            dt.Load(cmdRead)
            dgv.DataSource = dt
            dgv.Refresh()
            If dgv.ColumnCount > 1 Then
                dgv.Columns(0).Visible = False
            End If
            Return dgv.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
        Return 0
    End Function

    Public Function Encrypt(ByVal clearText As String) As String

        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
             &H65, &H64, &H76, &H65, &H64, &H65,
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function
    Public Function Decrypt(ByVal cipherText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
             &H65, &H64, &H76, &H65, &H64, &H65,
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function

    Sub Logs(ByVal transaction As String, Optional ByVal events As String = "*_Click")
        Try
            readQuery(String.Format("INSERT INTO `logs`(`dt`, `user_accounts_id`, `event`, `transactions`) VALUES ({0},{1},'{2}','{3}')", "now()",
                                    CurrentLoggedUser.id,
                                    events,
                                    transaction))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub CreateRecord(ByVal tableName As String, ByVal data As Dictionary(Of String, Object))
        Try
            Dim columns As String = String.Join(", ", data.Keys)
            Dim values As String = String.Join(", ", data.Keys.Select(Function(key) $"@{key}"))
            Dim sql As String = $"INSERT INTO {tableName} ({columns}) VALUES ({values})"

            openConn(db_name)
            cmd.Connection = conn
            cmd.CommandText = sql
            cmd.Parameters.Clear()
            For Each kvp As KeyValuePair(Of String, Object) In data
                cmd.Parameters.AddWithValue($"@{kvp.Key}", kvp.Value)
            Next
            cmd.ExecuteNonQuery() ' Executes the insert query
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Public Function ReadRecords(ByVal tableName As String, ByVal conditions As Dictionary(Of String, Object)) As DataTable
        Dim dt As New DataTable
        Try
            Dim whereClause As String = String.Join(" AND ", conditions.Keys.Select(Function(key) $"{key}=@{key}"))
            Dim sql As String = $"SELECT * FROM {tableName} WHERE {whereClause}"

            openConn(db_name)
            cmd.Connection = conn
            cmd.CommandText = sql
            cmd.Parameters.Clear()
            For Each kvp As KeyValuePair(Of String, Object) In conditions
                cmd.Parameters.AddWithValue($"@{kvp.Key}", kvp.Value)
            Next

            Using reader As MySqlDataReader = cmd.ExecuteReader()
                dt.Load(reader)
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
        Return dt
    End Function

    Public Sub UpdateRecord(ByVal tableName As String, ByVal data As Dictionary(Of String, Object), ByVal conditions As Dictionary(Of String, Object))
        Try
            Dim setClause As String = String.Join(", ", data.Keys.Select(Function(key) $"{key}=@{key}"))
            Dim whereClause As String = String.Join(" AND ", conditions.Keys.Select(Function(key) $"{key}=@cond_{key}"))
            Dim sql As String = $"UPDATE {tableName} SET {setClause} WHERE {whereClause}"

            openConn(db_name)
            cmd.Connection = conn
            cmd.CommandText = sql
            cmd.Parameters.Clear()
            For Each kvp As KeyValuePair(Of String, Object) In data
                cmd.Parameters.AddWithValue($"@{kvp.Key}", kvp.Value)
            Next
            For Each kvp As KeyValuePair(Of String, Object) In conditions
                cmd.Parameters.AddWithValue($"@cond_{kvp.Key}", kvp.Value)
            Next
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub DeleteRecord(ByVal tableName As String, ByVal conditions As Dictionary(Of String, Object))
        Try
            Dim whereClause As String = String.Join(" AND ", conditions.Keys.Select(Function(key) $"{key}=@{key}"))
            Dim sql As String = $"DELETE FROM {tableName} WHERE {whereClause}"

            openConn(db_name)
            cmd.Connection = conn
            cmd.CommandText = sql
            cmd.Parameters.Clear()
            For Each kvp As KeyValuePair(Of String, Object) In conditions
                cmd.Parameters.AddWithValue($"@{kvp.Key}", kvp.Value)
            Next
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' Function to resize the image automatically if it is too large
    Private Function ResizeImage(image As Image) As Image
        ' Define maximum allowed width and height (e.g., 800x600)
        Const maxWidth As Integer = 500
        Const maxHeight As Integer = 500

        ' If the image exceeds the maximum width or height, resize it
        If image.Width > maxWidth OrElse image.Height > maxHeight Then
            Dim ratio As Double = Math.Min(CDbl(maxWidth) / image.Width, CDbl(maxHeight) / image.Height)
            Dim newWidth As Integer = CInt(image.Width * ratio)
            Dim newHeight As Integer = CInt(image.Height * ratio)

            ' Create a new image with the resized dimensions
            Dim resizedImage As New Bitmap(image, newWidth, newHeight)
            Return resizedImage
        End If
        Return image
    End Function

    ' Function to convert an Image to a byte array with automatic resizing and compression
    Public Function ImageToByteArray(image As Image) As Byte()
        ' Automatically resize the image if it's too large
        image = ResizeImage(image)

        Using ms As New MemoryStream()
            ' Save the resized image to the memory stream in PNG format (can change to JPEG if needed for better compression)
            image.Save(ms, ImageFormat.Png)
            Return ms.ToArray()
        End Using
    End Function

    ' Function to convert a Byte Array to an Image
    Public Function ByteArrayToImage(byteArray As Byte()) As Image
        Using ms As New MemoryStream(byteArray)
            ' Read the memory stream as an image
            Return Image.FromStream(ms)
        End Using
    End Function

    Function GetTableData(tableName As String, Optional filter As String = "") As DataTable
        Dim query As String = $"SELECT * FROM {tableName}"
        If Not String.IsNullOrWhiteSpace(filter) Then
            query &= $" WHERE {filter}"
        End If

        Dim dt As New DataTable()
        Try
            readQuery(query)
            dt.Load(cmdRead)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
        Return dt
    End Function

    Function GetTableOrder(tableName As String, orderBy As String) As DataTable
        Dim query As String = $"SELECT * FROM {tableName}"

        ' Append the ORDER BY clause
        If Not String.IsNullOrWhiteSpace(orderBy) Then
            query &= $" {orderBy}"
        End If

        Dim dt As New DataTable()
        Try
            readQuery(query)
            dt.Load(cmdRead)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
        Return dt
    End Function



    Public Function GetPDLCount() As Integer
        Dim count As Integer = 0
        Try
            openConn(db_name)
            cmd.Connection = conn
            cmd.CommandText = "SELECT COUNT(*) FROM pdl"
            count = Convert.ToInt32(cmd.ExecuteScalar())
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        Finally
            conn.Close()
        End Try
        Return count
    End Function






    Public Sub LoadPDLInformation(pdlId As Integer)
        Dim pdlData As DataTable = GetTableData(pdl, $"pdl_id = {pdlId}")
        If pdlData.Rows.Count > 0 Then
            currentPDL = pdlData.Rows(0)
        End If

        Dim familyData As DataTable = GetTableData(family_background, $"pdl_id = {pdlId}")
        If familyData.Rows.Count > 0 Then
            currentFamily = familyData.Rows(0)
        End If

        Dim contactData As DataTable = GetTableData(emergency_contact, $"pdl_id = {pdlId}")
        If contactData.Rows.Count > 0 Then
            currentContacts = contactData.Rows(0)
        End If
    End Sub


    ' Load Criminal Case Information from the database and store it in the currentCase variable
    Public Sub LoadCriminalCaseInformation(pdlId As Integer)
        Dim criminalCaseData As DataTable = GetTableData(criminal_case, $"pdl_id = {pdlId}")
        If criminalCaseData.Rows.Count > 0 Then
            currentCase = criminalCaseData.Rows(0)
        End If
    End Sub


    ' Load Medical Information from the database and store it in the currentMedical variable
    Public Sub LoadMedicalInformation(pdlId As Integer)
        Dim medicalData As DataTable = GetTableData(medical, $"pdl_id = {pdlId}")
        If medicalData.Rows.Count > 0 Then
            currentMedical = medicalData.Rows(0)
        End If
    End Sub

    Public Sub ClearAllTextBoxes(parent As Control)
        For Each ctrl As Control In parent.Controls
            If TypeOf ctrl Is TextBox Then
                DirectCast(ctrl, TextBox).Clear()
            ElseIf ctrl.HasChildren Then
                ' Recursively clear textboxes in child containers (e.g., panels, group boxes)
                ClearAllTextBoxes(ctrl)
            End If
        Next
    End Sub





End Module
