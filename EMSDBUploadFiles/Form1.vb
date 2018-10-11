Imports System.Web
Imports System.Security
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class Form1
    'Dim con As New SqlConnection("Data Source=10.80.20.61,1433;Initial Catalog=EMSDB_QA;User ID=opsdev;Password=opsdev@123;")
    Dim con As New SqlConnection("Data Source=BLRPRODRTM\RTM_PROD_BLR;Initial Catalog=WorkFlowManagerDB;User ID=sa;Password=Prodrtm@123;")
    Dim cmd As SqlCommand
    Dim dt As New DataTable()
    Dim destloc As String = "\\10.80.20.251\emsdb_invpro\Automated_Download"
    Dim dtnow As String = Date.Now.AddDays(-1).ToString("MMddyyyy")
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim start1 As TimeSpan = New TimeSpan(7, 0, 0)
        Dim end1 As TimeSpan = New TimeSpan(7, 30, 0)
        Dim start2 As TimeSpan = New TimeSpan(9, 0, 0)
        Dim end2 As TimeSpan = New TimeSpan(9, 30, 0)
        Dim now As TimeSpan = DateTime.Now.TimeOfDay

        If ((now > start1) AndAlso (now < end1)) Then
            Dim result As DialogResult = MessageBox.Show("Files are downloading from VDrive. Please try after 7:30 AM", "Quit", MessageBoxButtons.OK, MessageBoxIcon.Information)
            If result = DialogResult.OK Then
                Application.Exit()
            End If
        End If

        If ((now > start2) AndAlso (now < end2)) Then
            Dim result As DialogResult = MessageBox.Show("Files are downloading from VDrive. Please try after 9:30 AM", "Quit", MessageBoxButtons.OK, MessageBoxIcon.Information)
            If result = DialogResult.OK Then
                Application.Exit()
            End If
        End If
        GetClients()
    End Sub

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        If String.IsNullOrEmpty(ddlClients.Text) OrElse String.IsNullOrWhiteSpace(ddlClients.Text) Then
            MessageBox.Show("Please select client")
            Exit Sub
        End If
        If String.IsNullOrEmpty(ddlSource.Text) OrElse String.IsNullOrWhiteSpace(ddlSource.Text) Then
            MessageBox.Show("Please select source")
            Exit Sub
        End If
        If Date.Now.DayOfWeek = DayOfWeek.Sunday Or Date.Now.DayOfWeek = DayOfWeek.Saturday Then
            Exit Sub
        End If
        Label1.Text = "Uploading please wait...."
        If Date.Now.DayOfWeek = DayOfWeek.Monday Then
            dtnow = Date.Now.AddDays(-3).ToString("MMddyyyy")
        End If

        Dim result As String = GetID()
        Dim id As Integer = Convert.ToInt32(result) + 1
        Dim foldername As String
        Dim newFolderName As String
        Dim filedest As String
        Dim _originalFiledest As String
        Dim destination As String
        Dim _originalDest As String
        foldername = "\Rename_" & ddlClients.Text.Trim & "\" & dtnow.ToString() & "\" & ddlSource.Text.Trim & "\Original"
        newFolderName = "\Rename_" & ddlClients.Text.Trim & "\" & dtnow.ToString() & "\" & ddlSource.Text.Trim
        _originalFiledest = destloc & foldername
        filedest = destloc & newFolderName
        createfolder(foldername, destloc, String.Empty)
        createfolder(newFolderName, destloc, String.Empty)
        Me.OpenFileDialog1.Filter = "All files (*.*)|*.*"
        Me.OpenFileDialog1.Multiselect = True
        Me.OpenFileDialog1.Title = "Select Files"
        Dim dr As DialogResult = Me.OpenFileDialog1.ShowDialog()
        If dr = System.Windows.Forms.DialogResult.OK Then
            For Each file1 As String In OpenFileDialog1.FileNames
                Dim file As String = Path.GetFileName(file1)
                Dim FileExtension As String = "." + file.Split(".").Last.ToLower
                _originalDest = System.IO.Path.Combine(_originalFiledest, file)
                destination = System.IO.Path.Combine(filedest, ddlClients.SelectedValue & "_" & id & FileExtension)

                FMove(file1, destination)
                FMove(file1, _originalDest)

                Try
                    Dim contentType1 As String = ContentType(FileExtension)
                    Dim fs As FileStream = New FileStream(file1, FileMode.Open, FileAccess.Read)
                    Dim br As BinaryReader = New BinaryReader(fs)
                    Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
                    br.Close()
                    fs.Close()



                    Using cmd As New SqlCommand("Insert into EMSDB_FileInfo (FI_OriginalName, FI_ReceiptDate, FI_Source, FI_ClientCode, FI_FileName, FI_ContentType, FI_Data, FI_CreatedOn, FI_Status) Values (@FI_OriginalName, @FI_ReceiptDate, @FI_Source, @FI_ClientCode, @FI_FileName, @FI_ContentType, @FI_Data, @FI_CreatedOn, @FI_Status)", con)
                        cmd.Parameters.AddWithValue("@FI_OriginalName", file)
                        cmd.Parameters.AddWithValue("@FI_ReceiptDate", Today)
                        cmd.Parameters.AddWithValue("@FI_Source", ddlSource.Text)
                        cmd.Parameters.AddWithValue("@FI_ClientCode", ddlClients.SelectedValue)
                        cmd.Parameters.AddWithValue("@FI_FileName", ddlClients.SelectedValue & "_" & id & FileExtension)
                        cmd.Parameters.AddWithValue("@FI_ContentType", contentType1)
                        cmd.Parameters.AddWithValue("@FI_Data", bytes)
                        cmd.Parameters.AddWithValue("@FI_CreatedOn", DateTime.Now)
                        cmd.Parameters.AddWithValue("@FI_Status", "1")
                        con.Open()
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End Using
                Catch ex As Exception

                End Try

                id = id + 1
            Next

            MessageBox.Show("Files uploaded successfully")
            Label1.Text = "Files uploaded successfully."
        End If
    End Sub

    Private Sub GetClients()
        Using da As New SqlDataAdapter("select * from EMSDB_Clients where C_Status=1", con)
            da.Fill(dt)
            ddlClients.DataSource = dt
            ddlClients.DisplayMember = "C_Name"
            ddlClients.ValueMember = "C_Code"
        End Using
    End Sub

    Private Sub createfolder(ByVal foldername As String, ByVal destloc As String, ByVal subfolder As String)
        If subfolder = String.Empty Then
            If Not Directory.Exists(destloc & "\" & foldername) Then
                Directory.CreateDirectory(destloc & "\" & foldername)
            End If
        Else
            If Not Directory.Exists(destloc & "\" & foldername & "\" & subfolder) Then
                Directory.CreateDirectory(destloc & "" & foldername.Trim & "\" & subfolder)
            End If
        End If
    End Sub

    Private Shared Sub FMove(ByVal source As String, ByVal destination As String)
        Dim array_length As Integer = CInt(Math.Pow(2, 19))
        Dim dataArray As Byte() = New Byte(array_length - 1) {}
        Using fsread As New FileStream(source, FileMode.Open, FileAccess.Read, FileShare.None, array_length)
            Using bwread As New BinaryReader(fsread)
                Using fswrite As New FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None, array_length)
                    Using bwwrite As New BinaryWriter(fswrite)
                        While True
                            Dim read As Integer = bwread.Read(dataArray, 0, array_length)
                            If 0 = read Then
                                Exit While
                            End If
                            bwwrite.Write(dataArray, 0, read)
                        End While
                    End Using
                End Using
            End Using
        End Using
        'File.Delete(source)
    End Sub

    Public Function GetID() As String
        Using cmd As New SqlCommand("Select TOP 1 FI_ID from EMSDB_FileInfo Order By FI_ID Desc", con)
            con.Open()
            Dim result As String = Convert.ToString(cmd.ExecuteScalar)
            con.Close()
            If Not String.IsNullOrEmpty(result) Then
                Return result
            Else
                Return "0"
            End If
        End Using
    End Function

    Function ContentType(ByVal FileExtension As String) As String
        Dim d As New Dictionary(Of String, String)
        'Images'
        d.Add(".bmp", "image/bmp")
        d.Add(".gif", "image/gif")
        d.Add(".jpeg", "image/jpeg")
        d.Add(".jpg", "image/jpeg")
        d.Add(".png", "image/png")
        d.Add(".tif", "image/tiff")
        d.Add(".tiff", "image/tiff")
        'Documents'
        d.Add(".doc", "application/msword")
        d.Add(".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        d.Add(".pdf", "application/pdf")
        'Slideshows'
        d.Add(".ppt", "application/vnd.ms-powerpoint")
        d.Add(".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        'Data'
        d.Add(".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d.Add(".xls", "application/vnd.ms-excel")
        d.Add(".csv", "text/csv")
        d.Add(".xml", "text/xml")
        d.Add(".txt", "text/plain")
        'Compressed Folders'
        d.Add(".zip", "application/zip")
        'Audio'
        d.Add(".ogg", "application/ogg")
        d.Add(".mp3", "audio/mpeg")
        d.Add(".wma", "audio/x-ms-wma")
        d.Add(".wav", "audio/x-wav")
        'Video'
        d.Add(".wmv", "audio/x-ms-wmv")
        d.Add(".swf", "application/x-shockwave-flash")
        d.Add(".avi", "video/avi")
        d.Add(".mp4", "video/mp4")
        d.Add(".mpeg", "video/mpeg")
        d.Add(".mpg", "video/mpeg")
        d.Add(".qt", "video/quicktime")
        Return d(FileExtension)
    End Function
End Class
