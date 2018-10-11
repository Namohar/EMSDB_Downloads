Imports System.Web
Imports System.Security
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class Form1
    'production database connection
    'Dim con As New SqlConnection("Data Source=BLRPRODRTM\RTM_PROD_BLR;Initial Catalog=Real_Time_Metrics_Dev;User ID=sa;Password=Prodrtm@123;")
    Dim con As New SqlConnection("Data Source=BLRPRODRTM\RTM_PROD_BLR;Initial Catalog=WorkFlowManagerDB;User ID=sa;Password=Prodrtm@123;")
    'Testing database connection
    ' Dim con As New SqlConnection("Data Source=10.80.20.61,1433;Initial Catalog=EMSDB_Dev;User ID=opsdev;Password=opsdev@123;")
    Dim cmd As SqlCommand
    Dim dirList As New List(Of String)
    Dim dt As New DataTable()
    'Vdrive source location changed on 29th june 2017.
    'Dim sourceloc As String = "\\vmdalapp01.symphonycmg.com\va\Renaming"
    Dim sourceloc As String = "\\10.0.35.16\va\Renaming"
    'Dim sourceloc As String = "\\10.0.35.16\VA"

    'Sdrive source location
    Dim Sdrive_sourceloc = "\\isgfs1\pdf_repository\HP ALU"
    'Invoice File's auto-downloaded path. 
    Dim destloc As String = "\\10.80.20.251\emsdb_invpro\Automated_Download"
    'dtnow is one day lesser then current date.
    Dim dtnow As String = Date.Now.AddDays(-1).ToString("MMddyyyy")
    Dim flagDownload As Integer = 0
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'folder creationn method for auto downloaded file's in our local repository, (\\10.80.20.251\emsdb_invpro)
        createfolder("Automated_Download", "\\10.80.20.251\emsdb_invpro", "")
        'while release please comment thse below two methods.
        Download_files_New()

        Download_S_Drive_files()

        '*************bgLongRunningProcess method contains the auto download schedule's every day. 
        bgLongRunningProcess.RunWorkerAsync()
    End Sub

    Private Sub Download_files()
        Try
            Label1.Text = Date.Now
            If Date.Now.DayOfWeek = DayOfWeek.Sunday Or Date.Now.DayOfWeek = DayOfWeek.Saturday Then
                Exit Sub
            End If

            If Date.Now.DayOfWeek = DayOfWeek.Monday Then
                dtnow = Date.Now.AddDays(-3).ToString("MMddyyyy")
            End If

            Dim result As String = GetID()
            Dim id As Integer = Convert.ToInt32(result) + 1
            dt = New DataTable()
            GetClients()

            For Each dr As DataRow In dt.Rows
                Dim foldername As String
                Dim _originalFoldername As String
                Dim filesource As String
                Dim filedest As String
                Dim _originalFiledest As String
                Dim destination As String
                Dim _originalDest As String

                foldername = "\Rename_" & dr("C_Name").ToString.Trim
                _originalFoldername = foldername & "\" & dtnow.ToString() & "\VDrive" & "\Original"

                filesource = sourceloc & foldername.Trim & "\" & dtnow.ToString
                filedest = destloc & foldername
                _originalFiledest = destloc & _originalFoldername
                createfolder(foldername, destloc, dtnow.ToString)
                createfolder(_originalFoldername, destloc, String.Empty)

                If CheckCreationDate(filedest) = False Then
                    If Not Directory.Exists(filesource) Then
                    Else

                        dirList = GetAllFiles(filesource)
                        If (dirList.Count > 0) Then
                            For Each filename As String In dirList
                                Dim file As String = Path.GetFileName(filename)
                                Dim lastFolderName As String = Path.GetFileName(Path.GetDirectoryName(filename))
                                _originalFiledest = String.Empty
                                _originalFiledest = destloc & foldername & "\" & dtnow.ToString() & "\VDrive" & "\Original"
                                _originalFiledest = _originalFiledest & "\" & lastFolderName
                                createfolder(_originalFoldername & "\" & lastFolderName, destloc, String.Empty)
                                _originalDest = System.IO.Path.Combine(_originalFiledest, file)
                                If System.IO.File.Exists(_originalDest) = False Then
                                    Dim FileExtension As String = "." + file.Split(".").Last.ToLower
                                    filedest = String.Empty
                                    filedest = destloc & foldername & "\" & dtnow.ToString & "\VDrive"
                                    filedest = filedest & "\" & lastFolderName
                                    createfolder(foldername, destloc, dtnow.ToString & "\VDrive" & "\" & lastFolderName)
                                    destination = System.IO.Path.Combine(filedest, dr("C_Code").ToString.Trim & "_" & id & FileExtension)

                                    'System.IO.File.Copy(filename, destination, True)
                                    FMove(filename, destination)
                                    'System.IO.File.Copy(filename, _originalDest, True)
                                    FMove(filename, _originalDest)

                                    Try
                                        Dim contentType1 As String = ContentType(FileExtension)
                                        Dim fs As FileStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                                        Dim br As BinaryReader = New BinaryReader(fs)
                                        Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
                                        br.Close()
                                        fs.Close()



                                        Using cmd As New SqlCommand("Insert into EMSDB_FileInfo (FI_OriginalName, FI_ReceiptDate, FI_Source, FI_ClientCode, FI_FileName, FI_ContentType, FI_Data, FI_CreatedOn, FI_Status) Values (@FI_OriginalName, @FI_ReceiptDate, @FI_Source, @FI_ClientCode, @FI_FileName, @FI_ContentType, @FI_Data, @FI_CreatedOn, @FI_Status)", con)
                                            cmd.Parameters.AddWithValue("@FI_OriginalName", file)
                                            cmd.Parameters.AddWithValue("@FI_ReceiptDate", Today)
                                            cmd.Parameters.AddWithValue("@FI_Source", "VDrive")
                                            cmd.Parameters.AddWithValue("@FI_ClientCode", dr("C_Code").ToString())
                                            cmd.Parameters.AddWithValue("@FI_FileName", dr("C_Code").ToString.Trim & "_" & id & FileExtension)
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

                                End If
                            Next
                        End If
                    End If
                End If
            Next
            Label2.Text = Date.Now
        Catch ex As Exception

        End Try

    End Sub

    'Download_files_New method downloads the Vdrive file's
    Private Sub Download_files_New()
        Try
            'DayOfWeek is name of the day ex: sunday, monday...
            If Date.Now.DayOfWeek = DayOfWeek.Sunday Or Date.Now.DayOfWeek = DayOfWeek.Saturday Then
                Exit Sub
            End If
            If Date.Now.DayOfWeek = DayOfWeek.Monday Then
                'dtnow object is if day is monday then we are subtracting from current date to less then 3 days for taking friday data. 
                dtnow = Date.Now.AddDays(-3).ToString("MMddyyyy")
            Else
                'dtnow object is if day is other then monday then we are subtracting from current date to less then 1 day for taking previous day data. 
                dtnow = Date.Now.AddDays(-1).ToString("MMddyyyy")
            End If
            'GetID() method will get maximum FI_ID in EMSDB_FileInfo table 
            Dim result As String = GetID()
            Dim id As Integer = Convert.ToInt32(result) + 1
            dt = New DataTable()
            'get active clients in emsdb_clients table.
            GetClients()
            'copy clients data into data table (dt).
            For Each dr As DataRow In dt.Rows
                'declaring empty variables.
                Dim foldername As String = String.Empty
                Dim _originalFoldername As String = String.Empty
                Dim filesource As String = String.Empty
                Dim filedest As String = String.Empty
                Dim _originalFiledest As String = String.Empty
                Dim destination As String = String.Empty
                Dim _originalDest As String = String.Empty


                foldername = "\Rename_" & dr("C_Name").ToString.Trim
                '_originalFoldername variable contains\Rename\client_name\current date\Vdrive\Original
                _originalFoldername = foldername & "\" & dtnow.ToString() & "\VDrive" & "\Original"
                'filesource contains Vdrive path with client name and date
                filesource = sourceloc & foldername.Trim & "\" & dtnow.ToString
                filedest = destloc & foldername
                '_originalFiledest contains \\10.80.20.251 auto downloads repositary path.
                _originalFiledest = destloc & _originalFoldername
                'create original folder \\10.80.20.251 auto downloads repositary path.
                createfolder(foldername, destloc, dtnow.ToString)
                'create folder \\10.80.20.251 auto downloads repositary path.
                createfolder(_originalFoldername, destloc, String.Empty)
                'check if folder already exist or not. if it's exist don't create again.
                If CheckCreationDate(filedest) = False Then
                    If Not Directory.Exists(filesource) Then
                    Else

                        Dim direc
                        'get all client folder and file's in Vdrive 
                        For Each direc In Directory.GetDirectories(filesource, "*", SearchOption.AllDirectories)
                            For Each filename As String In Directory.GetFiles(direc)
                                Dim file As String = Path.GetFileName(filename)
                                Dim lastFolderName As String = Path.GetFileName(Path.GetDirectoryName(filename))

                                _originalFiledest = String.Empty
                                _originalFiledest = destloc & foldername & "\" & dtnow.ToString() & "\VDrive" & "\Original"
                                _originalFiledest = _originalFiledest & "\" & lastFolderName
                                createfolder(_originalFoldername & "\" & lastFolderName, destloc, String.Empty)
                                _originalDest = System.IO.Path.Combine(_originalFiledest, file)
                                'checking in auto download path if folder exist it will skip that folder.
                                If System.IO.File.Exists(_originalDest) = False Then
                                    Dim FileExtension As String = "." + file.Split(".").Last.ToLower
                                    If FileExtension <> ".db" Then
                                        filedest = String.Empty
                                        filedest = destloc & foldername & "\" & dtnow.ToString & "\VDrive"
                                        filedest = filedest & "\" & lastFolderName
                                        createfolder(foldername, destloc, dtnow.ToString & "\VDrive" & "\" & lastFolderName)
                                        'getting each client name for create folder with client name.
                                        destination = System.IO.Path.Combine(filedest, dr("C_Code").ToString.Trim & "_" & id & FileExtension)

                                        'copy file's to Auto downloaded path with folder name
                                        FMove(filename, destination)
                                        'copy file's to Auto downloaded path within original folder name
                                        FMove(filename, _originalDest)

                                        Try
                                            Dim contentType1 As String = ContentType(FileExtension)
                                            Dim fs As FileStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                                            Dim br As BinaryReader = New BinaryReader(fs)
                                            'invoice's, all file formates converting into bytes 
                                            Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
                                            br.Close()
                                            fs.Close()
                                            'file_info table  insert query. data converting and inserting into Database
                                            Using cmd As New SqlCommand("Insert into EMSDB_FileInfo (FI_OriginalName, FI_ReceiptDate, FI_Source, FI_ClientCode, FI_FileName, FI_ContentType, FI_Data, FI_CreatedOn, FI_Status) Values (@FI_OriginalName, @FI_ReceiptDate, @FI_Source, @FI_ClientCode, @FI_FileName, @FI_ContentType, @FI_Data, @FI_CreatedOn, @FI_Status)", con)
                                                cmd.Parameters.AddWithValue("@FI_OriginalName", file)
                                                cmd.Parameters.AddWithValue("@FI_ReceiptDate", Today)
                                                cmd.Parameters.AddWithValue("@FI_Source", "VDrive")
                                                cmd.Parameters.AddWithValue("@FI_ClientCode", dr("C_Code").ToString())
                                                cmd.Parameters.AddWithValue("@FI_FileName", dr("C_Code").ToString.Trim & "_" & id & FileExtension)
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
                                        'auto increment values based on FI_ID
                                        id = id + 1
                                    End If
                                End If
                            Next
                        Next

                        For Each filename As String In Directory.GetFiles(filesource)
                            Dim file As String = Path.GetFileName(filename)
                            Dim lastFolderName As String = Path.GetFileName(Path.GetDirectoryName(filename))

                            _originalFiledest = String.Empty
                            _originalFiledest = destloc & foldername & "\" & dtnow.ToString() & "\VDrive" & "\Original"
                            _originalFiledest = _originalFiledest & "\" & lastFolderName
                            createfolder(_originalFoldername & "\" & lastFolderName, destloc, String.Empty)
                            _originalDest = System.IO.Path.Combine(_originalFiledest, file)

                            'checking if file is exist with same date it will skip and check for next folder.
                            If System.IO.File.Exists(_originalDest) = False Then
                                Dim FileExtension As String = "." + file.Split(".").Last.ToLower
                                If FileExtension <> ".db" Then
                                    filedest = String.Empty
                                    filedest = destloc & foldername & "\" & dtnow.ToString & "\VDrive"
                                    filedest = filedest & "\" & lastFolderName
                                    createfolder(foldername, destloc, dtnow.ToString & "\VDrive" & "\" & lastFolderName)
                                    'getting each client name for create folder with client name.
                                    destination = System.IO.Path.Combine(filedest, dr("C_Code").ToString.Trim & "_" & id & FileExtension)

                                    'System.IO.File.Copy(filename, destination, True)
                                    FMove(filename, destination)
                                    'System.IO.File.Copy(filename, _originalDest, True)
                                    FMove(filename, _originalDest)

                                    Try
                                        Dim contentType1 As String = ContentType(FileExtension)
                                        Dim fs As FileStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                                        Dim br As BinaryReader = New BinaryReader(fs)
                                        'invoice's all file formates converting into bytes 
                                        Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
                                        br.Close()
                                        fs.Close()
                                        'file_info table  insert query. data converting and inserting into Database
                                        Using cmd As New SqlCommand("Insert into EMSDB_FileInfo (FI_OriginalName, FI_ReceiptDate, FI_Source, FI_ClientCode, FI_FileName, FI_ContentType, FI_Data, FI_CreatedOn, FI_Status) Values (@FI_OriginalName, @FI_ReceiptDate, @FI_Source, @FI_ClientCode, @FI_FileName, @FI_ContentType, @FI_Data, @FI_CreatedOn, @FI_Status)", con)
                                            cmd.Parameters.AddWithValue("@FI_OriginalName", file)
                                            cmd.Parameters.AddWithValue("@FI_ReceiptDate", Today)
                                            cmd.Parameters.AddWithValue("@FI_Source", "VDrive")
                                            cmd.Parameters.AddWithValue("@FI_ClientCode", dr("C_Code").ToString())
                                            cmd.Parameters.AddWithValue("@FI_FileName", dr("C_Code").ToString.Trim & "_" & id & FileExtension)
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
                                    'auto increment values based on FI_ID
                                    id = id + 1
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        
        Catch ex As Exception
            Dim message1 As MailMessage = New MailMessage()
            Dim smtp As SmtpClient = New SmtpClient()
            'Once download complets send mail to RTM support.
            message1.From = New MailAddress("BLR-RTM-Server@tangoe.com")
            message1.To.Add(New MailAddress("RTM-Support@tangoe.com"))
            message1.To.Add(New MailAddress("namohar.m@tangoe.com"))
            message1.Subject = "EMSDB File Download Error"
            message1.Body = ex.Message
            message1.IsBodyHtml = False
            smtp.Port = 25
            smtp.Host = "outlook-south.tangoe.com"
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network
            smtp.EnableSsl = False
            smtp.Send(message1)
        End Try
    End Sub

    'code changed by namohar 01-03-2017
    Private Sub Download_S_Drive_files()
        Try

            If Date.Now.DayOfWeek = DayOfWeek.Sunday Or Date.Now.DayOfWeek = DayOfWeek.Saturday Then
                Exit Sub
            End If
            If Date.Now.DayOfWeek = DayOfWeek.Monday Then
                dtnow = Date.Now.AddDays(-3).ToString("MM-dd-yyyy")
            Else
                dtnow = Date.Now.AddDays(-1).ToString("MM-dd-yyyy")
            End If
            Dim result As String = GetID()
            Dim id As Integer = Convert.ToInt32(result) + 1
            dt = New DataTable()
            'get HP_Alcatel_LucentClient details.
            GetHP_Alcatel_LucentClient()
            For Each dr As DataRow In dt.Rows
                Dim foldername As String = String.Empty
                Dim _originalFoldername As String = String.Empty
                Dim filesource As String = String.Empty
                Dim filedest As String = String.Empty
                Dim _originalFiledest As String = String.Empty
                Dim destination As String = String.Empty
                Dim _originalDest As String = String.Empty

                foldername = "\Rename_HP_Alcatel Lucent"
                _originalFoldername = foldername & "\" & dtnow.ToString() & "\SDrive" & "\Original"

                '\\isgfs1\pdf_repository\HP ALU\02232017
                filesource = Sdrive_sourceloc & "\" & dtnow.ToString
                filedest = destloc & foldername
                _originalFiledest = destloc & _originalFoldername
                'create folder in local repository original.
                createfolder(foldername, destloc, dtnow.ToString)
                'create folder in local repository 
                createfolder(_originalFoldername, destloc, String.Empty)
                'check if folder already exist with same date.
                If CheckCreationDate(filedest) = False Then
                    '\\isgfs1\pdf_repository\HP ALU
                    'filesource:\\vmdalapp01.symphonycmg.com\va\Renaming\Rename_DHL Excel\02232017
                    If Not Directory.Exists(filesource) Then
                    Else

                        Dim direc
                        'get all client folder and file's in Sdrive 
                        For Each direc In Directory.GetDirectories(filesource, "*", SearchOption.AllDirectories)
                            'loop through each file in Sdrive Directory
                            For Each filename As String In Directory.GetFiles(direc)
                                Dim file As String = Path.GetFileName(filename)
                                'Actaul folder name 
                                Dim lastFolderName As String = Path.GetFileName(Path.GetDirectoryName(filename))

                                _originalFiledest = String.Empty
                                _originalFiledest = destloc & foldername & "\" & dtnow.ToString() & "\SDrive" & "\Original"
                                _originalFiledest = _originalFiledest & "\" & lastFolderName
                                createfolder(_originalFoldername & "\" & lastFolderName, destloc, String.Empty)
                                _originalDest = System.IO.Path.Combine(_originalFiledest, file)

                                If System.IO.File.Exists(_originalDest) = False Then
                                    Dim FileExtension As String = "." + file.Split(".").Last.ToLower
                                    filedest = String.Empty
                                    filedest = destloc & foldername & "\" & dtnow.ToString & "\SDrive"
                                    filedest = filedest & "\" & lastFolderName
                                    'Create folder in Auto downloaded path.
                                    createfolder(foldername, destloc, dtnow.ToString & "\SDrive" & "\" & lastFolderName)
                                    'get client file's  in Sdrive
                                    destination = System.IO.Path.Combine(filedest, dr("C_Code").ToString.Trim & "_" & id & FileExtension)

                                    'copy file's to Auto downloaded path with file name
                                    FMove(filename, destination)
                                    'copy file's to Auto downloaded path within original folder name
                                    FMove(filename, _originalDest)

                                    Try

                                        Dim contentType1 As String = ContentType(FileExtension)
                                        Dim fs As FileStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                                        Dim br As BinaryReader = New BinaryReader(fs)
                                        'file's were converted  into bytes 
                                        Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
                                        br.Close()
                                        fs.Close()
                                        'file info table  insert query.
                                        Using cmd As New SqlCommand("Insert into EMSDB_FileInfo (FI_OriginalName, FI_ReceiptDate, FI_Source, FI_ClientCode, FI_FileName, FI_ContentType, FI_Data, FI_CreatedOn, FI_Status) Values (@FI_OriginalName, @FI_ReceiptDate, @FI_Source, @FI_ClientCode, @FI_FileName, @FI_ContentType, @FI_Data, @FI_CreatedOn, @FI_Status)", con)
                                            cmd.Parameters.AddWithValue("@FI_OriginalName", file)
                                            cmd.Parameters.AddWithValue("@FI_ReceiptDate", Today)
                                            cmd.Parameters.AddWithValue("@FI_Source", "SDrive")
                                            cmd.Parameters.AddWithValue("@FI_ClientCode", dr("C_Code").ToString())
                                            cmd.Parameters.AddWithValue("@FI_FileName", dr("C_Code").ToString.Trim & "_" & id & FileExtension)
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
                                    'auto increment values based on FI_ID
                                    id = id + 1

                                End If
                            Next
                        Next
                        'get all client folder and file's in Sdrive 
                        'loop through each file in Sdrive Directory
                        For Each filename As String In Directory.GetFiles(filesource)
                            Dim file As String = Path.GetFileName(filename)
                            Dim lastFolderName As String = Path.GetFileName(Path.GetDirectoryName(filename))

                            _originalFiledest = String.Empty
                            _originalFiledest = destloc & foldername & "\" & dtnow.ToString() & "\SDrive" & "\Original"
                            _originalFiledest = _originalFiledest & "\" & lastFolderName
                            createfolder(_originalFoldername & "\" & lastFolderName, destloc, String.Empty)
                            _originalDest = System.IO.Path.Combine(_originalFiledest, file)

                            If System.IO.File.Exists(_originalDest) = False Then
                                Dim FileExtension As String = "." + file.Split(".").Last.ToLower
                                filedest = String.Empty
                                filedest = destloc & foldername & "\" & dtnow.ToString & "\SDrive"
                                filedest = filedest & "\" & lastFolderName
                                createfolder(foldername, destloc, dtnow.ToString & "\SDrive" & "\" & lastFolderName)
                                destination = System.IO.Path.Combine(filedest, dr("C_Code").ToString.Trim & "_" & id & FileExtension)

                                'copy file's to Auto downloaded path with file name
                                FMove(filename, destination)
                                'copy file's to Auto downloaded path within original folder name
                                FMove(filename, _originalDest)

                                Try
                                    Dim contentType1 As String = ContentType(FileExtension)
                                    Dim fs As FileStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
                                    Dim br As BinaryReader = New BinaryReader(fs)
                                    Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
                                    br.Close()
                                    fs.Close()
                                    'file info table  insert query.
                                    Using cmd As New SqlCommand("Insert into EMSDB_FileInfo (FI_OriginalName, FI_ReceiptDate, FI_Source, FI_ClientCode, FI_FileName, FI_ContentType, FI_Data, FI_CreatedOn, FI_Status) Values (@FI_OriginalName, @FI_ReceiptDate, @FI_Source, @FI_ClientCode, @FI_FileName, @FI_ContentType, @FI_Data, @FI_CreatedOn, @FI_Status)", con)
                                        cmd.Parameters.AddWithValue("@FI_OriginalName", file)
                                        cmd.Parameters.AddWithValue("@FI_ReceiptDate", Today)
                                        cmd.Parameters.AddWithValue("@FI_Source", "SDrive")
                                        cmd.Parameters.AddWithValue("@FI_ClientCode", dr("C_Code").ToString())
                                        cmd.Parameters.AddWithValue("@FI_FileName", dr("C_Code").ToString.Trim & "_" & id & FileExtension)
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
                                'auto increment values
                                id = id + 1

                            End If
                        Next
                    End If
                End If
            Next
       
        Catch ex As Exception
            Dim message1 As MailMessage = New MailMessage()
            Dim smtp As SmtpClient = New SmtpClient()
            'sending mail to rtm support team.
            message1.From = New MailAddress("BLR-RTM-Server@tangoe.com")
            message1.To.Add(New MailAddress("RTM-Support@tangoe.com"))

            message1.Subject = "EMSDB File Download Error"
            message1.Body = ex.Message
            message1.IsBodyHtml = False
            smtp.Port = 25
            smtp.Host = "outlook-south.tangoe.com"
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network
            smtp.EnableSsl = False
            smtp.Send(message1)
        End Try
    End Sub
    'code changed by namohar 01-03-2017
    Private Sub GetHP_Alcatel_LucentClient()
        dt = New DataTable()
        Using da As New SqlDataAdapter("select * from EMSDB_Clients where C_Status=1 and C_ID=8", con)
            da.Fill(dt)
        End Using
    End Sub

    'copy file from V or S drive and put into \\10.80.20.251\emsdb_invpro\Automated_Download 
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

    End Sub
    'create client folders in \\10.80.20.251\emsdb_invpro\Automated_Download"
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

    'get maximum FI_ID in EMSDB_FileInfo table 
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
    'input invoice file formates 
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
    'getting all client file's in S or V drive 
    Public Shared Function GetAllFiles(ByVal _directory As [String]) As List(Of [String])
        'Return Directory.GetFiles(_directory, "*.*", SearchOption.AllDirectories).ToList()
        Return Directory.EnumerateFiles(_directory, "*", SearchOption.AllDirectories).ToList
    End Function
    'checking folder name if folder exist with same date it will skip and go to next folder...
    Private Function CheckCreationDate(ByVal dirpath As String) As Boolean
        Dim di As New DirectoryInfo(dirpath)
        Dim dirs() As DirectoryInfo = di.GetDirectories()
        Dim creationTime As DateTime
        Dim dirsize As Double

        For Each dir As DirectoryInfo In dirs
            creationTime = dir.LastAccessTime.Date
            'dirsize = FldrSize(dirpath)
        Next
        If creationTime <> Now.Date.AddDays(-1) Or dirsize = 0 Then
            Return False
        ElseIf creationTime = Now.Date.AddDays(-1) Then
            Return True
        End If
        If dirs.Length = 0 Then
            Return False
        End If
    End Function

    'writting error log in \\10.80.20.251\emsdb_invpro\AutomatedDownload_ErrorLog"
    Public Sub To_WriteError(ByVal pstrErrorDes As String)

        Dim strFile As String = "\\10.80.20.251\emsdb_invpro\Log\AutomatedDownload_ErrorLog.txt"

        If (Not File.Exists(strFile)) Then
            Dim sw As New System.IO.StreamWriter(strFile, False)
            sw = File.CreateText(strFile)
            sw.WriteLine(pstrErrorDes & DateTime.Now)
            sw.Close()
        Else
            Dim sw As New System.IO.StreamWriter(strFile, True)
            sw.WriteLine(pstrErrorDes & DateTime.Now)
            sw.Close()
        End If
    End Sub
    'get active clients in EMSDB_Clients table.
    Private Sub GetClients()
        dt = New DataTable()
        Using da As New SqlDataAdapter("select * from EMSDB_Clients where C_Status=1", con)
            da.Fill(dt)
        End Using
    End Sub
    'Auto download process , we are sheduling time slots for auto download.
    Private Sub bgLongRunningProcess_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgLongRunningProcess.DoWork
        Dim start1 As TimeSpan = New TimeSpan(6, 45, 0)
        Dim end1 As TimeSpan = New TimeSpan(6, 50, 0)
        Dim start2 As TimeSpan = New TimeSpan(9, 15, 0)
        Dim end2 As TimeSpan = New TimeSpan(9, 20, 0)
        Dim start3 As TimeSpan = New TimeSpan(6, 51, 0)
        Dim end3 As TimeSpan = New TimeSpan(6, 56, 0)
        Dim now As TimeSpan = TimeSpan.Parse(DateTime.Now.TimeOfDay.ToString("hh\:mm\:ss"))
        If ((now > start1) AndAlso (now < end1)) Then
            Download_files_New()

        End If
        If ((now > start2) AndAlso (now < end2)) Then
            Download_files_New()

        End If

        If ((now > start3) AndAlso (now < end3)) Then
            Download_S_Drive_files()
        End If

    End Sub

    Private Sub bgLongRunningProcess_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgLongRunningProcess.RunWorkerCompleted

        bgLongRunningProcess.RunWorkerAsync()
    End Sub


    Private Sub bgLongRunningProcess_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgLongRunningProcess.ProgressChanged
        'ProgressBar1.Value = e.ProgressPercentage
    End Sub
End Class
