Imports System.Data
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Security.Policy
Imports Microsoft.Data.SqlClient
Imports Windows.Media
Imports Microsoft.Office.Interop.Outlook
Imports System.Net.Mail
Imports Windows.Networking
Imports System.Runtime.InteropServices
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Configuration
Imports System.Windows.Forms.FolderBrowserDialog

Class MainWindow

    'Dim savePdfFolder As String = "I:\AppData\dba\Technology Review - Semi Annual\SQL SERVER\_2024\Q3_Q4\reports"
    Dim labelSavePdfFolder As String = "Save PDF To Folder:"
    Dim labelSaveSentEmailFolder As String = "Save Sent Email To Folder:"

    Dim savePdfFolder As String = DBAccessReviewSender.Settings.Default.PDF_Folder
    Dim saveSentEmailFolder As String = DBAccessReviewSender.Settings.Default.Draft_Email_Folder
    '"I:\AppData\dba\Technology Review - Semi Annual\ORACLE\_Databases\_On Premise Oracle Databases\reports"
    '"I:\AppData\dba\Technology Review - Semi Annual\SQL SERVER\_2024\Q3_Q4\sent_email"

    Dim ToggleIsSelectedState = False
    Dim carbonCopyAppAdmins = "debwong@fhb.com"

    Dim DueDate As Date = DBAccessReviewSender.Settings.Default.Response_Due_Date

    Dim sqlStr As String = "SELECT GroupName
, Status
, IsSelected
, HostNm
, EnvCd
, ManagerFullNm
, Attachment
, ManagerEmpNr
, CC
, ManagerEmail
, RevRptID FROM dbo.TReviewRptsGrouper
ORDER BY  GroupName, HostNm, EnvCd"

    Dim builder As New SqlConnectionStringBuilder With {
            .InitialCatalog = DBAccessReviewSender.Settings.Default.SQL_DB_Name,
            .DataSource = DBAccessReviewSender.Settings.Default.SQL_Server_Name,
            .IntegratedSecurity = True,
            .ConnectTimeout = 3,
            .Encrypt = 0
        }

    Dim targetConn As New SqlConnection(builder.ConnectionString)


    Dim aDataSet As DataSet
    Dim adapter As New SqlDataAdapter

    Dim StatusType = New Dictionary(Of String, String) From {{"PDFExported", "Report Exported"},
                                                             {"EmailSent", "Email Sent"},
                                                             {"Approved", "Approved"},
                                                             {"CSARPending", "Access Request Pending"},
                                                             {"CSARComplete", "Access Request Completed"},
                                                             {"OtherAction", "Other follow-up in progress"}}


    Public Sub DraftEmailOutlook()
        Dim msgTemplateFilePath As String
        msgTemplateFilePath = savePdfFolder & "\EmailTemplate.oft"
        Dim appOutlook As New Microsoft.Office.Interop.Outlook.Application()
        'Dim msg As appOutlook.CreateItemFromTemplate(msgTemplateFilePath)
        ' Loop through the rows in the data grid

        ' Sort the data grid by Group (not working)
        ' Dim colIdx As Int32
        'colIdx = 4
        'Dim sortColumn = Grid1.Columns[colIdx]
        '        Grid1.Items.SortDescriptions.Clear()
        'Grid1.Items.SortDescriptions.Add(New ComponentModel.SortDescription(Column.SortMemberPath, ListSortDirection.Ascending.))

        Dim currentGroup As String
        Dim priorGroup As String
        currentGroup = ""
        priorGroup = ""


        For Each row As DataRowView In Grid1.ItemsSource

            If row.Item("IsSelected") = False Then
                'Return break exit next?
                Continue For
            End If

            Dim msg
            Dim attachments

            ' new group? 
            currentGroup = row.Item("GroupName")
            If currentGroup <> priorGroup Then

                ' Optional save msg to email sent folder. 
                'If msg IsNot Nothing Then
                '    msg.SaveAs(saveSentEmailFolder + "\" + row.Item("RevRptID") + ".msg")
                'End If


                msg = appOutlook.CreateItemFromTemplate(msgTemplateFilePath)
                '********          msg.Display()
                attachments = msg.Attachments
                'Dim toEmails = "pking@fhb.m"
                '#pktodo remove
                Dim toEmails = row.Item("ManagerEmail")

                DueDate = DueDatePicker.SelectedDate
                Dim strDueDate As String = DueDate.ToLongDateString()

                Dim tempBody As String = msg.Body
                'tempBody = tempBody.Replace("##DUE_DATE", "9/30/2024")
                tempBody = tempBody.Replace("##DUE_DATE", strDueDate)
                msg.Body = tempBody

                'If (Not IsDBNull(row.Item("CCEmails"))) Then
                '    If Not String.IsNullOrWhiteSpace(row.Item("CCEmails")) Then
                '        toEmails = toEmails & row.Item("CCEmails")
                '    End If

                'End If
                Dim AdditionalRecipients
                If IsDBNull(row.Item("CC")) Then
                    row.Item("CC") = String.Empty
                End If

                If row.Item("CC") <> String.Empty Then
                    AdditionalRecipients = carbonCopyAppAdmins + ";" + row.Item("CC")
                Else
                    AdditionalRecipients = carbonCopyAppAdmins
                End If
                msg.CC = AdditionalRecipients

                'toEmails = toEmails & ";" & carbonCopyAppAdmins

                msg.Recipients.Add(toEmails)
                ' reset prior as we are in a new group at this point
                priorGroup = currentGroup

            End If

            ' Attach PDF Access Review Report
            'Dim olByValue As Microsoft.Office.Interop.Outlook.olb

            Dim attFilePath As String
            attFilePath = savePdfFolder & "\" & row.Item("Attachment")
            attachments.Add(attFilePath)
            msg.Save()
            row.Item("Status") = "Notification Drafted"
            'msg.Close()
        Next

        ' Need to send the last email at this point.  

        Marshal.FinalReleaseComObject(appOutlook)
        'appOutlook.Close()
        'appOutlook.Quit()
        MsgBox("Check you Outlook drafts folder for the results",, "Completed")

    End Sub


    Public Sub ExportPDFs()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        'Dim myURI = "http//itd-rpt-dossrst/ReportServer?%2FData+Ops%2FOracle+Database+User+Account+Reval+-+ALL&HostName_Filter_Parameter=%%&Host_Name_Parameter=dstrp&rs:Command = Render&rs:Format = PDF"
        Dim myURI = ""
        Dim BaseURI = "http://itd-rpt-dossrst/ReportServer?/Data%20Ops/Oracle%20Database%20User%20Account%20Reval%20-%20By%20Users%20Manager"

        ' Dim t As Task = GetUrlContentAsync(myURI)
        ' Dim t As Task = DownloadContentAsync(myURI, "C\localdocs\ws\_latest_AccessReview_app\reports\report.pdf")
        ' Dim t As Task = DownloadContentAsync(myURI, savePdfFolder + "\" + " report.pdf")

        't.Wait()
        Dim reportParms As String = ""

        For Each row As DataRowView In Grid1.ItemsSource
            ' we need to pass in mgr# and attachmentname

            If row.Item("IsSelected") = True Then

                If row.Item("Attachment") IsNot Nothing And row.Item("ManagerEmpNr") IsNot Nothing Then
                    ' &HostName_Filter_Parameter=%%&Host_Name_Parameter=cme_t &PARM_ManagerEmpNr=9999   &rs:Command = Render&rs:Format = PDF"
                    reportParms = "&HostName_Filter_Parameter=%%&Host_Name_Parameter=" & row.Item("HostNm") & "&PARM_ManagerEmpNr=" & row.Item("ManagerEmpNr") & "&rs:Command=Render&rs:Format=PDF"
                    myURI = BaseURI + reportParms
                    Dim t As Task = DownloadContentAsync(myURI, savePdfFolder & "\" & row.Item("Attachment"))
                    Threading.Thread.Sleep(500)

                    row.Item("Status") = StatusType.Item("PDFExported")
                End If
            End If

        Next

        MsgBox("Exports Completed",, "Completed")


    End Sub

    Async Function GetUrlContentAsync(ByVal url As String) As Task
        Using client As New HttpClient()

            ' client.
            Dim response As HttpResponseMessage = Await client.GetAsync(url)

            If response.IsSuccessStatusCode Then
                Dim content As String = Await response.Content.ReadAsStringAsync()
                Console.WriteLine(content)
            Else
                Console.WriteLine("Error:   " & response.StatusCode.ToString())
            End If
        End Using
    End Function

    Async Function DownloadContentAsync(ByVal url As String, ByVal outputPath As String) As Task

        Dim handler As New HttpClientHandler()
        handler.UseDefaultCredentials = True

        Using client As New HttpClient(handler)

            Dim response As HttpResponseMessage = Await client.GetAsync(url)
            response.EnsureSuccessStatusCode()

            Using stream As Stream = Await response.Content.ReadAsStreamAsync()
                Using fileStream As FileStream = File.Create(outputPath)
                    Await stream.CopyToAsync(fileStream)
                End Using
            End Using

        End Using
    End Function

    Public Function SelectRows(queryStr As String, conn As SqlConnection, adapter As SqlDataAdapter) As DataSet


        Dim myDataSet As New DataSet
        'Try
        If conn.State <> ConnectionState.Open Then
            conn.Open()
        End If
        With adapter
            .SelectCommand = New SqlCommand(queryStr) With {
                .Connection = conn
            }
        End With

        adapter.Fill(myDataSet)
        conn.Close()
        'Catch ex As Exception

        'End Try
        Return myDataSet

    End Function

    Public Sub SaveRows()


        If targetConn.State <> ConnectionState.Open Then
            targetConn.Open()
        End If

        Dim commandBuilder As New SqlCommandBuilder(adapter)
        adapter.UpdateCommand = commandBuilder.GetUpdateCommand()
        adapter.Update(aDataSet.Tables(0))

        aDataSet.Tables(0).AcceptChanges()

        ' Save Application Settings as well
        DBAccessReviewSender.Settings.Default.PDF_Folder = savePdfFolder
        DBAccessReviewSender.Settings.Default.Draft_Email_Folder = saveSentEmailFolder

        DueDate = DueDatePicker.SelectedDate
        DBAccessReviewSender.Settings.Default.Response_Due_Date = DueDate
        ' DBAccessReviewSender.Settings.Default.Admin_Emails = carbonCopyAppAdmins   ' <-- no UI to update these settings. 
        DBAccessReviewSender.Settings.Default.Save()

    End Sub

    Public Sub LoadServerList()

        ' This call is required by the designer. "{d:SampleData ItemCount=5}"/>
        InitializeComponent()

        LabelPDFPath.Content = labelSavePdfFolder + " " + savePdfFolder
        LabelSentEmailPath.Content = labelSaveSentEmailFolder + " " + saveSentEmailFolder
        Label_DB_Name.Content = DBAccessReviewSender.Settings.Default.SQL_DB_Name

        ' add filters if set
        If TextSidFilter.Text <> "SID filter" And TextSidFilter.Text <> "" Then
            TextSidFilter.Text = TextSidFilter.Text.Replace(";", "")
            If sqlStr.Contains("WHERE") Then
                sqlStr = sqlStr + " AND HostNm LIKE '%" + TextSidFilter.Text + "%'"
            Else
                sqlStr = sqlStr + " WHERE HostNm LIKE '%" + TextSidFilter.Text + "%'"
            End If
        End If

        ' fill the dataset via the sql query
        aDataSet = SelectRows(sqlStr, targetConn, adapter)
        ' bind the data grid to the dataset (table 0)
        Grid1.DataContext = aDataSet.Tables(0).DefaultView
        DueDate = DBAccessReviewSender.Settings.Default.Response_Due_Date
        DueDatePicker.SelectedDate = DueDate


    End Sub

    Private Sub DataGrid1_AutoGeneratingColumn(sender As Object, e As DataGridAutoGeneratingColumnEventArgs)

        If e.PropertyName = "Status" Then
            Dim cb = New DataGridComboBoxColumn()
            cb.Header = "Status"
            cb.Width = "200"
            cb.ItemsSource = New String() {"Not started",
                        "Report Exported",
                        "Notification Drafted",
                        "Notification Sent",
                        "Reviewed - Approved As Is",
                        "Reviewed - Changes",
                        "Reviewed - Follow-up required",
                        "Reminder Sent",
                        "No response",
                        "Completed - closed"}
            cb.SelectedValueBinding = New Binding("Status")
            e.Column = cb
        End If

    End Sub


    Public Sub ToggleSelectAll()

        If ToggleIsSelectedState = True Then
            ToggleIsSelectedState = False
        Else
            ToggleIsSelectedState = True
        End If

        For Each row As DataRowView In Grid1.ItemsSource

            row.Item("IsSelected") = ToggleIsSelectedState

        Next

    End Sub


    Public Sub PdfFolderPicker()
        Dim fp = New Forms.FolderBrowserDialog()
        fp.InitialDirectory = savePdfFolder
        If (fp.ShowDialog) = System.Windows.Forms.DialogResult.OK Then
            savePdfFolder = fp.SelectedPath
            LabelPDFPath.Content = labelSavePdfFolder + " " + savePdfFolder
        End If

    End Sub

    Public Sub EmailFolderPicker()
        Dim fp = New Forms.FolderBrowserDialog()
        fp.InitialDirectory = savePdfFolder
        If (fp.ShowDialog) = System.Windows.Forms.DialogResult.OK Then
            saveSentEmailFolder = fp.SelectedPath
            LabelSentEmailPath.Content = labelSaveSentEmailFolder + " " + saveSentEmailFolder
        End If

        ' 
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub OnLoad(sender As Object, e As RoutedEventArgs)
        Call LoadServerList()
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Call LoadServerList()
    End Sub

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        Call ToggleSelectAll()
    End Sub

    Private Sub EmailFolder_Button_Click(sender As Object, e As RoutedEventArgs) 'Handles ButtonChangeEmailFolder.Click
        Call EmailFolderPicker()


    End Sub

    Private Sub PDFFolder_Button_Click(sender As Object, e As RoutedEventArgs) 'Handles ButtonChangePDFFolder.Click
        Call PdfFolderPicker()
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As RoutedEventArgs) Handles SaveButton.Click
        Call SaveRows()

    End Sub

    Private Sub ExportPDFsButton_Click(sender As Object, e As RoutedEventArgs) Handles ExportPDFsButton.Click
        Call ExportPDFs()
    End Sub

    Private Sub Button_Click_3(sender As Object, e As RoutedEventArgs)

        Call DraftEmailOutlook()
    End Sub

    Private Sub Button_Click_4(sender As Object, e As RoutedEventArgs)
        Close()
    End Sub
End Class

