Imports DJLib
Imports DJLib.Dbtools
Imports System.ComponentModel
Imports System.IO

Public Class FormDownloadCurrency
    Private WithEvents backgroundworker1 As New BackgroundWorker
    Private WithEvents backgroundworker2 As New BackgroundWorker
    Private WithEvents Download1 As Download
    Private useProxy As Boolean
    Private ProxyServer As String
    Private User As String
    Private Password As String
    Private Server As String

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim DataCollection As New Collection
        Dim dbtools1 As New Dbtools(myUserid, myPassword)
        Dim sqlstr As String = "select pdt.cvalue,pdt.paramname from rmparamdt pdt" & _
                                " left join rmparamhd ph on ph.paramhdid = pdt.paramhdid" & _
                                " where ph.paramname = 'DownloadOptions' and pdt.paramname <> 'useproxy' " & _
                                " order by pdt.ivalue"
        Dim errorMessage As String = vbEmpty
        If dbtools1.GetDataCollection(sqlstr, DataCollection, errorMessage) Then
            sqlstr = "select pdt.bvalue,pdt.paramname from rmparamdt pdt left join rmparamhd ph on ph.paramhdid = pdt.paramhdid where ph.paramname = 'DownloadOptions' and pdt.paramname = 'useproxy' order by pdt.ivalue"
            dbtools1.GetDataCollection(sqlstr, DataCollection, errorMessage)
            TextBox1.Text = DataCollection.Item("url").ToString
            TextBox2.Text = DataCollection.Item("saveto").ToString
            useProxy = DataCollection.Item("useproxy").ToString
            ProxyServer = DataCollection.Item("proxyserver").ToString
            User = DataCollection.Item("username").ToString
            Password = DataCollection.Item("password").ToString
            Server = DataCollection.Item("server").ToString
            'Download begin
            Call StartDownload()
        Else
            MsgBox(errorMessage)
        End If
    End Sub

    Private Sub ButtonStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStop.Click
        If backgroundworker1.IsBusy Then
            Dim i As Integer = MsgBox("Do you want to stop the process?", vbYesNo)
            If i = 6 Then
                backgroundworker1.CancelAsync()
            End If
        End If
        Me.Close()
    End Sub

    Private Sub StartDownload()
        Dim Message As String = vbEmpty
        backgroundworker1.WorkerReportsProgress = True
        backgroundworker1.WorkerSupportsCancellation = True
        MoveToFolder(My.Settings.tmpFolder, TextBox2.Text & "\" & Path.GetFileName(TextBox1.Text), Message)
        Download1 = New Download(TextBox1.Text, TextBox2.Text, useProxy, ProxyServer, User, Password, Server)
        ProgressBar1.Value = 0
        backgroundworker1.RunWorkerAsync()
    End Sub

    Private Sub backgroundworker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundworker1.DoWork
        Dim message As String = String.Empty
        Download1.Start(message, backgroundworker1, e)
        If Download1.getResult Then
            backgroundworker1.ReportProgress(1, "Download Completed")
            'FormMenu.setBubbleMessage("Download Currency", "Download Complete")
        Else
            FormMenu.setBubbleMessage("Download Currency", message)
            backgroundworker1.ReportProgress(1, message)
        End If
    End Sub

    Private Sub Download1_StateChange(ByVal progress As Long, ByVal obj As Object) Handles Download1.StateChange
        backgroundworker1.ReportProgress(progress, obj)
    End Sub

    Private Sub backgroundworker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles backgroundworker1.ProgressChanged
        If backgroundworker1.CancellationPending = True Then
            FormMenu.setBubbleMessage("Download Currency", "User Cancelled")
        End If
        If e.ProgressPercentage = 1 Then
            TextDesc.Text = e.UserState
        ElseIf e.ProgressPercentage = 6 Then
            ProgressBar1.Style = ProgressBarStyle.Continuous
        ElseIf e.ProgressPercentage = 7 Then
            ProgressBar1.Style = ProgressBarStyle.Marquee
        Else
            If Download1.Filesize > 0 Then
                ProgressBar1.Maximum = Download1.Filesize
                ProgressBar1.Value = e.ProgressPercentage
                TextBox3.Text = Download1.bytesToLabel(Download1.Filesize)
                TextDesc.Text = Format(ProgressBar1.Value / ProgressBar1.Maximum, "0%") & " [" & Download1.bytesToLabel(e.ProgressPercentage) & "]" 'Download1.progress.ToString & " of " & Download1.Filesize.ToString & " " & e.ProgressPercentage & " " & e.UserState  
                Me.Text = "[" & Format(ProgressBar1.Value / ProgressBar1.Maximum, "0%") & "] - Download Currency"          
            End If
        End If
    End Sub

    Private Sub backgroundworker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundworker1.RunWorkerCompleted
        If Download1.getResult Then
            If cbImportDb.Checked Then
                Call ImportData()
            End If
        End If
    End Sub

    Private Sub ImportData()
        backgroundworker2.WorkerReportsProgress = True
        backgroundworker2.WorkerSupportsCancellation = True
        backgroundworker2.RunWorkerAsync()
    End Sub

    Private Sub backgroundworker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundworker2.DoWork
        Call Import()
    End Sub

    Private Sub backgroundworker2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles backgroundworker2.ProgressChanged
        If e.ProgressPercentage = 1 Then
            TextDesc.Text = e.UserState
        End If
    End Sub

    Private Sub backgroundworker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundworker2.RunWorkerCompleted
        FormMenu.setBubbleMessage("Download Currency", "Import Complete")
        If CheckBoxClose.Checked Then
            Me.Close()
        End If
    End Sub

    Private Sub Import()
        Dim dbtools1 As New Dbtools(myUserid, myPassword)
        Dim list As New List(Of String)
        Dim myHeader As String()
        Dim myHeader2 As String()
        Dim myDetail As String()
        Dim DateTx As String
        Dim line As String
        Dim myFullPath As String = System.AppDomain.CurrentDomain.BaseDirectory & "mytext.txt"
        Dim myvalue As String
        Dim sqlstr As String
        Dim result As Boolean
        Dim message As String = String.Empty
        Dim startdate As Date = Today
        Dim enddate As Date = Today
        Dim currdate As Date = Today
        backgroundworker2.ReportProgress(7, "Marquee..")
        backgroundworker2.ReportProgress(1, "Create Temp File..")
        If File.Exists(myFullPath) Then
            File.Delete(myFullPath)
        End If
        Dim fs As FileStream
        fs = File.Create(myFullPath)
        fs.Close()
        backgroundworker2.ReportProgress(1, "Open File..")

        'Using myStream As StreamReader = File.OpenText(Download1.getSaveTo & "\" & Path.GetFileName(Download1.getUrl))
        Using myStream As StreamReader = File.OpenText(Download1.getSaveTo & "\2.csv")
            line = myStream.ReadLine
            Do While (Not line Is Nothing)
                list.Add(line)
                line = myStream.ReadLine
            Loop
        End Using
        'myHeader = list(7).Split(New Char() {";"c})
        myHeader = list(2).Split(New Char() {";"c})
        myHeader2 = list(0).Split(New Char() {";"c})
        'convert to file
        backgroundworker2.ReportProgress(1, "Converting File..")
        Using writer As StreamWriter = New StreamWriter(myFullPath)
            For i = 6 To list.Count - 1
                myDetail = list(i).Split(";")
                If myDetail.Length > 1 Then
                    If IsNumeric(myDetail(1)) Then
                        DateTx = DateFormatDDMMYYYY(myDetail(0))
                        currdate = CDate(String.Format("{0}-{1}-{2}", myDetail(0).Substring(6, 4), myDetail(0).Substring(3, 2), myDetail(0).Substring(0, 2)))
                        'get startdate and enddate
                        If startdate > currdate Then
                            startdate = currdate
                        End If

                        If enddate < currdate Then
                            enddate = currdate
                        End If
                        For j = 1 To myDetail.Length - 1
                            myvalue = dbtools1.ValidNum(myDetail(j))
                            'If myvalue <> "\N" Then
                            If IsNumeric(myvalue) Then
                                'writer.WriteLine(DateTx & vbTab & myHeader(j) & vbTab & myvalue)
                                'writer.WriteLine(DateTx & vbTab & myHeader(j).Substring(myHeader(j).IndexOf("(") + 1, 3) & vbTab & myvalue)
                                If myHeader(j) <> "" Then
                                    writer.WriteLine(DateTx & vbTab & myHeader(j).Substring(myHeader(j).IndexOf("(") + 1, 3) & vbTab & myvalue)
                                Else
                                    writer.WriteLine(DateTx & vbTab & myHeader2(j).Substring(myHeader2(j).IndexOf("(") + 1, 3) & vbTab & myvalue)
                                End If
                                'writer.WriteLine(DateTx & vbTab & myHeader(j).Substring(myHeader(j).IndexOf("(") + 1, 3) & vbTab & myvalue)
                                'Else
                                '    writer.WriteLine(DateTx & vbTab & myHeader(j) & vbTab & "\N")
                            End If
                        Next
                    End If
                End If
            Next
        End Using
        backgroundworker2.ReportProgress(1, "Import to DB..")
        'sqlstr = "Delete from rmcrcytx;select setval('crcytx_crcytxid_seq',1,false);copy rmcrcytx(crcydate,crcycode,crcyamount) from stdin;"
        sqlstr = String.Format("Delete from rmcrcytx  where crcydate >= '{0:yyyy-MM-dd}' and crcydate <= '{1:yyyy-MM-dd}';with seq as (select crcytxid +1 as myint  from rmcrcytx order by crcytxid desc limit 1) select setval('crcytx_crcytxid_seq',max(myint),false) from seq;copy rmcrcytx(crcydate,crcycode,crcyamount) from stdin;", startdate, enddate)
        backgroundworker2.ReportProgress(1, "Importing..")
        result = dbtools1.copy(sqlstr, myFullPath, message)


        If Not result Then
            MsgBox(message)
        Else
            'create table master crcy
            Dim DataSet1 As New DataSet
            sqlstr = "select crcycode from rmcrcytx group by crcycode order by crcycode;"
            dbtools1.getDataSet(sqlstr, DataSet1)
            myFullPath = System.AppDomain.CurrentDomain.BaseDirectory & "crcycode.txt"
            Using writer As StreamWriter = New StreamWriter(myFullPath)
                For Each row As DataRow In DataSet1.Tables(0).Rows
                    writer.WriteLine(row.Item(0).ToString)
                Next
            End Using

            sqlstr = "Delete from rmtbcrcy;select setval('rmtbcrcy_rmtbcrcyid_seq',1,false);copy rmtbcrcy(crcycode) from stdin;"
            result = dbtools1.copy(sqlstr, myFullPath, message)

            backgroundworker2.ReportProgress(1, "Import Completed")
        End If

    End Sub


    Private Sub FormDownloadCurrency_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class