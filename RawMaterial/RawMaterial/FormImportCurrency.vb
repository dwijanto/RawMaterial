Imports System.Threading
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports DJLib
Imports DJLib.Dbtools

Public Class FormImportCurrency
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    'Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim D As New ProgressReportDelegate(AddressOf ProgressReport)
    'Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim OpenFileDialog1 As New OpenFileDialog
    Dim mySelectedPath As String

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        If Not myThread.IsAlive Then

            With OpenFileDialog1
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    mySelectedPath = .FileName
                    Try
                        myThread = New System.Threading.Thread(AddressOf DoWork)
                        myThread.SetApartmentState(ApartmentState.MTA)
                        myThread.Start()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End With
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        Me.Invoke(D, New Object() {2, "Read File.."})
        Dim result As Boolean = ImportTextFile(myselectedpath, errMsg)
        If result Then
            sw.Stop()
            Me.Invoke(D, New Object() {2, String.Format("Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString)})
            Me.Invoke(D, New Object() {2, TextBox2.Text & "Done."})
            Me.Invoke(D, New Object() {5, "Set to continuous mode again"})
        Else
            errSB.Append(errMsg & vbCrLf)
            Me.Invoke(D, New Object(), {3, errSB.ToString})
        End If
        sw.Stop()

    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        Select Case id
            Case 2
                TextBox2.Text = message
            Case 3
                TextBox3.Text = message
            Case 1
                TextBox1.Text = message
            Case 5
                ProgressBar1.Style = ProgressBarStyle.Continuous
            Case 6
                ProgressBar1.Style = ProgressBarStyle.Marquee
            Case 7
                Dim myvalue = message.ToString.Split(",")
                ProgressBar1.Minimum = 1
                ProgressBar1.Value = myvalue(0)
                ProgressBar1.Maximum = myvalue(1)
        End Select
    End Sub

    Private Function ImportTextFile(ByVal FileName As String, ByRef errMsg As String) As Boolean
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
        Me.Invoke(D, New Object() {6, "set to Marque"})
        Using myStream As StreamReader = File.OpenText(FileName)
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
        Me.Invoke(D, New Object() {1, "Converting File.."})
        Me.Invoke(D, New Object() {5, "set to Continuous"})
        Using writer As StreamWriter = New StreamWriter(myFullPath)
            For i = 6 To list.Count - 1
                Me.Invoke(D, New Object() {7, i + 1 & "," & list.Count})
                myDetail = list(i).Split(";")
                If myDetail.Length > 1 Then
                    If IsNumeric(myDetail(1)) Then
                        DateTx = DateFormatDDMMYYYY(myDetail(0))
                        currdate = CDate(String.Format("{0}-{1}-{2}", myDetail(0).Substring(6, 4), myDetail(0).Substring(3, 2), myDetail(0).Substring(0, 2)))
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
                                'Else
                                '    writer.WriteLine(DateTx & vbTab & myHeader(j) & vbTab & "\N")
                            End If
                        Next
                    End If

                    
                End If
            Next
        End Using
        Me.Invoke(D, New Object() {2, "Import to DB.."})
        'sqlstr = "Delete from rmcrcytx;select setval('crcytx_crcytxid_seq',1,false);copy rmcrcytx(crcydate,crcycode,crcyamount) from stdin;"
        sqlstr = String.Format("Delete from rmcrcytx  where crcydate >= '{0:yyyy-MM-dd}' and crcydate <= '{1:yyyy-MM-dd}';with seq as (select crcytxid +1 as myint  from rmcrcytx order by crcytxid desc limit 1) select setval('crcytx_crcytxid_seq',max(myint),false) from seq;copy rmcrcytx(crcydate,crcycode,crcyamount) from stdin;", startdate, enddate)
        Me.Invoke(D, New Object() {1, "Importing.."})
        Me.Invoke(D, New Object() {6, "Set To Marque"})
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
            Me.Invoke(D, New Object() {1, "Import Completed"})
        End If
        Me.Invoke(D, New Object() {5, "Set To continuous"})
        Return True
    End Function


End Class
Public Class myData
    Public Property filename As String
    Public Property rownumber As Long
    Public Property data As Object

    Public Sub New(ByVal filename As String, ByVal rownumber As Long, ByVal data As Object)
        Me.filename = filename
        Me.rownumber = rownumber
        Me.data = data
    End Sub
End Class