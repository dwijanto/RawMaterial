Imports DJLib
Imports DJLib.Dbtools
Imports System.Threading
Public Class FormMenu

    Private DynamicMenu1 As DynamicMenu
    Private MenuStrip1 As MenuStrip
    Private StatustStrip1 As StatusStrip
    Private ToolStripStatusLable1 As ToolStripStatusLabel
    Private WithEvents ToolStripButton1 As ToolStripButton
    Private dbtools1 As New Dbtools(myUserid, myPassword)
    Private DataTable1 As New DataTable

    Public Sub New()

        StatustStrip1 = New System.Windows.Forms.StatusStrip
        ToolStripStatusLable1 = New System.Windows.Forms.ToolStripStatusLabel
        ToolStripButton1 = New System.Windows.Forms.ToolStripButton

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'Me.Text = "App.Version: " & dbtools1.getVer & " :: Server: " & ConnectionStringCollections.Item("HOST") & ", Database: " & ConnectionStringCollections.Item("DATABASE") & ", Userid: " & dbtools1.Userid
        Me.Text = "App.Version: " & Application.ProductVersion & " :: Server: " & ConnectionStringCollections.Item("HOST") & ", Database: " & ConnectionStringCollections.Item("DATABASE") & ", Userid: " & dbtools1.Userid
        Me.Location = New Point(300, 10)
    End Sub

    Private Sub FormMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim errMessage As String = vbNull
        ' If Not dbtools1.getData("Select isactive,programid,parentid,myorder,description,programname,icon,iconindex from program  where isactive order by parentid,myorder", DataTable1, errMessage) Then
        Dim sqlstr As String = String.Empty
        If ConnectionStringCollections.Item("HOST") = "hon14nt" Then
            sqlstr = "Select isactive,programid,parentid,myorder,description,programname,icon,iconindex,_groupname from rmprogram " & _
                                " left join _groupuser gu on gu._membername = '" & myUserid & "'" & _
                                " where isactive and  members ~ '\m" & myGroup & "\y' order by parentid,myorder"
        Else
            sqlstr = "Select isactive,programid,parentid,myorder,description,programname,icon,iconindex,_groupname from rmprogram " & _
                                " left join _groupuser gu on gu._membername = '" & myUserid & "'" & _
                                " where isactive and  members ~ ('\m'" & "|| _groupname ||" & "'\y') order by parentid,myorder"
        End If
        If Not dbtools1.getData(sqlstr, DataTable1, errMessage) Then
            MsgBox(errMessage)
        Else
            If DataTable1.Rows.Count > 0 Then
                DynamicMenu1 = New DynamicMenu(Me, DataTable1, ImageList1)
                'DataGridView1.DataSource = DataTable1
                DynamicMenu1.LoadMenu(MenuStrip1)
                Me.MainMenuStrip = MenuStrip1
                Me.Controls.Add(MenuStrip1)

            Else
                MessageBox.Show("You don't have any access. Please contact admin!", "No menu found for this user")
                Me.Close()
            End If

        End If
    End Sub

    Private Sub MenuItemOnClick_mDownloadOptions(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor

        FormDownloadOptions.ShowDialog()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mDownloadCurrency(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        FormDownloadCurrency.Show()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mImportCurrency(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        FormImportCurrency.Show()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mImportRawMaterial(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        FormImportRawMaterial.Show()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mChart1(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        Chart1.Show()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mChart2(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        Chart2.Show()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mChart3(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        Chart3.Show()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mRawMaterial(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()
        FormRawMaterial.Show()
        Application.DoEvents()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mRawMaterialCT(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        FormRawMaterialCT.Show()
        Application.DoEvents()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub MenuItemOnClick_mCurrency(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        FormCurrency.Show()
        Application.DoEvents()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub MenuItemOnClick_mExit(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i As Integer = MsgBox("Are you sure?", MsgBoxStyle.OkCancel)
        If i = 1 Then
            For i = 1 To (My.Application.OpenForms.Count - 1)
                My.Application.OpenForms.Item(1).Close()
            Next
            fadeout(Me)
            Me.Close()
        End If
    End Sub
    Private Sub MenuItemOnClick_mHelpFile(ByVal sender As Object, ByVal e As System.EventArgs)
        'Process.Start("\\172.22.10.08\DJSoft\RawMaterial\help\User Guide - Raw Material Trend - Rev1.pdf")
        'Process.Start("C:\Program Files\DJSoft\RawMaterial\help\User Guide - Raw Material Trend.pdf")
        Process.Start(Application.StartupPath & "\help\User Guide - Raw Material Trend.pdf")
    End Sub
    Private Sub MenuItemOnClick_mTemplate(ByVal sender As Object, ByVal e As System.EventArgs)
        'Process.Start("C:\Program Files\DJSoft\RawMaterial\help\Material Trend - Template Common DB.xltx")
        Process.Start(Application.StartupPath & "\help\Material Trend - Template Common DB.xltx")
    End Sub
    Private Sub MenuItemOnClick_mAbout(ByVal sender As Object, ByVal e As System.EventArgs)
        FormAbout.Show()
    End Sub


    Protected Friend Sub setBubbleMessage(ByVal title As String, ByVal message As String)
        NotifyIcon1.Visible = True
        NotifyIcon1.BalloonTipText = message
        NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
        NotifyIcon1.BalloonTipTitle = title
        NotifyIcon1.ShowBalloonTip(200)
        ShowballonWindow(200)
    End Sub
    Private Sub ShowballonWindow(ByVal timeout As Integer)
        If timeout <= 0 Then
            Exit Sub
        End If
        Dim timeoutcount As Integer = 0
        While (timeoutcount < timeout)
            Thread.Sleep(1)
            timeoutcount += 1
        End While
        NotifyIcon1.Visible = False
    End Sub

End Class