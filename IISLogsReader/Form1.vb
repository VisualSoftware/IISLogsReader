Imports System.IO

Public Class Form1

    Dim LastItem As Integer
    Dim LogsDir As String
    Dim GlobalImageListItemIcons As Integer
    Dim GeoIPLoaded As Boolean = False
    Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
    Dim ShellEXE As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListView1.Items.Clear()
        'load IIS logs
        Dim filesystemumad As New FileStream(MetroComboBox2.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
        Dim streamreaderumad As New StreamReader(filesystemumad)
        Dim line As String = streamreaderumad.ReadLine()

        While line = Nothing = False
            If Not line.StartsWith("#") Then
                Dim NewLogLine() As String = line.Split(" ")
                Dim NewLogItem As New ListViewItem
                NewLogItem.Text = ""
                NewLogItem.SubItems.Add(NewLogLine(0))
                NewLogItem.SubItems.Add(NewLogLine(1))
                NewLogItem.SubItems.Add(NewLogLine(2))
                NewLogItem.SubItems.Add(NewLogLine(3))
                NewLogItem.SubItems.Add(NewLogLine(4))
                NewLogItem.SubItems.Add(NewLogLine(5))
                NewLogItem.SubItems.Add(NewLogLine(6))
                NewLogItem.SubItems.Add(NewLogLine(7))
                NewLogItem.SubItems.Add(NewLogLine(8))
                NewLogItem.SubItems.Add(NewLogLine(9))
                NewLogItem.SubItems.Add(NewLogLine(10))
                NewLogItem.SubItems.Add(NewLogLine(11))
                NewLogItem.SubItems.Add(NewLogLine(12))
                NewLogItem.SubItems.Add(NewLogLine(13))
                NewLogItem.SubItems.Add(NewLogLine(14))
                If ComboBox1.SelectedIndex = 0 Then
                    'HTTP Response Icon
                    If NewLogLine(11).StartsWith("1") Then
                        'informational message
                        NewLogItem.ImageIndex = 4
                    ElseIf NewLogLine(11).StartsWith("2") Then
                        'success message
                        NewLogItem.ImageIndex = 0
                    ElseIf NewLogLine(11).StartsWith("3") Then
                        'redirection message
                        NewLogItem.ImageIndex = 5
                    ElseIf NewLogLine(11).StartsWith("4") Then
                        'client error message
                        NewLogItem.ImageIndex = 1
                    ElseIf NewLogLine(11).StartsWith("5") Then
                        'server error message
                        NewLogItem.ImageIndex = 2
                    Else
                        'unknown
                        NewLogItem.ImageIndex = 3
                    End If
                ElseIf ComboBox1.SelectedIndex = 1 Then
                    'Browser Agent Icon
                    If NewLogLine(9).Contains("rv:12") Then
                        'IE 12 / MS Edge [Spartan]
                        NewLogItem.ImageIndex = 6
                    ElseIf NewLogLine(9).Contains("rv:11") Then
                        'IE 11
                        NewLogItem.ImageIndex = 23
                    ElseIf NewLogLine(9).Contains("MSIE+10") Then
                        NewLogItem.ImageIndex = 23
                    ElseIf NewLogLine(9).Contains("MSIE+9") Then
                        NewLogItem.ImageIndex = 24
                    ElseIf NewLogLine(9).Contains("MSIE+8") Then
                        NewLogItem.ImageIndex = 25
                    ElseIf NewLogLine(9).Contains("MSIE+7") Then
                        NewLogItem.ImageIndex = 25
                    ElseIf NewLogLine(9).Contains("MSIE+6") Then
                        NewLogItem.ImageIndex = 26
                    ElseIf NewLogLine(9).Contains("Firefox") Then
                        NewLogItem.ImageIndex = 7
                    ElseIf NewLogLine(9).Contains("Chrome") Then
                        NewLogItem.ImageIndex = 8
                    ElseIf NewLogLine(9).Contains("FDM") Then
                        NewLogItem.ImageIndex = 28
                    ElseIf NewLogLine(9).Contains("Presto") Then
                        'Opera
                        NewLogItem.ImageIndex = 13
                    ElseIf NewLogLine(9).Contains("Opera") Then
                        NewLogItem.ImageIndex = 13
                    ElseIf NewLogLine(9).Contains("Safari") Then
                        If Not NewLogLine(9).Contains("Chrome") Then
                            NewLogItem.ImageIndex = 9
                        End If
                    Else
                        If NewLogLine(9).Contains("MSIE") Then
                            NewLogItem.ImageIndex = 27
                        Else
                            NewLogItem.ImageIndex = 3
                        End If
                    End If
                ElseIf ComboBox1.SelectedIndex = 2 Then
                    'OS
                    If NewLogLine(9).Contains("Windows+NT+6.3") Then
                        'Windows 8.1
                        NewLogItem.ImageIndex = 10
                    ElseIf NewLogLine(9).Contains("Windows+Phone+8") Then
                        NewLogItem.ImageIndex = 10
                    ElseIf NewLogLine(9).Contains("Windows+Phone+OS+7.8") Then
                        NewLogItem.ImageIndex = 21
                    ElseIf NewLogLine(9).Contains("Windows+Phone+OS+7.5") Then
                        NewLogItem.ImageIndex = 21
                    ElseIf NewLogLine(9).Contains("Windows+Phone+OS+7.0") Then
                        NewLogItem.ImageIndex = 22
                    ElseIf NewLogLine(9).Contains("Windows+NT+10.1") Then
                        'Windows Redstone
                        NewLogItem.ImageIndex = 20
                    ElseIf NewLogLine(9).Contains("Windows+NT+6.4") Then
                        'Windows 10 Alpha
                        NewLogItem.ImageIndex = 19
                    ElseIf NewLogLine(9).Contains("Windows+NT+10.0") Then
                        'Windows 10
                        NewLogItem.ImageIndex = 19
                    ElseIf NewLogLine(9).Contains("Windows+NT+6.2") Then
                        'Windows 8
                        NewLogItem.ImageIndex = 10
                    ElseIf NewLogLine(9).Contains("Windows+NT+6.1") Then
                        'Windows 7
                        NewLogItem.ImageIndex = 14
                    ElseIf NewLogLine(9).Contains("Windows+NT+6.0") Then
                        If NewLogLine(9).Contains("MSIE+6") Then
                            'Windows Longhorn
                            NewLogItem.ImageIndex = 17
                        Else
                            'Windows Vista
                            NewLogItem.ImageIndex = 15
                        End If
                    ElseIf NewLogLine(9).Contains("Windows+NT+5.2") Then
                        'Windows XP X64 / Server 2003
                        NewLogItem.ImageIndex = 16
                    ElseIf NewLogLine(9).Contains("Windows+NT+5.1") Then
                        'Windows XP
                        NewLogItem.ImageIndex = 16
                    ElseIf NewLogLine(9).Contains("Mac") Then
                        NewLogItem.ImageIndex = 11
                    ElseIf NewLogLine(9).Contains("iPod") Then
                        NewLogItem.ImageIndex = 11
                    ElseIf NewLogLine(9).Contains("iPhone") Then
                        NewLogItem.ImageIndex = 11
                    ElseIf NewLogLine(9).Contains("iPad") Then
                        NewLogItem.ImageIndex = 11
                    ElseIf NewLogLine(9).Contains("AppleTV") Then
                        NewLogItem.ImageIndex = 11
                    ElseIf NewLogLine(9).Contains("Linux") Then
                        NewLogItem.ImageIndex = 12
                    Else
                        If NewLogLine(9).Contains("Windows") Then
                            'Windows
                            NewLogItem.ImageIndex = 18
                        Else
                            NewLogItem.ImageIndex = 3
                        End If
                    End If
                Else
                    'IP to Country
                    If GeoIPLoaded = False Then
                        For Each line In IO.File.ReadAllLines(My.Application.Info.DirectoryPath & "\GeoIP\IP2LOCATION-LITE-DB1.CSV")
                            Dim Newentry As New ListViewItem
                            Dim newitem() As String = line.Replace("""", "").Split(",")
                            Newentry.Text = newitem(0)
                            Newentry.SubItems.Add(newitem(1))
                            Newentry.SubItems.Add(newitem(2))
                            Newentry.SubItems.Add(newitem(3))
                            ListView2.Items.Add(Newentry)
                        Next
                        GeoIPLoaded = True
                    End If
                    Try

                        Dim GetIP() As String = NewLogLine(8).Split(".")

                        Dim x As UInteger = GetIP(0) * 16777216
                        Dim y As UInteger = GetIP(1) * 65536
                        Dim z As UInteger = GetIP(2) * 256
                        Dim t As UInteger = x + y + z + GetIP(3)
                        For n As Integer = 0 To ListView2.Items.Count - 1
                            Dim selecteditm As ListViewItem = ListView2.Items(n)
                            Dim LowerBound As UInteger = selecteditm.Text
                            Dim UpperBound As UInteger = selecteditm.SubItems(1).Text
                            Dim NumberVal As UInteger = t
                            If (LowerBound <= NumberVal) AndAlso (NumberVal <= UpperBound) Then
                                Dim lvItem2 As ListViewItem = _ListView3.FindItemWithText(selecteditm.SubItems(2).Text, True, 0, True)
                                If (lvItem2 IsNot Nothing) Then
                                    NewLogItem.ImageIndex = lvItem2.SubItems(1).Text
                                Else
                                    ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\GeoIP\" & selecteditm.SubItems(2).Text & ".png"))
                                    GlobalImageListItemIcons += 1
                                    NewLogItem.ImageIndex = GlobalImageListItemIcons
                                    Dim NewCCAdded As New ListViewItem
                                    NewCCAdded.Text = selecteditm.SubItems(2).Text
                                    NewCCAdded.SubItems.Add(GlobalImageListItemIcons)
                                    ListView3.Items.Add(NewCCAdded)
                                End If
                                Exit For
                            Else

                            End If
                        Next

                    Catch ex As Exception
                        NewLogItem.ImageIndex = 3
                    End Try
                End If
                ListView1.Items.Add(NewLogItem)
            End If
            line = streamreaderumad.ReadLine()
        End While
        streamreaderumad.Close()
        filesystemumad.Close()
        If CheckBox1.Checked = True Then
            'autorefresh is on, scroll to last item
            ListView1.Items(ListView1.Items.Count - 1).EnsureVisible()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "IIS Logs Reader - v" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-OK.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-WARN.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ERR.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-UNK.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-INFO.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-move.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ie12.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ff.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-chrome.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-safari.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-win8.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-mac.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-linux.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-opera.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-7.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-vista.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-xp.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-lh.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-2000.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-9.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-10.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-wp75.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-wp70.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ie1011.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ie9.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ie78.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ie6.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ie45.png"))
        ImageList3.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\downloader.png"))
        ImageList4.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-http.png"))
        ImageList4.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-ie.png"))
        ImageList4.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\WR-win.png"))
        ImageList4.Images.Add(System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\Resources\euflag48b.png"))
        GlobalImageListItemIcons = ImageList3.Images.Count - 1
        ComboBox1.ImageList = ImageList4
        ComboBox1.Items.Add(New ComboBoxExItem("HTTP Status Code", 0))
        ComboBox1.Items.Add(New ComboBoxExItem("Web Browser", 1))
        ComboBox1.Items.Add(New ComboBoxExItem("Operating System", 2))
        ComboBox1.Items.Add(New ComboBoxExItem("Country Flag (IPv4)", 3))
        ComboBox1.SelectedIndex = 0
        If My.Application.CommandLineArgs.Count = 0 Then
            'No args
            If IO.File.Exists(My.Application.Info.DirectoryPath & "\Logsdir.dat") Then
                LogsDir = IO.File.ReadAllText(My.Application.Info.DirectoryPath & "\Logsdir.dat")
            Else
                LogsDir = "C:\inetpub\logs\LogFiles\"
            End If
            If IO.Directory.Exists(LogsDir) = False Then
                'Ask user for a new dir
                MetroFramework.MetroMessageBox.Show(Me, "The folder '" & LogsDir & "' does not exist. Please choose a new folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ChangeLogsFolderToolStripMenuItem_Click(sender, e)
            Else

            End If
            Try
                MetroComboBox1.Items.Clear()
                For Each item In IO.Directory.EnumerateDirectories(LogsDir)
                    MetroComboBox1.Items.Add(item)
                Next
                MetroComboBox1.SelectedIndex = 0
            Catch ex As Exception
                MetroFramework.MetroMessageBox.Show(Me, "Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            'Args
            For i As Integer = 0 To CommandLineArgs.Count - 1
                ShellEXE = ShellEXE & " " & CommandLineArgs(i)
            Next
            If ShellEXE.Contains("-rundevversion") Then
                If CommandLineArgs.Count = 1 Then
                    'enable features
                Else

                End If
            Else
                MetroComboBox1.Enabled = False
                MetroComboBox2.Items.Add(ShellEXE)
                MetroComboBox2.Items.Add("Other...")
                MetroComboBox2.SelectedIndex = 0
            End If
        End If



    End Sub

    Private Sub MetroComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MetroComboBox1.SelectedIndexChanged
        'ADD LOG FILES FROM SELECTED FOLDER
        MetroComboBox2.Items.Clear()
        Try
            Dim myFile As String
            Dim mydir As String = MetroComboBox1.Text
            For Each myFile In Directory.GetFiles(mydir, "*.log")
                MetroComboBox2.Items.Add(myFile)
            Next
            MetroComboBox2.Items.Add("Other...")
            MetroComboBox2.SelectedIndex = 0
        Catch ex As Exception
            MetroFramework.MetroMessageBox.Show(Me, "Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub MetroComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MetroComboBox2.SelectedIndexChanged
        If MetroComboBox2.Text = "Other..." Then
            'Select other file
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.Cancel Then
                MetroComboBox2.SelectedIndex = LastItem
            Else
                MetroComboBox2.Items.Add(OpenFileDialog1.FileName)
                MetroComboBox2.SelectedIndex = MetroComboBox2.Items.Count - 1
            End If
        Else
            LastItem = MetroComboBox2.SelectedIndex
            Button1_Click(sender, e)
        End If
    End Sub

    Private Sub MetroComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Button1_Click(sender, e)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Dim findstr As String = InputBox("Find what: ", "Find")
        If findstr = Nothing Then

        Else
            Button1_Click(sender, e)
            'Find
            For n As Integer = 0 To ListView1.Items.Count - 1
                Dim lvItem As ListViewItem = _ListView1.FindItemWithText(findstr, True, n, True)
                If (lvItem IsNot Nothing) Then

                    lvItem.BackColor = Color.Purple
                    lvItem.ForeColor = Color.White
                Else

                End If
            Next

        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Button1_Click(sender, e)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Button1_Click(sender, e)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Process.Start("https://visualsoftware.wordpress.com")
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Process.Start("https://www.twitter.com/VisualSoftCorp")
    End Sub

    Private Sub ChangeLogsFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeLogsFolderToolStripMenuItem.Click
        If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.Cancel Then

        Else
            LogsDir = FolderBrowserDialog1.SelectedPath & "\"
            IO.File.WriteAllText(My.Application.Info.DirectoryPath & "\Logsdir.dat", LogsDir)
            Try
                MetroComboBox1.Items.Clear()
                For Each item In IO.Directory.EnumerateDirectories(LogsDir)
                    MetroComboBox1.Items.Add(item)
                Next
                MetroComboBox1.SelectedIndex = 0
            Catch ex As Exception
                MetroFramework.MetroMessageBox.Show(Me, "Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening
        Try
            For Each item As ListViewItem In ListView1.SelectedItems
                If item.Selected = True Then
                    CopyURIToolStripMenuItem.Enabled = True
                    CopySourceIPToolStripMenuItem.Enabled = True
                    CopyBrowserAgentToolStripMenuItem.Enabled = True
                Else
                    CopyURIToolStripMenuItem.Enabled = False
                    CopySourceIPToolStripMenuItem.Enabled = False
                    CopyBrowserAgentToolStripMenuItem.Enabled = False
                End If
                Exit For
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CopyURIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyURIToolStripMenuItem.Click
        Try
            For Each item As ListViewItem In ListView1.SelectedItems
                If item.Selected = True Then
                    Clipboard.SetText(item.SubItems(5).Text)
                Else

                End If
                Exit For
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CopySourceIPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopySourceIPToolStripMenuItem.Click
        Try
            For Each item As ListViewItem In ListView1.SelectedItems
                If item.Selected = True Then
                    Clipboard.SetText(item.SubItems(9).Text)
                Else

                End If
                Exit For
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CopyBrowserAgentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyBrowserAgentToolStripMenuItem.Click
        Try
            For Each item As ListViewItem In ListView1.SelectedItems
                If item.Selected = True Then
                    Clipboard.SetText(item.SubItems(10).Text.Replace("+", " "))
                Else

                End If
                Exit For
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim CurrentFileIndex As Integer = MetroComboBox2.Items.Count
            Dim MyFiles() As String
            Dim i As Integer
            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                MetroComboBox2.Items.Add(MyFiles(i))
            Next
            MetroComboBox2.SelectedIndex = CurrentFileIndex
        End If
    End Sub
End Class

Class ComboBoxEx
    Inherits ComboBox
    Private m_imageList As ImageList
    Public Property ImageList() As ImageList
        Get
            Return m_imageList
        End Get
        Set(value As ImageList)
            m_imageList = value
        End Set
    End Property

    Public Sub New()
        DrawMode = DrawMode.OwnerDrawFixed
    End Sub

    Protected Overrides Sub OnDrawItem(ea As DrawItemEventArgs)
        Try
            ea.DrawBackground()
            ea.DrawFocusRectangle()
            Dim item As ComboBoxExItem
            Dim imageSize As Size
            imageSize.Height = 48
            imageSize.Width = 48
            Dim bounds As Rectangle = ea.Bounds

            Try
                item = DirectCast(Items(ea.Index), ComboBoxExItem)

                If item.ImageIndex <> -1 Then
                    m_imageList.Draw(ea.Graphics, bounds.Left, bounds.Top, item.ImageIndex)
                    ea.Graphics.DrawString(item.Text, ea.Font, New SolidBrush(ea.ForeColor), bounds.Left + imageSize.Width, bounds.Top)
                Else
                    ea.Graphics.DrawString(item.Text, ea.Font, New SolidBrush(ea.ForeColor), bounds.Left, bounds.Top)
                End If
            Catch
                If ea.Index <> -1 Then
                    ea.Graphics.DrawString(Items(ea.Index).ToString(), ea.Font, New SolidBrush(ea.ForeColor), bounds.Left, bounds.Top)
                Else
                    ea.Graphics.DrawString(Text, ea.Font, New SolidBrush(ea.ForeColor), bounds.Left, bounds.Top)
                End If
            End Try

            MyBase.OnDrawItem(ea)
        Catch ex As Exception

        End Try
    End Sub
End Class

Class ComboBoxExItem
    Private _text As String
    Public Property Text() As String
        Get
            Return _text
        End Get
        Set(value As String)
            _text = value
        End Set
    End Property

    Private _imageIndex As Integer
    Public Property ImageIndex() As Integer
        Get
            Return _imageIndex
        End Get
        Set(value As Integer)
            _imageIndex = value
        End Set
    End Property

    Public Sub New()
        Me.New("")
    End Sub

    Public Sub New(text As String)
        Me.New(text, -1)
    End Sub

    Public Sub New(text As String, imageIndex As Integer)
        _text = text
        _imageIndex = imageIndex
    End Sub

    Public Overrides Function ToString() As String
        Return _text
    End Function
End Class