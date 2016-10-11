Imports System.IO
Imports System.IO.Packaging
Imports System.Web
Imports Ionic.Zip


Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        ProgressZip.Value = 0
    End Sub
    ''' <summary>
    ''' 压缩指定的文件夹
    ''' </summary>
    ''' <param name="folderName">待压缩文件夹</param>
    ''' <param name="zipFileName">压缩文件路径</param>
    ''' <param name="overrideExisting">是否覆盖已存在的文件</param>
    ''' <returns></returns>
    Private Function ZipFolder(ByVal folderName As String, ByVal zipFileName As String, ByVal overrideExisting As Boolean) As Boolean
        If (folderName.EndsWith("\")) Then
            folderName = folderName.Remove(folderName.Length - 1)
        End If
        Dim result As Boolean = False
        If (Not Directory.Exists(folderName)) Then
            Return result
        End If

        If (Not overrideExisting And File.Exists(zipFileName)) Then
            Return result
        End If
        If (overrideExisting And File.Exists(zipFileName)) Then
            File.Delete(zipFileName)
        End If
        Try
            Using zip As ZipFile = New ZipFile(zipFileName, System.Text.Encoding.GetEncoding("gb2312")) '支持中文
                AddHandler zip.SaveProgress, AddressOf MySaveProgress

                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.Default


                For Each fileName As String In Directory.EnumerateFiles(folderName, "*", SearchOption.AllDirectories)
                    zip.AddFile(fileName, "")
                Next
                zip.Save()
                zip.Dispose()
            End Using
            result = True
        Catch ex As Exception
            Throw New Exception("Error zipping folder " + folderName, ex)
        End Try
        Return result
    End Function
    ' <summary>
    ' 压缩指定的文件
    ' </summary>
    ' <param name="souFileName">源文件路径</param>
    ' <param name="zipFileName">压缩文件路径</param>
    ' <param name="overrideExisting">是否覆盖已存在的文件</param>
    ' <returns></returns>
    Private Function ZipFiles(ByVal souFileName As String, ByVal zipFileName As String, ByVal overrideExisting As Boolean) As Boolean
        Dim result As Boolean = False
        If (Not File.Exists(souFileName)) Then
            Return result
        End If

        If (Not overrideExisting And File.Exists(zipFileName)) Then
            Return result
        End If
        If (overrideExisting And File.Exists(zipFileName)) Then
            File.Delete(zipFileName)
        End If
        Try
            Using zip As ZipFile = New ZipFile(zipFileName, System.Text.Encoding.GetEncoding("gb2312")) '支持中文
                AddHandler zip.SaveProgress, AddressOf MySaveProgress
                zip.AddFile(souFileName, "")

                zip.Save()
                zip.Dispose()
            End Using
            result = True
        Catch ex As Exception
            Throw New Exception("Error zipping File " + souFileName, ex)
        End Try
        Return result
    End Function
    '''<summary>
    '''解压缩文件
    '''</summary>
    '''<param name="zipFileName">压缩文件路径</param>
    '''<param name="extDirectory">解压目标目录</param>
    '''<param name="sinFileName">指定解压文件名称</param>
    '''<param name="overrideExisting">是否覆盖已存在的文件</param>
    '''<returns></returns>
    Private Function UnZipFile(ByVal zipFileName As String, ByVal extDirectory As String, Optional ByVal sinFileName As String = "", Optional ByVal overrideExisting As Boolean = True) As Boolean
        Dim result As Boolean = False
        If Not File.Exists(zipFileName) Then Return result
        Try
            Dim options As New ReadOptions
            options.Encoding = System.Text.Encoding.GetEncoding("gb2312") '支持中文
            options.StatusMessageWriter = System.Console.Out
            Using zip As ZipFile = ZipFile.Read(zipFileName, options)
                Dim ExtractAction As ExtractExistingFileAction = IIf(overrideExisting, ExtractExistingFileAction.OverwriteSilently, ExtractExistingFileAction.DoNotOverwrite)
                AddHandler zip.ExtractProgress, AddressOf MyExtractProgress
                If sinFileName.Length = 0 Then
                    'For Each e As ZipEntry In zip '逐个文件解压（速度慢）
                    '    e.Extract(extDirectory, ExtractAction)
                    'Next
                    zip.ExtractAll(extDirectory, ExtractAction) '全部解压（速度快）
                Else
                    If zip.EntryFileNames.Contains(sinFileName) Then
                        zip(sinFileName).Extract(extDirectory, ExtractAction)
                    End If
                End If
                zip.Dispose()
            End Using
            result = True
        Catch ex As Exception
            Throw New Exception("Error unzipping file " + zipFileName, ex)
        End Try
        Return result
    End Function



 





    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ZipFolder("D:\Report", "D:\Report.zip", True)
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        UnZipFile("D:\PY.cos", "D:\Report1", , True)
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ZipFiles("D:\test.hlx", "D:\test.zip", True)
    End Sub
    Public Sub MySaveProgress(ByVal sender As Object, ByVal e As SaveProgressEventArgs)
        Dim Info As String = ""
        If (e.EventType = ZipProgressEventType.Saving_Started) Then
            Info = String.Format("正在保存: {0}……", e.ArchiveName)
        ElseIf (e.EventType = ZipProgressEventType.Saving_Completed) Then
            Info = String.Format("完成: {0}", e.ArchiveName)
        ElseIf (e.EventType = ZipProgressEventType.Saving_BeforeWriteEntry) Then
            Info = String.Format("  正在写入: {0} ({1}/{2})", e.CurrentEntry.FileName, e.EntriesSaved, e.EntriesTotal)
        ElseIf (e.EventType = ZipProgressEventType.Saving_EntryBytesRead) Then
            ProgressZip.Maximum = CInt(e.TotalBytesToTransfer)
            ProgressZip.Value = CInt(e.BytesTransferred)
            Info = String.Format("     {0}/{1} ({2:N0}%)", e.BytesTransferred, e.TotalBytesToTransfer, (CDbl(e.BytesTransferred) / (0.01 * e.TotalBytesToTransfer)))
        End If
        ListBox1.Items.Add(Info)
        ListBox1.TopIndex = ListBox1.Items.Count - 1
    End Sub

    Public Sub MyExtractProgress(ByVal sender As Object, ByVal e As ExtractProgressEventArgs)
        Dim Info As String = ""
        If (e.EventType = ZipProgressEventType.Extracting_EntryBytesWritten) Then
            ProgressZip.Maximum = CInt(e.TotalBytesToTransfer)
            ProgressZip.Value = CInt(e.BytesTransferred)
            Info = String.Format("   {0}/{1} ({2:N0}%)", e.BytesTransferred, e.TotalBytesToTransfer, (CDbl(e.BytesTransferred) / (0.01 * e.TotalBytesToTransfer)))
        ElseIf (e.EventType = ZipProgressEventType.Extracting_BeforeExtractEntry) Then
            Info = String.Format("解压缩: {0}", e.CurrentEntry.FileName)
        End If
        ListBox1.Items.Add(Info)
        ListBox1.TopIndex = ListBox1.Items.Count - 1
    End Sub

    
End Class
