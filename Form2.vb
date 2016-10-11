

Imports ICSharpCode.SharpZipLib.Zip
Imports System.IO

Public Class Form2



    Public Shared Sub ZipFiles(ByVal topDirectoryName As String, ByVal zipedFileName As String, ByVal compresssionLevel As Integer, ByVal password As String, ByVal comment As String)
        Using ms As MemoryStream = New MemoryStream


            'Using zos As ZipOutputStream = New ZipOutputStream(File.Open(zipedFileName, FileMode.OpenOrCreate))
            Using zos As ZipOutputStream = New ZipOutputStream(ms)
                If compresssionLevel <> 0 Then
                    zos.SetLevel(compresssionLevel) '设置压缩级别  
                End If

                If Not String.IsNullOrEmpty(password) Then
                    zos.Password = password '设置zip包加密密码  
                End If

                If Not String.IsNullOrEmpty(comment) Then
                    zos.SetComment(comment) '设置zip包的注释  
                End If

                '循环设置目录下所有的*.jpg文件（支持子目录搜索）  


                For Each f In Directory.GetFiles(topDirectoryName, "*.jpg", SearchOption.AllDirectories)
                    If File.Exists(f) Then
                        Dim item As FileInfo = New FileInfo(f)
                        Dim fs As FileStream = File.OpenRead(item.FullName)

                        Dim buffer() As Byte = New Byte(fs.Length) {}
                        fs.Read(buffer, 0, buffer.Length)

                        Dim enTry As ZipEntry = New ZipEntry(item.Name)
                        zos.PutNextEntry(enTry)
                        zos.Write(buffer, 0, buffer.Length)
                    End If
                Next
            End Using
            Dim bf() As Byte = ms.ToArray
            Dim ff = File.Create(zipedFileName, bf.Length)
            ff.Write(bf, 0, bf.Length)
            ff.Close()

        End Using
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            ZipFiles("D:\Report", "D:\Report.zip", 0, "", "")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

    End Sub
End Class