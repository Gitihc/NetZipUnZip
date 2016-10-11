Imports System.Web

Imports System.IO
Imports Ionic.Zip
Imports ICSharpCode.SharpZipLib.Zip

Public Class Handler1
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        Dim flag As String = context.Request("flag")
        Select Case flag
            Case "1"
                IonicZip(context)
            Case "2"
                ICSharpZip(context)
        End Select

    End Sub


    Public Sub ICSharpZip(ByVal context As HttpContext)

        'ICSharpCode.SharpZipLib.Zip

        Dim FileBody() As Byte      'zip 文件
        Dim FileName As String      'zip 文件名称
        '从数据库取出二进制文件
        Dim dt As DataTable = DBUtil.ExecuteHasQuery(String.Format(" SELECT FileName,FileBody FROM UserFiles WHERE ID IN(1,2,3)"))
        '多文件 打包下载
        Using ms As MemoryStream = New MemoryStream
            Using zs As ICSharpCode.SharpZipLib.Zip.ZipOutputStream = New ICSharpCode.SharpZipLib.Zip.ZipOutputStream(ms)

                For Each row In dt.Rows
                    Dim fn As String = row("FileName")
                    If String.IsNullOrEmpty(FileName) Then
                        FileName = fn
                        FileName = Path.GetFileNameWithoutExtension(FileName)
                        FileName = String.Format(FileName & "等{0}个文件.zip", dt.Rows.Count)
                    End If
                    Dim fb() As Byte = row("FileBody")
                    Dim m As MemoryStream = New MemoryStream(fb)
                    Dim zn As ICSharpCode.SharpZipLib.Zip.ZipEntry = New ICSharpCode.SharpZipLib.Zip.ZipEntry(fn)
                    zs.PutNextEntry(zn)
                    zs.Write(fb, 0, fb.Length)
                Next
                zs.Close()
                FileBody = ms.ToArray
            End Using
        End Using

        context.Response.ContentType = "application/octet-stream"
        context.Response.Clear()
        context.Response.Buffer = True
        context.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(FileName))
        context.Response.BinaryWrite(FileBody)
        context.Response.Flush()
        HttpContext.Current.ApplicationInstance.CompleteRequest()
    End Sub


    Public Sub IonicZip(ByVal context As HttpContext)
        'Ionic.Zip
        Dim FileBody() As Byte      'zip 文件
        Dim FileName As String      'zip 文件名称
        '从数据库取出二进制文件
        Dim dt As DataTable = DBUtil.ExecuteHasQuery(String.Format(" SELECT FileName,FileBody FROM UserFiles WHERE ID IN(1,2,3)"))
        '多文件 打包下载
        Using ms As MemoryStream = New MemoryStream
            Using zs As Ionic.Zip.ZipOutputStream = New Ionic.Zip.ZipOutputStream(ms)

                zs.AlternateEncoding = System.Text.Encoding.GetEncoding("gb2312")
                zs.AlternateEncodingUsage = ZipOption.Always
                For Each row In dt.Rows
                    Dim fn As String = row("FileName")
                    If String.IsNullOrEmpty(FileName) Then
                        FileName = fn
                        FileName = Path.GetFileNameWithoutExtension(FileName)
                        FileName = String.Format(FileName & "等{0}个文件.zip", dt.Rows.Count)
                    End If
                    Dim fb() As Byte = row("FileBody")
                    Dim m As MemoryStream = New MemoryStream(fb)
                    zs.PutNextEntry(fn)
                    zs.Write(fb, 0, fb.Length)
                Next
                zs.Close()
                FileBody = ms.ToArray
            End Using
        End Using

        context.Response.ContentType = "application/octet-stream"
        context.Response.Clear()
        context.Response.Buffer = True
        context.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(FileName))
        context.Response.BinaryWrite(FileBody)
        context.Response.Flush()
        HttpContext.Current.ApplicationInstance.CompleteRequest()
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class