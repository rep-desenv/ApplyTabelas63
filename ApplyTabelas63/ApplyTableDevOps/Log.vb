Imports System.IO
Imports System.Text
Imports System.IO.Directory

Module Log


    Public Sub GravaLogErro(ByVal Descricao As String)


        fs = New FileStream(sDirLog + sNomeArquivoErro, FileMode.Append)
        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
        mysw.WriteLine("[" & sUser & "] " & Now + " " + Descricao)
        mysw.Close()

    End Sub

    Public Sub GravaLog(ByVal Descricao As String)

        fs = New FileStream(sDirLog + sNomeArquivoLog, FileMode.Append)
        mysw = New StreamWriter(fs, System.Text.Encoding.Default)
        mysw.WriteLine("[" & sUser & "] " & Now + " " + Descricao)
        mysw.Close()

    End Sub

End Module
