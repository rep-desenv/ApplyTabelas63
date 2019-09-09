Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.IO
Imports System.Text
Imports System.IO.Directory

Module Apply_de_Tabelas

    Public sDirLog As String
    Public sNomeArquivoLog As String
    Public sNomeArquivoErro As String
    Public sUser As String

    Public fs As FileStream
    Public mysw As StreamWriter

    Public Const WM_COMMAND = &H111
    Public Const BN_CLICKED = 0
    Public Const BM_CLICK As Integer = &HF5

    Public Const WM_CLEAR = &H303

    Private Const WM_GETTEXT As Integer = &HD
    Public Const WM_SETTEXT As Long = &HC

    Public Const GW_CHILD As Long = 5

    Public Const WM_KEYDOWN As Long = &H100

    Private Const WM_DESTROY = &H2
    Private Const WM_CLOSE = &H10

    Public Const VK_RETURN As Long = &HD
    Private Const WM_GETTEXTLENGTH As Integer = &HE

    Public Const WM_CHAR As Int32 = &H102
    Private Const TVM_GETNEXTITEM = &H110A
    Private Const TVGN_CARET = 9
    Public Const LB_SETCURSEL = &H186
    Const LB_ERR = -1


    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function


    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer

    End Function

    Private Declare Function SendMessageByString Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, _
                                                                                        ByVal wParam As Integer, ByVal lParam As String) _
                                                                                    As Integer


    Declare Function SendMessageHM Lib "user32.dll" Alias "SendMessageA" ( _
            ByVal hWnd As IntPtr, _
            ByVal wMsg As Int32, _
            ByVal wParam As Int32, _
            ByVal lParam As String) As Int32

    Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As IntPtr) As IntPtr

    Public Declare Function GetMenu Lib "user32" Alias "GetMenu" (ByVal hwnd As Int32) As Int32
    Public Declare Function GetSubMenu Lib "user32" Alias "GetSubMenu" (ByVal hMenu As IntPtr, ByVal nPos As Int32) As IntPtr
    Private Declare Function GetMenuItemID Lib "user32" (ByVal MenuHandle As IntPtr, ByVal Index As IntPtr) As IntPtr

    Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer, _
                                                                          ByVal lpString As String, ByVal nMaxCount As Integer) As Integer

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Function GetDlgItem(
    ByVal hDlg As IntPtr,
    ByVal nIDDlgItem As Integer) As IntPtr
    End Function

    Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal uCmd As Integer) As IntPtr

    <DllImport("user32.dll")> _
    Private Function IsWindowEnabled(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Private Declare Auto Function PostMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, _
   ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr

    Private Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, _
    ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr
    <DllImport("user32.dll")> _
    Private Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    '<DllImport("user32", SetLastError:=True, CharSet:=CharSet.Auto)>
    'Private Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    'End Function

    Declare Auto Function FindWindow Lib "user32.dll" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    <DllImport("user32", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function FindWindowEx(ByVal hWndParent As IntPtr, ByVal hWndChildAfter As IntPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As IntPtr
    End Function
    Public Function MakeLong(ByVal lowPart As Short, ByVal highPart As Short) As Integer
        Return CInt(CUShort(lowPart) Or CUInt(highPart << 16))
    End Function

    <DllImport("user32")>
    Public Function GetDlgCtrlID(ByVal hWnd As IntPtr) As Integer
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> Public Function WritePrivateProfileString _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Function GetPrivateProfileString( _
      ByVal lpAppName As String, _
      ByVal lpKeyName As String, _
      ByVal lpDefault As String, _
      ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    End Function

    Public Function lerCFG(ByVal strIniFile As String, ByVal strKey As String, ByVal strItem As String) As String
        Dim strValue As StringBuilder = New StringBuilder(255)
        Dim intSize As Integer
        intSize = GetPrivateProfileString(strKey, strItem, "", strValue, 255, strIniFile)
        Return strValue.ToString
    End Function
    Private Function GetText(ByVal WindowHandle As Integer) As String
        Dim TextLen As Integer = SendMessage(WindowHandle, WM_GETTEXTLENGTH, 0, 0) + 1
        Dim Buffer As String = New String(" "c, TextLen)
        SendMessageByString(WindowHandle, WM_GETTEXT, TextLen, Buffer)
        Return Buffer
    End Function

    Sub Main(ByVal args() As String)

        Dim IniciaSiebel As Long
        Dim TotalArquivos As Integer = 0
        Dim TotalArquivosConflito As Integer = 0
        Dim TotalArquivosSemConflito As Integer = 0
        Dim TotalArquivosErro As Integer = 0

        Dim sToolsExe As String
        Dim sCFG As String

        Dim sPassword As String
        Dim sDataSource As String
        Dim sDirImp As String
        Dim Apply As Boolean
        Dim StartWindow As String
        Dim sArquivo As String
        Dim sPassword_Schema As String
        Dim NomeArquivo As String
        Dim HoraInicio As String
        Dim hWndTools As IntPtr
        Dim hWndSelect As IntPtr
        Dim hWndEdit As IntPtr
        Dim hWndBookmark As IntPtr
        Dim hWndLstEdit As IntPtr
        Dim hMenu As IntPtr
        Dim hSubMenu As IntPtr
        Dim hItem As IntPtr
        Dim PreencheuPassword As Boolean

        Dim Successfully As Boolean


        Dim B() As Byte
        Dim I As Integer
        Dim Tamanho As Long
        Dim SemExtensao As Long

        Dim sODBC As String

        HoraInicio = Now

        'Console.WriteLine(" ")
        'Console.WriteLine("Versao : 1.0")
        'Console.WriteLine(" ")
        'Console.WriteLine("Hora Inicio : " & HoraInicio)
        'Console.WriteLine(" ")

        '*****Fazendo a leitura dos parâmetros de inicialização*****'
        Try

            sNomeArquivoLog = "ApplyTabelas63.log"          'Especificando arquivo de log de execução'
            sNomeArquivoErro = "ErroApplyTabelas63.log"     'Especificando arquivo de log de erros'

            'GravaLog("Hora Início : " & HoraInicio)
            'Console.WriteLine("")
            'GravaLog("Versão: " & My.Application.Info.Version.ToString() & ")")
            'Console.WriteLine("")

            Console.WriteLine("Check Point: Recebendo parâmetros de execução.")

            'sToolsExe = args(0)
            'sCFG = args(1)
            'sUser = Trim$(args(2))
            'sPassword = Trim$(args(3))
            'sDirImp = Trim$(args(4))
            'sDirLog = Trim$(args(5))
            'sDataSource = Trim$(args(6))
            ''sODBC = Trim$(args(6))
            ''sPassword_Schema = Trim$(args(7))
            'sPassword_Schema = Trim$(args(7))

            Console.WriteLine("Check Point: Recebendo parâmetros de execução => OK")

            sToolsExe = "C:\sea630\tools\BIN\siebdev.exe"
            'sCFG = "c:\sea630\tools\bin\tools_B10.cfg"
            sCFG = "c:\sea630\tools\bin\tools_local.cfg"
            'sUser = "RPERES"
            sUser = "E_CARVALHO"
            'sPassword = "trv3keay"
            sPassword = "E_CARVALHO"
            sDirImp = "C:\Temp\imp"
            'sDataSource = "SERVER"
            sDataSource = "Server"
            'sPassword_Schema = "siebel"
            sPassword_Schema = "E_CARVALHO"

            'sDirLog = System.AppDomain.CurrentDomain.BaseDirectory

            Console.WriteLine("Versão: " & My.Application.Info.Version.ToString())
            Console.WriteLine("CFG :  " + sCFG)
            Console.WriteLine("UserName : " + sUser)
            Console.WriteLine("DataSource : " + sDataSource)
            Console.WriteLine("Dir Importacao : " + sDirImp)
            Console.WriteLine("Dir Log : " + sDirLog)
            'Console.WriteLine("sPassword_Schema : " + sPassword_Schema)

            '*****Escrevendo no arquivo de log informações de início da execução*****'
            fs = New FileStream(sDirLog + sNomeArquivoLog, FileMode.Append)
            mysw = New StreamWriter(fs, System.Text.Encoding.Default)

            mysw.WriteLine("[" & sUser & "] " & Now + " Hora Início : " & HoraInicio)
            mysw.WriteLine("[" & sUser & "] " & Now + " Versão: " & My.Application.Info.Version.ToString())
            mysw.WriteLine("[" & sUser & "] " & Now + " CFG :  " + sCFG)
            mysw.WriteLine("[" & sUser & "] " & Now + " UserName : " + sUser)
            mysw.WriteLine("[" & sUser & "] " & Now + " DataSource : " + sDataSource)
            mysw.WriteLine("[" & sUser & "] " & Now + " Dir Importacao : " + sDirImp)
            mysw.WriteLine("[" & sUser & "] " & Now + " Dir Log : " + sDirLog)
            'mysw.WriteLine("[" & sUser & "] " & Now + " Password ODBC : " + sPassword_Schema)
            mysw.Close()

            '*****Finalizando escrita no arquivo de log informações de início da execução*****'          

        Catch ex As Exception

            Console.WriteLine("Exception Parâmetro de entada :  " + Err.Description)
            GravaLogErro(" Exception Parâmetro de entada :  " + Err.Description)

            Exit Sub

        End Try

        Console.WriteLine("Check Point: Inciando Siebel Client.")
        '*****Inicializando Siebel utilizando os parâmetros de entrada da aplicação*****'

        IniciaSiebel = Shell(sToolsExe + " /c " + sCFG + " /U " + sUser + " /P " + sPassword + " /D " + sDataSource)

        GravaLog("Hora Inicio : " & Now)

        Thread.Sleep(2000)

        '****Ler o CFG do Siebel para obter o parâmetro ApplicationTitle*****'
        StartWindow = lerCFG(sCFG, "Siebel", "ApplicationTitle")

        '*****Localizando a janela do Siebel Tools*****'
        Do
            hWndTools = FindWindow(Nothing, StartWindow + " - Siebel Repository")
        Loop While hWndTools = 0

        Console.WriteLine("Check Point: Inciando Siebel Client => OK.")

        Thread.Sleep(1000)

        Console.WriteLine("Check Point: Localizando Arquivos 'TBL_*' Dentro do Diretório.")
        '*****Localizando os .sifs dentro do diretório e carregando no vetor*****'
        Dim fileEntries As String() = Directory.GetFiles(sDirImp, "*TBL_*")

        Console.WriteLine("Check Point: Localizando Arquivos 'TBL_*' Dentro do Diretório => OK.")

        For Each sFileName In fileEntries

            GravaLog("Iniciando Apply da tabela: " + System.IO.Path.GetFileName(sFileName) + " - Inicio: " + Now)
            Console.WriteLine("Iniciando Apply da tabela: " + System.IO.Path.GetFileName(sFileName) + " - Inicio: " + Now)

            NomeArquivo = ""
            sArquivo = ""
            NomeArquivo = System.IO.Path.GetFileName(sFileName)
            Tamanho = Len(NomeArquivo)
            SemExtensao = Tamanho - 8
            NomeArquivo = Mid(NomeArquivo, 5)
            NomeArquivo = Left(NomeArquivo, SemExtensao)

            My.Computer.Clipboard.SetText(NomeArquivo)

            'Next sFileName

            SetForegroundWindow(hWndTools)

            Console.WriteLine("Check Point: Navegando até Tela de Tabelas.")

            '*****Coloca ponteiro no sub menu de Bookmarks que vai para a tela de Tabelas*****'
            hMenu = GetMenu(hWndTools)
            hSubMenu = GetSubMenu(hMenu, 4)
            hItem = GetMenuItemID(hSubMenu, 1)
            PostMessage(hWndTools, WM_COMMAND, hItem, IntPtr.Zero)

            Console.WriteLine("Check Point: Navegando até Tela de Tabelas => OK.")

            '*****Localiza a janela de bookmarks*****'
            Console.WriteLine("Check Point: Executando Bookmarks.")
            Do

                hWndBookmark = FindWindow(Nothing, "Bookmarks")

                If hWndBookmark = 0 Then
                    Thread.Sleep(1000)
                End If
            Loop While hWndBookmark = 0

            Do

                SetForegroundWindow(hWndBookmark)

                '*****Localiza o ListBox na janela de bookmarks*****'
                hWndLstEdit = FindWindowEx(hWndBookmark, IntPtr.Zero, "ListBox", vbNullString)
                'SetForegroundWindow(hWndLstEdit)

                '*****Seleciona o primeiro item na janela de bookmarks*****'
                My.Computer.Keyboard.SendKeys("{DOWN}", True)
                'RetVal = SendMessage(hWndLstEdit, LB_SETCURSEL, 0, 0)
                'Thread.Sleep(2000)
                'SendMessage(hWndLstEdit, LB_SETCURSEL, 1, 0)

                '*****Clica no botão "Go to"******'
                My.Computer.Keyboard.SendKeys("%(G)", True)

                Console.WriteLine("Check Point: Executando Bookmarks => OK.")

            Loop While hWndBookmark = 0

            '*****Seleciona opção para executar uma nova Query*****'
            Console.WriteLine("Check Point: Executando Query de Tabela.")

            'My.Computer.Keyboard.SendKeys("%(Q)", True)
            hSubMenu = GetSubMenu(hMenu, 5)
            hItem = GetMenuItemID(hSubMenu, 0)
            PostMessage(hWndTools, WM_COMMAND, hItem, IntPtr.Zero)

            '*****Localiza a Data Table e insere o nome da tabela para pesquisa*****'

            'Dim B() As Byte = System.Text.Encoding.Default.GetBytes(sFileName)
            'Dim I As Integer
            'For I = 4 To sFileName.Length - 1
            'SendMessage(hWndEdit, WM_CHAR, B(I), 0)
            'Next I

            My.Computer.Keyboard.SendKeys("%(E)", True)
            Thread.Sleep(1000)
            My.Computer.Keyboard.SendKeys("P", True)
            Thread.Sleep(1000)
            My.Computer.Keyboard.SendKeys("%(Q)", True)
            Thread.Sleep(1000)
            My.Computer.Keyboard.SendKeys("X", True)
            Thread.Sleep(2000)

            Console.WriteLine("Check Point: Executando Query de Tabela => OK.")

            ' ''Paliativo para verificar se a tabela existe:
            ''hSubMenu = GetSubMenu(hMenu, 6)
            ''hItem = GetMenuItemID(hSubMenu, 0)
            ''PostMessage(hWndTools, WM_COMMAND, hItem, IntPtr.Zero)


            ''Dim hWndReport As IntPtr = FindWindow(Nothing, "Siebel Report Viewer")

            ''Do
            ''    hWndReport = FindWindow(Nothing, "Siebel Report Viewer")

            ''    If hWndReport = 0 Then
            ''        Thread.Sleep(1000)
            ''    End If
            ''Loop While hWndReport = 0


            ''Thread.Sleep(3000)

            ''Dim hWndReportMsg As IntPtr = FindWindow(Nothing, "Siebdev")

            ''For index As Integer = 1 To 3
            ''    hWndReportMsg = FindWindow(Nothing, "Siebdev")

            ''    If hWndReportMsg = 0 Then
            ''        Thread.Sleep(1000)
            ''    Else

            ''        Dim hWndOkButtonErroReport As IntPtr = FindWindowEx(hWndReportMsg, IntPtr.Zero, "Button", "OK")

            ''        If hWndOkButtonErroReport <> 0 Then

            ''            Do
            ''                hWndOkButtonErroReport = FindWindowEx(hWndReportMsg, IntPtr.Zero, "Button", "OK")

            ''                If hWndOkButtonErroReport = 0 Then
            ''                    Thread.Sleep(1000)
            ''                End If
            ''            Loop While hWndOkButtonErroReport = 0

            ''            SendMessage(hWndReportMsg, WM_COMMAND, MakeLong(GetDlgCtrlID(hWndOkButtonErroReport), BN_CLICKED), hWndOkButtonErroReport)
            ''            GravaLogErro("Tabela não encontrada: " + NomeArquivo + ";")
            ''            SendMessage(hWndReport, WM_CLOSE, 0, 0)
            ''            GoTo proximo

            ''        End If
            ''    End If
            ''Next

            ''SendMessage(hWndReport, WM_CLOSE, 0, 0)
            Console.WriteLine("Check Point: Digitando Password.")
            Thread.Sleep(100)
            Console.WriteLine("Check Point: Apertar Botão Apply.")
            My.Computer.Keyboard.SendKeys("%(l)", True)
            Thread.Sleep(1000)
            Console.WriteLine("Check Point: Apertar Botão Apply => OK.")


            My.Computer.Keyboard.SendKeys("{TAB}", True)
            Thread.Sleep(100)
            My.Computer.Keyboard.SendKeys("{TAB}", True)
            Thread.Sleep(100)
            My.Computer.Keyboard.SendKeys("{TAB}", True)
            Thread.Sleep(100)
            My.Computer.Keyboard.SendKeys("{TAB}", True)
            Thread.Sleep(100)
            My.Computer.Keyboard.SendKeys("{TAB}", True)
            Thread.Sleep(100)
            Console.WriteLine("Check Point: Navegou até a janela de senha.")
            My.Computer.Keyboard.SendKeys(sPassword_Schema, True)
            Thread.Sleep(100)



            'My.Computer.Keyboard.SendKeys("{TAB}", True)
            'Thread.Sleep(1000)
            'My.Computer.Keyboard.SendKeys(sODBC, True)
            'Thread.Sleep(1000)
            'My.Computer.Keyboard.SendKeys("%(E)", True)
            'Thread.Sleep(1000)
            'My.Computer.Keyboard.SendKeys("SIEBEL", True)
            'Thread.Sleep(1000)

            Console.WriteLine("Check Point: Digitando Password => OK.")

            Console.WriteLine("Check Point: Apertando Botão de Apply.")

            My.Computer.Keyboard.SendKeys("%(A)", True)
            'Thread.Sleep(1000)

            ''''''Escrever a variável da password'''''

            'Do

            '    hWndSelect = FindWindow(Nothing, "Apply Schema")

            '    If hWndSelect = 0 Then
            '        Thread.Sleep(1000)
            '    End If

            'Loop While hWndSelect = 0

            'hWndEdit = FindWindowEx(hWndSelect, IntPtr.Zero, "Edit", vbNullString)

            'SetForegroundWindow(hWndEdit)

            'SendMessage(hWndEdit, WM_CLEAR, hItem, IntPtr.Zero)

            'B = System.Text.Encoding.Default.GetBytes(sPassword_Schema)
            'For I = 0 To sPassword_Schema.Length - 1
            'SendMessage(hWndEdit, WM_CHAR, B(I), 0)
            'Next I
            ''''''Fim Escrever a password''''''

            'My.Computer.Keyboard.SendKeys("%(A)", True)
            'Thread.Sleep(1000)

            '' ''Do While Apply = True
            '' ''    Thread.Sleep(5000)

            '' ''    Dim hWndApply As IntPtr = FindWindow(Nothing, "Apply")

            '' ''    If hWndApply = 0 Then
            '' ''        Apply = False
            '' ''        Exit Do
            '' ''    End If

            '' ''Loop

            Dim hWndApply As IntPtr = FindWindow(Nothing, "Apply")

            Do

                hWndApply = FindWindow(Nothing, "Apply")

                If hWndApply = 0 Then
                    Thread.Sleep(1000)
                End If

            Loop While hWndApply = 1

            Console.WriteLine("Check Point: Apertando Botão de Apply => OK.")

            Console.WriteLine("Check Point: Aguardando Conclusão do Apply.")
            '***************************
            'Aguardando mensagem abaixo:
            '***************************
            '            Siebel()
            '---------------------------
            'Changes successfully applied.
            '---------------------------
            '            OK()

            'Thread.Sleep(2000)

            Dim hWndSuccessfully As IntPtr = FindWindow(Nothing, "Siebel")

            Do
                hWndSuccessfully = FindWindow(Nothing, "Siebel")

                If hWndSuccessfully = 0 Then
                    Thread.Sleep(1000)
                End If
            Loop While hWndSuccessfully = 0

            Dim hWndOkButtonErro As IntPtr = FindWindowEx(hWndSuccessfully, IntPtr.Zero, "Button", "OK")

            If hWndOkButtonErro <> 0 Then

                Do
                    hWndOkButtonErro = FindWindowEx(hWndSuccessfully, IntPtr.Zero, "Button", "OK")

                    If hWndOkButtonErro = 0 Then
                        Thread.Sleep(1000)
                    End If
                Loop While hWndOkButtonErro = 0

                SendMessage(hWndSuccessfully, WM_COMMAND, MakeLong(GetDlgCtrlID(hWndOkButtonErro), BN_CLICKED), hWndOkButtonErro)
                GravaLog("Apply aplicado com sucesso em: " + System.IO.Path.GetFileName(sFileName) + " - Fim: " + Now)
                Console.WriteLine("Apply aplicado com sucesso em: " + System.IO.Path.GetFileName(sFileName) + " - Fim: " + Now)
                Console.WriteLine("Check Point: Aguardando Conclusão do Apply => OK.")
            End If

            'Thread.Sleep(3000)
proximo:
            TotalArquivos = TotalArquivos + 1

        Next sFileName

        'Thread.Sleep(1000)

        SendMessage(hWndTools, WM_CLOSE, 0, 0)

        Console.WriteLine("")
        Console.WriteLine("Total de arquivos : " & TotalArquivos)
        Console.WriteLine("Hora Inicio : " & HoraInicio)
        Console.WriteLine("Hora Fim : " & Now)


        GravaLog("Hora Fim : " & Now)


        Thread.Sleep(1000)

fim:
        Try


        Catch ex As Exception


            GravaLogErro(" Erro Exception - " & ex.Message)
            GravaLog("Hora Fim : " & Now)

            Thread.Sleep(1000)

            SendMessage(hWndTools, WM_CLOSE, 0, 0)


            Console.WriteLine("Exception")
            Console.WriteLine(ex.Message)

            GravaLogErro(" Erro Exception - " & ex.Message)

            SendMessage(hWndTools, WM_CLOSE, 0, 0)

            Exit Sub


        End Try


    End Sub

End Module
