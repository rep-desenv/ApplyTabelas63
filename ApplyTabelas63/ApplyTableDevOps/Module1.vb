Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.IO
Imports System.Text
Imports System.IO.Directory

Module Module1

    Private Const TVGN_CHILD As Integer = &H4
    Private Const TV_FIRST As Integer = &H1100
    Private Const TVM_GETNEXTITEM As Integer = (TV_FIRST + 10)
    Private Const TVGN_NEXT As Integer = &H1
    Private Const TVGN_ROOT As Integer = &H0
    Private Const TVM_SELECTITEM As Integer = (TV_FIRST + 11)
    Private Const TVGN_CARET As Integer = &H9
    Private Const BM_CLICK As Integer = &HF5
    Public Const WM_KEYDOWN As Long = &H100
    Public Const VK_RETURN As Long = &HD
    Const WM_CHAR As Int32 = &H102
    Public Const WM_CLEAR = &H303
    Public Const WM_COMMAND = &H111
    Private Const WM_CLOSE = &H10
    Const WM_SETTEXT = &HC

    Public fs As FileStream
    Public mysw As StreamWriter

    Public sDirLog As String
    Public sNomeArquivoLog As String
    Public sNomeArquivoErro As String
    Public sUser As String


    Public Declare Function GetMenu Lib "user32" Alias "GetMenu" (ByVal hwnd As Int32) As Int32
    Public Declare Function GetSubMenu Lib "user32" Alias "GetSubMenu" (ByVal hMenu As IntPtr, ByVal nPos As Int32) As IntPtr
    Private Declare Function GetMenuItemID Lib "user32" (ByVal MenuHandle As IntPtr, ByVal Index As IntPtr) As IntPtr


    Declare Auto Function FindWindow Lib "user32.dll" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    ' Private Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, _
    'ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr


    Private Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, _
   ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr


    Private Declare Auto Function PostMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, _
   ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr

    Private Function GetRootNodeHandle(ByVal tvHandle As IntPtr, ByVal NodeIndex As Integer) As IntPtr
        Dim hNode As IntPtr = SendMessage(tvHandle, TVM_GETNEXTITEM, TVGN_ROOT, IntPtr.Zero) 'Get 1st root node handle
        'Iterate to the node wanted if (NodeIndex) is not not the 1st node (0) and get its handle
        For i As Integer = 1 To NodeIndex
            hNode = SendMessage(tvHandle, TVM_GETNEXTITEM, TVGN_NEXT, hNode)
            MsgBox(hNode.ToString)
        Next
        Return hNode 'Return the root node handle
    End Function

    Private Function GetChildNodeHandle(ByVal tvHandle As IntPtr, ByVal hParentNode As IntPtr, ByVal ChildNodeIndex As Integer) As IntPtr
        Dim hChildNode As IntPtr = SendMessage(tvHandle, TVM_GETNEXTITEM, TVGN_CHILD, hParentNode) 'Get the 1st child node
        'Iterate to the child node wanted if (ChildNodeIndex) is not not the 1st node (0) and get its handle
        For i As Integer = 1 To ChildNodeIndex
            hChildNode = SendMessage(tvHandle, TVM_GETNEXTITEM, TVGN_NEXT, hChildNode)
        Next
        Return hChildNode 'Return the child node handle
    End Function

    <DllImport("user32", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function FindWindowEx(ByVal hWndParent As IntPtr, ByVal hWndChildAfter As IntPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As IntPtr
    End Function

    <DllImport("user32.dll")> _
    Private Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> Public Function GetPrivateProfileString( _
      ByVal lpAppName As String, _
      ByVal lpKeyName As String, _
      ByVal lpDefault As String, _
      ByVal lpReturnedString As StringBuilder, _
      ByVal nSize As Integer, _
      ByVal lpFileName As String) As Integer
    End Function

    Public Function lerCFG(ByVal strIniFile As String, ByVal strKey As String, ByVal strItem As String) As String
        Dim strValue As StringBuilder = New StringBuilder(255)
        Dim intSize As Integer
        intSize = GetPrivateProfileString(strKey, strItem, "", strValue, 255, strIniFile)
        Return strValue.ToString
    End Function

    Sub Main(ByVal args() As String)

        Dim HoraInicio As String
        Dim IniciaSiebel As Long

        Dim sToolsExe As String
        Dim sCFG As String
        Dim sPassword As String
        Dim sDataSource As String
        Dim sDirImp As String
        Dim StartWindow As String
        Dim sPasswordApply As String

        Dim hWndTools As IntPtr


        Dim NomeArquivo As String
        Dim sArquivo As String
        Dim Contador As Integer = 0
        Dim retval

        '' Versão 4.6 - Alterado para usar SiebelApp
        '' Versão 4.6.2 - Alterado para Apagar WF
        '' Versão 4.6.3 - Alterado para Apagar Regra de negócio
        '' Versão 4.6.4 - Alterado para Apagar Obj. Repositorio Applet - Unique
        '' Versão 4.6.6 - Alterado para Apagar Consulta de Regra de negócio
        '' Versão 4.6.7 - Alterado para Apagar Aoes de Regra de negócio
        ''Versão 4.6.8 - Alterado para Apagar Condicoes de Regra de negócio - BUG
        ''Versão 4.7 -  Alterado para Apagar Acoes de Regra de negócio - BUG
        ''Versão 4.7.2 -  Alterado para Apagar VIEW Unique
        ''Versão 4.7.3 -  Alterado para Apagar BC e BS Unique
        ''Versão 4.7.4 -  Alterado para colocar EDM como overwrite
        ''Versão 4.7.5 -  Alterado para colocar Business Service Client

        ''Distribuições Diferentes - Inicio
        ''Versão 4.7.6 -  Alterado para Apagar Mapas de Valores EAI
        ''Versão 4.7.7 -  Ajuste para RGN
        ''Distribuições Diferentes - Fim

        ''Versão 4.7.8 -  Alterado para Apagar Mapas de Valores EAI + Ajuste de RGN



        HoraInicio = Now

        Console.WriteLine(" ")
        With My.Application.Info.Version
            Console.WriteLine("Versao Apply Table : " & .Major & "." & .Minor & " (Build " & .Build & "." & .Revision & ")")
        End With
        Console.WriteLine(" ")
        Console.WriteLine("Hora Inicio : " & HoraInicio)
        Console.WriteLine(" ")

        sNomeArquivoLog = "ApplyTableDevops.log"
        sNomeArquivoErro = "ErroApplyTableDevops.log"

        Try


            'sToolsExe = args(0)
            'sCFG = args(1)
            'sUser = Trim$(args(2))
            'sPassword = Trim$(args(3))
            'sDirImp = Trim$(args(4))
            'sDataSource = Trim$(args(5))
            'sPasswordApply = Trim$(args(6))


            sToolsExe = "c:\sea630\tools\BIN\siebdev.exe"
            sCFG = "c:\sea630\tools\bin\tools_local.cfg"
            'sCFG = "c:\sea630\tools\bin\tools_B10.cfg"
            sUser = "E_CARVALHO"
            'sUser = "RPERES"
            sPassword = "E_CARVALHO"
            'sPassword = "trv3keay"
            sDirImp = "c:\temp\imp\"
            sDataSource = "Local"
            'sDataSource = "Server"
            sPasswordApply = "E_CARVALHO"
            'sPasswordApply = "siebel"



            sDirLog = System.AppDomain.CurrentDomain.BaseDirectory


            Console.WriteLine("CFG :  " + sCFG)
            Console.WriteLine("UserName : " + sUser)
            Console.WriteLine("DataSource : " + sDataSource)
            Console.WriteLine("Dir Importacao : " + sDirImp)



            fs = New FileStream(sDirLog + sNomeArquivoLog, FileMode.Append)
            mysw = New StreamWriter(fs, System.Text.Encoding.Default)
            mysw.WriteLine("[" & sUser & "] " & Now + " CFG :  " + sCFG)
            mysw.WriteLine("[" & sUser & "] " & Now + " UserName : " + sUser)
            mysw.WriteLine("[" & sUser & "] " & Now + " DataSource : " + sDataSource)
            mysw.WriteLine("[" & sUser & "] " & Now + " Dir Importacao : " + sDirImp)
            mysw.Close()


        Catch ex As Exception

            Console.WriteLine("Exception Parâmetro de entada :  " + Err.Description)
            GravaLogErro(" Exception Parâmetro de entada :  " + Err.Description)

            Exit Sub

        End Try


        Try

            GravaLog("Hora Inicio : " & Now)

            Console.WriteLine("Inicio processo de Apply ...")
            Console.WriteLine("")

            IniciaSiebel = Shell(sToolsExe + " /c " + sCFG + " /U " + sUser + " /P " + sPassword + " /D " + sDataSource)


            Dim fileEntries As String() = Directory.GetFiles(sDirImp, "TBL_*.sif")


            StartWindow = lerCFG(sCFG, "Siebel", "ApplicationTitle")

            Do
                hWndTools = FindWindow(Nothing, StartWindow + " - Siebel Repository")
            Loop While hWndTools = 0

            For Each sFileName In fileEntries

                Dim Warning As IntPtr
                Dim BtnOk As IntPtr
                Dim BtnApply As IntPtr
                Dim ApplySchema As IntPtr

                Dim hWndWizardSiebeldev As IntPtr
                Dim hWndWizardSiebeldevError As IntPtr
                Dim hWndWizardSiebel As IntPtr

                Dim BtnOkoDBC As IntPtr
                Dim Edit5 As IntPtr
                Dim Edit6 As IntPtr


                NomeArquivo = ""
                sArquivo = ""
                NomeArquivo = System.IO.Path.GetFileName(sFileName)

                ' Console.WriteLine(" NomeArquivo : " & NomeArquivo)

                NomeArquivo = NomeArquivo.Replace("TBL_", "")
                NomeArquivo = Trim(NomeArquivo.Replace(".sif", ""))

                'Console.WriteLine("Novo NomeArquivo : " & NomeArquivo)
                'Console.WriteLine("Tamanho NomeArquivo : " & NomeArquivo.Length.ToString)

                Console.WriteLine("Tabela: " & NomeArquivo)





                Thread.Sleep(1000)

                SetForegroundWindow(hWndTools)

                Dim hWndTable2 As IntPtr = FindWindowEx(hWndTools, IntPtr.Zero, "AfxControlBar42", "Object Explorer")
                Dim hWndTable3 As IntPtr = FindWindowEx(hWndTable2, IntPtr.Zero, "#32770", "Object Explorer")
                Dim hWndTable4 As IntPtr = FindWindowEx(hWndTable3, IntPtr.Zero, "SysTreeView32", "Tree1")


                If Contador = 0 Then

                    Dim hRootNode As IntPtr = GetRootNodeHandle(hWndTable4, 0)
                    Dim hChildNode As IntPtr = GetChildNodeHandle(hWndTable4, hRootNode, 29)


                    If hChildNode <> IntPtr.Zero Then
                        'SendMessage(hWndTable4, TVM_EXPAND, TVE_EXPAND, hChildNode) 'Expand the node
                        SendMessage(hWndTable4, TVM_SELECTITEM, TVGN_CARET, hChildNode)
                    End If
                Else
                    Dim hRootNodi As IntPtr = GetRootNodeHandle(hWndTable4, 0)
                    Dim hChildNodi As IntPtr = GetChildNodeHandle(hWndTable4, hRootNodi, 0)


                    If hChildNodi <> IntPtr.Zero Then
                        SendMessage(hWndTable4, TVM_SELECTITEM, TVGN_CARET, hChildNodi)
                    End If

                    Dim hRootNodi1 As IntPtr = GetRootNodeHandle(hWndTable4, 0)
                    Dim hChildNodi1 As IntPtr = GetChildNodeHandle(hWndTable4, hRootNodi1, 29)


                    If hChildNodi1 <> IntPtr.Zero Then
                        SendMessage(hWndTable4, TVM_SELECTITEM, TVGN_CARET, hChildNodi1)
                    End If

                End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Thread.Sleep(2000)

                Dim hMenu As IntPtr = GetMenu(hWndTools)
                Dim hSubMenu As IntPtr = GetSubMenu(hMenu, 5)
                Dim hItem As IntPtr = GetMenuItemID(hSubMenu, 0)
                PostMessage(hWndTools, WM_COMMAND, hItem, IntPtr.Zero)

                Thread.Sleep(1000)

                Dim hWndJanela As IntPtr = FindWindow(Nothing, StartWindow + " - Siebel Repository - Table List")
                Dim hWndMDI As IntPtr = FindWindowEx(hWndJanela, IntPtr.Zero, "MDIClient", Nothing)
                Dim hWndTables As IntPtr = FindWindowEx(hWndMDI, IntPtr.Zero, Nothing, "Tables")
                Dim hWndDialogo As IntPtr = FindWindowEx(hWndTables, IntPtr.Zero, "#32770", Nothing)
                Thread.Sleep(1000)
                Dim teste As IntPtr = FindWindowEx(hWndDialogo, IntPtr.Zero, "AfxWnd42", Nothing)
                Dim hWndEdit As IntPtr = FindWindowEx(teste, IntPtr.Zero, "Edit", Nothing)

                SetForegroundWindow(teste)



                SendMessage(hWndEdit, WM_CLEAR, hItem, IntPtr.Zero)

                Thread.Sleep(1000)

                Dim B() As Byte = System.Text.Encoding.Default.GetBytes(NomeArquivo)
                Dim I As Integer
                For I = 0 To NomeArquivo.Length - 1
                    SendMessage(hWndEdit, WM_CHAR, B(I), 0)
                Next I

                Dim hMenu1 As IntPtr = GetMenu(hWndTools)
                Dim hSubMenu1 As IntPtr = GetSubMenu(hMenu1, 5)
                Dim hItem1 As IntPtr = GetMenuItemID(hSubMenu1, 2)
                PostMessage(hWndTools, WM_COMMAND, hItem1, IntPtr.Zero)


                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Thread.Sleep(1000)

                Dim SiebelTools As IntPtr = FindWindow(Nothing, StartWindow + " - Siebel Repository - Table List")
                Dim Table As IntPtr = FindWindowEx(SiebelTools, IntPtr.Zero, "MDIClient", Nothing)
                Dim Table1 As IntPtr = FindWindowEx(Table, IntPtr.Zero, "AfxFrameOrView42", Nothing)
                Dim Table2 As IntPtr = FindWindowEx(Table1, IntPtr.Zero, "#32770", Nothing)
                Dim Table3 As IntPtr = FindWindowEx(Table2, IntPtr.Zero, "Button", "Apply")

                Dim Aplicar As IntPtr = FindWindowEx(Table2, IntPtr.Zero, "Button", "Aplicar")


                'SetForegroundWindow(Table3)


                If Table3 <> 0 Then
                    retval = PostMessage(Table3, BM_CLICK, BM_CLICK, 0)
                ElseIf Aplicar <> 0 Then
                    retval = PostMessage(Aplicar, BM_CLICK, BM_CLICK, 0)
                End If

                ' Dim retval = PostMessage(Table3, BM_CLICK, BM_CLICK, 0)
                ' Dim retval1 = PostMessage(Table3, BM_CLICK, BM_CLICK, 0)


                If sDataSource.ToUpper = "LOCAL" Then
                    Do
                        Warning = FindWindowEx(0, IntPtr.Zero, "#32770", "Warning")

                        If Warning <> 0 Then
                            Thread.Sleep(1000)
                            PostMessage(Warning, WM_KEYDOWN, VK_RETURN, 0)
                            'Dim retval2 = PostMessage(BtnOk, BM_CLICK, BM_CLICK, 0)
                        End If
                    Loop While Warning = 0
                End If

                Thread.Sleep(1000)

                ApplySchema = FindWindowEx(vbNullString, IntPtr.Zero, "#32770", "Apply Schema")
                SetForegroundWindow(ApplySchema)


                Dim Edit0 As IntPtr = FindWindowEx(ApplySchema, IntPtr.Zero, "Edit", Nothing)
                Dim Edit1 As IntPtr = FindWindowEx(ApplySchema, Edit0, "Edit", Nothing)
                Dim Edit2 As IntPtr = FindWindowEx(ApplySchema, Edit1, "Edit", Nothing)
                Dim Edit3 As IntPtr = FindWindowEx(ApplySchema, Edit2, "Edit", Nothing)
                Dim Edit4 As IntPtr = FindWindowEx(ApplySchema, Edit3, "Edit", Nothing)

                SendMessage(Edit4, WM_SETTEXT, IntPtr.Zero, vbNullString)



                Dim B2() As Byte = System.Text.Encoding.Default.GetBytes(sPasswordApply)
                Dim I2 As Integer
                For I2 = 0 To sPasswordApply.Length - 1
                    SendMessage(Edit4, WM_CHAR, B2(I2), 0)
                Next I2


                BtnApply = FindWindowEx(ApplySchema, IntPtr.Zero, "Button", Nothing)

                Dim retval4 = PostMessage(BtnApply, BM_CLICK, BM_CLICK, 0)
                Dim retval5 = PostMessage(BtnApply, BM_CLICK, BM_CLICK, 0)

                Do
                    Thread.Sleep(2000)

                    hWndWizardSiebel = FindWindowEx(0, IntPtr.Zero, "#32770", "Siebel")
                    If hWndWizardSiebel <> 0 Then
                        Thread.Sleep(1000)
                        PostMessage(hWndWizardSiebel, WM_KEYDOWN, VK_RETURN, 0)
                        Console.WriteLine("Apply realizado com sucesso")
                        Console.WriteLine(" ")
                    End If



                    hWndWizardSiebeldev = FindWindowEx(0, IntPtr.Zero, "#32770", "siebdev")
                    If hWndWizardSiebeldev <> 0 Then
                        Thread.Sleep(1000)
                        BtnOkoDBC = FindWindowEx(hWndWizardSiebeldev, IntPtr.Zero, "Button", "OK")
                        Dim retval6 = PostMessage(BtnOkoDBC, BM_CLICK, BM_CLICK, 0)
                        hWndWizardSiebel = 0
                        Console.WriteLine("Erro: ODBC")
                        GravaLogErro("Erro: ODBC")
                        Thread.Sleep(1000)

                        hWndWizardSiebeldevError = FindWindowEx(0, IntPtr.Zero, "#32770", "Error")

                        If hWndWizardSiebeldevError <> 0 Then
                            BtnOkoDBC = FindWindowEx(hWndWizardSiebeldevError, IntPtr.Zero, "Button", "OK")
                            Dim retval7 = PostMessage(BtnOkoDBC, BM_CLICK, BM_CLICK, 0)
                            hWndWizardSiebel = 0
                            GoTo Finish
                        End If

                        GoTo Finish
                    End If




                Loop While hWndWizardSiebel = 0

                Contador = Contador + 1

                ' SendMessage(hWndTools, WM_CLOSE, 0, 0)

            Next

Finish:


            Console.WriteLine("Total de Tabelas : " & Contador)
            Console.WriteLine(" ")

            SendMessage(hWndTools, WM_CLOSE, 0, 0)

            GravaLog("Hora Fim : " & Now)
            GravaLog("Total de Tabelas : " & Contador)

        Catch ex As Exception

            Console.WriteLine("Erro Exception : " & ex.Message)
            Console.WriteLine("")
            GravaLogErro(" Erro Exception :  " + ex.Message)

        End Try



    End Sub

End Module
