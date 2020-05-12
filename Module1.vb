Imports System.Xml


Module Module1
    Public ArgsCount As String = My.Application.CommandLineArgs.Count
    Public Args = My.Application.CommandLineArgs
    Public NewConfigFolder As String = ""

    ''' <summary>
    ''' Получение папки с актуальным конфигом
    ''' </summary>
    ''' <returns></returns>
    Public Function GetPVDConfigFolder() As String
        Dim retval As String = ""
        Try
            Dim DefaultFolder As String = Environ("appdata").ToString & "\..\Local\ООО_«ТЕКТУС.ИТ»\" '
            If NewConfigFolder <> "" Then
                retval = NewConfigFolder
            Else
                If FileIO.FileSystem.DirectoryExists(DefaultFolder) = True Then
                    Dim Intro1 = FileIO.FileSystem.GetDirectories(DefaultFolder)
                    retval = GetNewestConfigFolder(Intro1)
                End If
            End If
        Catch ex As Exception
        End Try
        Return retval
    End Function

    ''' <summary>
    ''' Сканирует общую директорию с несколькими конфигами и возвращает последнюю изменённую диреткорию (конфиг)
    ''' </summary>
    ''' <param name="Directories">Путь до общей директории</param>
    ''' <returns></returns>
    Public Function GetNewestConfigFolder(ByVal Directories As ObjectModel.ReadOnlyCollection(Of String)) As String
        Dim retval As String = "", tmprez As String = "", ArrayDate() As Date = {}, ArrayPathToPVDConfigFolder() As String = {}, CurrConfig As String = "", CurrConfigFolder As String = "", FullDirCount As Integer = 0, PVDVersion = ""
        Try
            PVDVersion = GetPVDVersion()
            If PVDVersion <> "" Then
                For Each Dir As String In Directories
                    CurrConfigFolder = Dir & "\" & PVDVersion
                    CurrConfig = CurrConfigFolder & "\user.config"
                    If FileIO.FileSystem.FileExists(CurrConfig) = True Then
                        FullDirCount += 1
                        ReDim Preserve ArrayDate(FullDirCount - 1)
                        ReDim Preserve ArrayPathToPVDConfigFolder(FullDirCount - 1)
                        ArrayDate(FullDirCount - 1) = FileIO.FileSystem.GetFileInfo(CurrConfig).LastWriteTime
                        ArrayPathToPVDConfigFolder(FullDirCount - 1) = CurrConfigFolder
                    End If
                Next
            End If
            Dim MaxDate = ArrayDate.Max().ToString
            For i = 0 To FullDirCount - 1
                If ArrayDate(i).ToString = MaxDate Then
                    retval = ArrayPathToPVDConfigFolder(i).ToString
                End If
            Next
        Catch ex As Exception
        End Try
        Return retval
    End Function


    Public Function GetPVDConfigFile() As String
        Dim retval As String = ""
        Try
            Dim PVDConfigFolder As String = GetPVDConfigFolder()
            If PVDConfigFolder <> "" Then retval = PVDConfigFolder & "\user.config" Else retval = ""
        Catch ex As Exception
        End Try
        Return retval
    End Function

    Public Function GetPVDBinEXE() As String
        Dim retval As String = ""
        Try
            Dim BinFolder = GetPVDBinFolder()
            If BinFolder <> "" Then
                retval = BinFolder & "SRCC.Terra.PVD.Reception.exe"
                If FileIO.FileSystem.FileExists(retval) = False Then retval = ""
            End If
        Catch ex As Exception
        End Try
        Return retval
    End Function

    ''' <summary>
    ''' Получение папки по умолчанию установленного ПК ПВД 2
    ''' </summary>
    ''' <returns></returns>
    Public Function GetPVDBinFolder() As String
        Dim retval As String = ""
        Try
            Dim DefaultFolder As String = Environ("appdata").ToString & "\..\Local\Programs\ПК ПВД (клиент)\Pvd\"
            If FileIO.FileSystem.DirectoryExists(DefaultFolder) = True Then retval = DefaultFolder Else retval = ""
        Catch ex As Exception
        End Try
        Return retval
    End Function

    ''' <summary>
    ''' Получение версии исполняемого файла ПК ПВД 2 из папки по умолчанию
    ''' </summary>
    ''' <returns></returns>
    Public Function GetPVDVersion() As String
        Dim retval As String = ""
        Try
            Dim BinFolder = GetPVDBinFolder()
            If BinFolder <> "" Then
                Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(BinFolder & "SRCC.Terra.PVD.Reception.exe")
                If myFileVersionInfo.FileVersion <> "" Then retval = myFileVersionInfo.FileVersion Else retval = ""
            End If
        Catch ex As Exception
        End Try
        Return retval
    End Function

    ''' <summary>
    ''' Получение текущего адреса сервера ПК ПВД 2
    ''' </summary>
    ''' <returns></returns>
    Public Function GetPVDServer() As String
        Dim retval As String = ""
        Try
            Dim IsFinded As Boolean = False
            Dim ConfigFile = GetPVDConfigFile()
            If ConfigFile <> "" And FileIO.FileSystem.FileExists(ConfigFile) = True Then
                Dim ConfigAsXML As New XmlDocument
                ConfigAsXML.Load(ConfigFile)
                Dim elevent As XmlNodeList = ConfigAsXML.GetElementsByTagName("SRCC.Terra.PVD.Common.Settings.ClientSettingsBase")
                For Each node As XmlNode In elevent
                    For Each node2 As XmlNode In node
                        If IsFinded = True Then Exit For
                        If node2.Name = "setting" Then
                            If node2.Attributes.Count <> 0 Then
                                For i = 0 To node2.Attributes.Count - 1
                                    If IsFinded = True Then Exit For
                                    If node2.Attributes.Item(i).Value.ToString = "SRCC.Terra.PVD.Windows.Tasks.Logon.PVDLogonDialog.Server" Then
                                        retval = node2.FirstChild.InnerText
                                        IsFinded = True
                                    End If
                                Next
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
        End Try
        Return retval
    End Function

    ''' <summary>
    ''' Установка адреса сервера ПК ПВД 2
    ''' </summary>
    ''' <param name="NewServer">Новый адреса сервера</param>
    ''' <returns></returns>
    Public Function SetPVDServer(ByVal NewServer As String)
        Dim retval As String = ""
        Try
            Dim IsFinded As Boolean = False
            Dim ConfigFile = GetPVDConfigFile()
            If ConfigFile <> "" And FileIO.FileSystem.FileExists(ConfigFile) = True Then
                Dim ConfigAsXML As New XmlDocument
                ConfigAsXML.Load(ConfigFile)
                Dim elevent As XmlNodeList = ConfigAsXML.GetElementsByTagName("SRCC.Terra.PVD.Common.Settings.ClientSettingsBase") '("SRCC.Terra.PVD.Windows.Tasks.Logon.PVDLogonDialog.Server")
                For Each node As XmlNode In elevent
                    For Each node2 As XmlNode In node
                        If IsFinded = True Then Exit For
                        If node2.Name = "setting" Then
                            If node2.Attributes.Count <> 0 Then
                                For i = 0 To node2.Attributes.Count - 1
                                    If IsFinded = True Then Exit For
                                    If node2.Attributes.Item(i).Value.ToString = "SRCC.Terra.PVD.Windows.Tasks.Logon.PVDLogonDialog.Server" Then
                                        node2.FirstChild.InnerText = NewServer
                                        IsFinded = True
                                    End If
                                Next
                            End If
                        End If
                    Next
                Next
                ConfigAsXML.Save(ConfigFile)
            End If
        Catch ex As Exception
        End Try
        Return retval
    End Function

    ''' <summary>
    ''' Поиск запущенного ПК ПВД 2
    ''' </summary>
    ''' <returns></returns>
    Public Function FindRunnedPVD() As Boolean
        Dim Retval As Boolean = False
        Try
            Dim Proc As Process() = Process.GetProcessesByName("SRCC.Terra.PVD.Reception")
            If Proc.Length <> 0 Then
                Retval = True
            End If
        Catch ex As Exception

        End Try
        Return Retval
    End Function

    ''' <summary>
    ''' Уведомление о событиях
    ''' </summary>
    ''' <param name="title">Заголовок</param>
    ''' <param name="text">Текст уведомления</param>
    Public Sub SetStatus(ByVal title As String, ByVal text As String, Optional CustomToolTipIcon As ToolTipIcon = ToolTipIcon.Info)
        With Form1.NotifyIcon1
            .Icon = Form1.Icon
            .BalloonTipIcon = CustomToolTipIcon
            .BalloonTipTitle = title
            .BalloonTipText = text
        End With
        Form1.NotifyIcon1.ShowBalloonTip(5000)
    End Sub


End Module