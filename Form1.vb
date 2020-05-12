Public Class Form1
    Public SecondAfterExit As Integer = 5 'интвервал в секундах до скрытия уведомлений 
    Public CounterForExit As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
        Me.WindowState = FormWindowState.Minimized
        'get
        Dim isGetArgsFinded As Boolean = False
        If ArgsCount = 1 Or ArgsCount >= 2 Then
            If Args.Item(0).ToString.ToLower = "/get" Then
                isGetArgsFinded = True
                MsgBox(GetPVDServer())
            End If
        End If
        'set
        If ArgsCount >= 2 Then
            If ArgsCount >= 4 Then
                If Args.Item(2).ToString.ToLower = "/configfolder" Then
                    NewConfigFolder = Args.Item(3).ToString
                End If
            End If
            If Args.Item(0).ToString.ToLower = "/set" Then
                If FindRunnedPVD() = False Then
                    SetPVDServer(Args.Item(1).ToString)
                    SetStatus("Сервер ПК ПВД 2 переключён на ", Args.Item(1).ToString)
                Else
                    SetStatus("ПК ПВД 2", "Найден запущенный ПК ПВД 2, необходимо завершить с ним работу, для переключения сервера", ToolTipIcon.Warning)
                End If

            End If
        End If
        Try
            If Not (isGetArgsFinded) Then
                Dim EXE = GetPVDBinEXE()
                If EXE <> "" Then System.Diagnostics.Process.Start(EXE)
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        Timer1.Enabled = True

        'End
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        CounterForExit += 1
        If CounterForExit >= SecondAfterExit Then
            NotifyIcon1.Visible = False
            NotifyIcon1.Dispose()
            End
        End If
    End Sub
End Class
