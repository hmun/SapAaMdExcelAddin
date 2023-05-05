Public Class SapAaMdAddIn

    Private Sub SapAaMdAddIn_Startup() Handles Me.Startup
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    Private Sub SapAaMdAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
