Imports System.Runtime.InteropServices
Imports Notification.DesktopNotifications
Imports Notification.DesktopNotifications.NotificationActivator

<ClassInterface(ClassInterfaceType.None)>
<ComSourceInterfaces(GetType(INotificationActivationCallback))>
<Guid("16300E8D-3B63-4967-AAFA-3E93BB9CE7EE"), ComVisible(True)>
Public Class MyNotificationActivator
    Inherits NotificationActivator

    Public Overrides Sub OnActivated(arguments As String, userInput As NotificationUserInput, appUserModelId As String)
        '

    End Sub
End Class
