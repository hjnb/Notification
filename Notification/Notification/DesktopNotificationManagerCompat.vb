' ******************************************************************
' Copyright (c) Microsoft. All rights reserved.
' This code Is licensed under the MIT License (MIT).
' THE CODE Is PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS Or IMPLIED,
' INCLUDING BUT Not LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS Or COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
' DAMAGES Or OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
' TORT Or OTHERWISE, ARISING FROM, OUT OF Or IN CONNECTION WITH
' THE CODE Or THE USE Or OTHER DEALINGS IN THE CODE.
' ******************************************************************

' License for the RegisterActivator portion of code from FrecherxDachs

'The MIT License (MIT)

'Copyright(c) 2020 Michael Dietrich

'Permission Is hereby granted, free Of charge, to any person obtaining a copy
'of this software And associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, And/Or sell
'copies of the Software, And to permit persons to whom the Software Is
'furnished to do so, subject to the following conditions:

'The above copyright notice And this permission notice shall be included In
'all copies Or substantial portions Of the Software.

'THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or
'IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS Or COPYRIGHT HOLDERS BE LIABLE For ANY CLAIM, DAMAGES Or OTHER
'LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or OTHERWISE, ARISING FROM,
'OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS IN
'THE SOFTWARE.

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports Windows.UI.Notifications
Imports DesktopNotifications.NotificationActivator

Namespace DesktopNotifications
    Public Class DesktopNotificationManagerCompat
        Public Const TOAST_ACTIVATED_LAUNCH_ARG As String = "-ToastActivated"

        Private Shared _registeredAumidAndComServer As Boolean
        Private Shared _aumid As String
        Private Shared _registeredActivator As String

        Public Shared Sub RegisterAumidAndComServer(Of T As {NotificationActivator, New})(aumid As String)
            If (String.IsNullOrWhiteSpace(aumid)) Then
                Throw New ArgumentException("You must provide an AUMID.", NameOf(aumid))
            End If

            ' If running as Desktop Bridge
            If (DesktopBridgeHelpers.IsRunningAsUwp()) Then
                ' Clear the AUMID since Desktop Bridge doesn't use it, and then we're done.
                ' Desktop Bridge apps are registered with platform through their manifest.
                ' Their LocalServer32 key Is also registered through their manifest.
                _aumid = Nothing
                _registeredAumidAndComServer = True
                Return
            End If

            _aumid = aumid

            Dim exePath As String = Process.GetCurrentProcess().MainModule.FileName
            RegisterComServer(Of T)(exePath)

            _registeredAumidAndComServer = True
        End Sub

        Private Shared Sub RegisterComServer(Of T As {NotificationActivator, New})(exePath As String)
            ' We register the EXE to start up when the notification Is activated
            Dim regString As String = String.Format("SOFTWARE\\Classes\\CLSID\\{{{0}}}\\LocalServer32", GetType(T).GUID)
            Dim key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(regString)

            ' Include a flag so we know this was a toast activation And should wait for COM to process
            ' We also wrap EXE path in quotes for extra security
            key.SetValue(Nothing, """" & exePath & """" & " " & TOAST_ACTIVATED_LAUNCH_ARG)
        End Sub

        Public Shared Sub RegisterActivator(Of T As {NotificationActivator, New})()
            ' Big thanks to FrecherxDachs for figuring out the following code which works in .NET Core 3: https : //github.com/FrecherxDachs/UwpNotificationNetCoreTest
            Dim uuid = GetType(T).GUID
            Dim _cookie As UInt32
            CoRegisterClassObject(uuid, New NotificationActivatorClassFactory(Of T)(), CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE, _cookie)

            _registeredActivator = True
        End Sub

        <ComImport()>
        <Guid("00000001-0000-0000-C000-000000000046")>
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Private Interface IClassFactory

            <PreserveSig()>
            Function CreateInstance(pUnkOuter As IntPtr, ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer

            <PreserveSig()>
            Function LockServer(fLock As Boolean) As Integer
        End Interface

        Private Const CLASS_E_NOAGGREGATION As Integer = -2147221232
        Private Const E_NOINTERFACE As Integer = -2147467262
        Private Const CLSCTX_LOCAL_SERVER As Integer = 4
        Private Const REGCLS_MULTIPLEUSE As Integer = 1
        Private Const S_OK As Integer = 0
        Private Shared ReadOnly IUnknownGuid As Guid = New Guid("00000000-0000-0000-C000-000000000046")

        Private Class NotificationActivatorClassFactory(Of T As {NotificationActivator, New})
            Implements IClassFactory

            Public Function CreateInstance(pUnkOuter As IntPtr, ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer Implements IClassFactory.CreateInstance
                ppvObject = IntPtr.Zero

                If (pUnkOuter <> IntPtr.Zero) Then
                    Marshal.ThrowExceptionForHR(CLASS_E_NOAGGREGATION)
                End If


                If (riid = GetType(T).GUID OrElse riid = IUnknownGuid) Then
                    ' Create the instance of the .NET object
                    ppvObject = Marshal.GetComInterfaceForObject(New T(), GetType(NotificationActivator.INotificationActivationCallback))
                Else
                    ' The object that ppvObject points to does Not support the
                    ' interface identified by riid.
                    Marshal.ThrowExceptionForHR(E_NOINTERFACE)
                End If

                Return S_OK

            End Function

            Public Function LockServer(fLock As Boolean) As Integer Implements IClassFactory.LockServer
                Return S_OK
            End Function

        End Class

        <DllImport("ole32.dll")>
        Private Shared Function CoRegisterClassObject(<MarshalAs(UnmanagedType.LPStruct)> rclsid As Guid, <MarshalAs(UnmanagedType.IUnknown)> pUnk As Object, dwClsContext As UInt32, flags As UInt32, ByRef lpdwRegister As UInt32) As Integer
        End Function

        Public Shared Function CreateToastNotifier() As ToastNotifier
            EnsureRegistered()

            If (_aumid <> Nothing) Then
                ' Non-Desktop Bridge
                Return ToastNotificationManager.CreateToastNotifier(_aumid)
            Else
                ' Desktop Bridge
                Return ToastNotificationManager.CreateToastNotifier()
            End If
        End Function

        Public Shared ReadOnly Property History As DesktopNotificationHistoryCompat
            Get
                EnsureRegistered()

                Return New DesktopNotificationHistoryCompat(_aumid)
            End Get
        End Property

        Private Shared Sub EnsureRegistered()
            ' If Not registered AUMID yet
            If Not _registeredAumidAndComServer Then
                ' Check if Desktop Bridge
                If (DesktopBridgeHelpers.IsRunningAsUwp()) Then
                    ' Implicitly registered, all good!
                    _registeredAumidAndComServer = True
                Else
                    ' Otherwise, incorrect usage
                    Throw New Exception("You must call RegisterAumidAndComServer first.")
                End If
            End If

            ' If Not registered activator yet
            If Not _registeredActivator Then
                ' Incorrect usage
                Throw New Exception("You must call RegisterActivator first.")
            End If
        End Sub

        Public Shared ReadOnly Property CanUseHttpImages As Boolean
            Get
                Return DesktopBridgeHelpers.IsRunningAsUwp()
            End Get
        End Property

        ' <summary>
        ' Code from https://github.com/qmatteoq/DesktopBridgeHelpers/edit/master/DesktopBridge.Helpers/Helpers.cs
        ' </summary>
        Private Class DesktopBridgeHelpers
            Const APPMODEL_ERROR_NO_PACKAGE As Long = 15700L

            <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
            Shared Function GetCurrentPackageFullName(ByRef packageFullNameLength As Integer, packageFullName As StringBuilder) As Integer
            End Function

            Private Shared _isRunningAsUwp As Boolean?

            Public Shared Function IsRunningAsUwp() As Boolean
                If (_isRunningAsUwp Is Nothing) Then
                    If (IsWindows7OrLower) Then
                        _isRunningAsUwp = False
                    Else
                        Dim length As Integer = 0
                        Dim sb As StringBuilder = New StringBuilder(0)
                        Dim result As Integer = GetCurrentPackageFullName(length, sb)

                        sb = New StringBuilder(length)
                        result = GetCurrentPackageFullName(length, sb)

                        _isRunningAsUwp = (result <> APPMODEL_ERROR_NO_PACKAGE)
                    End If
                End If

                Return _isRunningAsUwp.Value
            End Function

            Private Shared ReadOnly Property IsWindows7OrLower As Boolean
                Get
                    Dim versionMajor As Integer = Environment.OSVersion.Version.Major
                    Dim versionMinor As Integer = Environment.OSVersion.Version.Minor
                    Dim Version As Double = versionMajor + CDbl(versionMinor) / 10
                    Return Version <= 6.1
                End Get
            End Property
        End Class
    End Class

    ' <summary>
    ' Manages the toast notifications for an app including the ability the clear all toast history And removing individual toasts.
    ' </summary>
    Public NotInheritable Class DesktopNotificationHistoryCompat
        Private _aumid As String
        Private _history As ToastNotificationHistory

        ' <summary>
        ' Do Not call this. Instead, call <see cref="DesktopNotificationManagerCompat.History"/> to obtain an instance.
        ' </summary>
        ' <param name="aumid"></param>
        Sub New(aumid As String)
            _aumid = aumid
            _history = ToastNotificationManager.History
        End Sub

        ' <summary>
        ' Removes all notifications sent by this app from action center.
        ' </summary>
        Public Sub Clear()
            If (_aumid <> Nothing) Then
                _history.Clear(_aumid)
            Else
                _history.Clear()
            End If
        End Sub

        ' <summary>
        ' Gets all notifications sent by this app that are currently still in Action Center.
        ' </summary>
        ' <returns>A collection of toasts.</returns>
        Public Function GetHistory() As IReadOnlyList(Of ToastNotification)
            Return If(_aumid <> Nothing, _history.GetHistory(_aumid), _history.GetHistory())
        End Function

        ' <summary>
        ' Removes an individual toast, with the specified tag label, from action center.
        ' </summary>
        ' <param name="tag">The tag label of the toast notification to be removed.</param>
        Public Sub Remove(tag As String)
            If (_aumid <> Nothing) Then
                _history.Remove(tag, String.Empty, _aumid)
            Else
                _history.Remove(tag)
            End If
        End Sub

        ' <summary>
        ' Removes a toast notification from the action using the notification's tag and group labels.
        ' </summary>
        ' <param name="tag">The tag label of the toast notification to be removed.</param>
        ' <param name="group">The group label of the toast notification to be removed.</param>
        Public Sub Remove(tag As String, group As String)
            If (_aumid <> Nothing) Then
                _history.Remove(tag, group, _aumid)
            Else
                _history.Remove(tag, group)
            End If
        End Sub

        ' <summary>
        ' Removes a group of toast notifications, identified by the specified group label, from action center.
        ' </summary>
        ' <param name="group">The group label of the toast notifications to be removed.</param>
        Public Sub RemoveGroup(group As String)
            If (_aumid <> Nothing) Then
                _history.RemoveGroup(group, _aumid)
            Else
                _history.RemoveGroup(group)
            End If
        End Sub
    End Class

    ' <summary>
    ' Apps must implement this activator to handle notification activation.
    ' </summary>NotificationActivator.INotificationActivationCallback
    Public MustInherit Class NotificationActivator
        Implements NotificationActivator.INotificationActivationCallback

        Public Sub Activate(appUserModelId As String, invokedArgs As String, data As NOTIFICATION_USER_INPUT_DATA(), dataCount As UInteger) Implements INotificationActivationCallback.Activate
            OnActivated(invokedArgs, New NotificationUserInput(data), appUserModelId)
        End Sub

        ' <summary>
        ' This method will be called when the user clicks on a foreground Or background activation on a toast. Parent app must implement this method.
        ' </summary>
        ' <param name="arguments">The arguments from the original notification. This Is either the launch argument if the user clicked the body of your toast, Or the arguments from a button on your toast.</param>
        ' <param name="userInput">Text And selection values that the user entered in your toast.</param>
        ' <param name="appUserModelId">Your AUMID.</param>
        Public MustOverride Sub OnActivated(arguments As String, userInput As NotificationUserInput, appUserModelId As String)

        ' These are the New APIs for Windows 10
        <StructLayout(LayoutKind.Sequential), Serializable()>
        Public Structure NOTIFICATION_USER_INPUT_DATA
            <MarshalAs(UnmanagedType.LPWStr)>
            Public Key As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Value As String
        End Structure

        <ComImport(), Guid("53E31837-6600-4A81-9395-75CFFE746F94"), ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Public Interface INotificationActivationCallback
            Sub Activate(
                <[In], MarshalAs(UnmanagedType.LPWStr)>
            appUserModelId As String,
                <[In], MarshalAs(UnmanagedType.LPWStr)>
            invokedArgs As String,
                <[In], MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=3)>
            data As NOTIFICATION_USER_INPUT_DATA(),
                <[In], MarshalAs(UnmanagedType.U4)>
            dataCount As UInteger)
        End Interface
    End Class

    ' <summary>
    ' Text And selection values that the user entered on your notification. The Key Is the ID of the input, And the Value Is what the user entered.
    ' </summary>
    Public Class NotificationUserInput
        Implements IReadOnlyDictionary(Of String, String)

        Private _data As NotificationActivator.NOTIFICATION_USER_INPUT_DATA()

        Sub New(Data As NotificationActivator.NOTIFICATION_USER_INPUT_DATA())
            _data = Data
        End Sub

        'Public String this[string key] => _data.First(i => i.Key == key).Value;
        Public ReadOnly Property Item(key As String) As String Implements IReadOnlyDictionary(Of String, String).Item
            Get
                Return _data.First(Function(i) i.Key = key).Value
            End Get
        End Property


        'Public IEnumerable<String> Keys => _data.Select(i => i.Key);
        Public ReadOnly Property Keys As IEnumerable(Of String) Implements IReadOnlyDictionary(Of String, String).Keys
            Get
                Return _data.Select(Function(i) i.Key)
            End Get
        End Property

        'Public IEnumerable<String> Values => _data.Select(i => i.Value);
        Public ReadOnly Property Values As IEnumerable(Of String) Implements IReadOnlyDictionary(Of String, String).Values
            Get
                Return _data.Select(Function(i) i.Value)
            End Get
        End Property

        'Public int Count => _data.Length;
        Public ReadOnly Property Count As Integer Implements IReadOnlyDictionary(Of String, String).Count
            Get
                Return _data.Length
            End Get
        End Property

        Public Function ContainsKey(key As String) As Boolean Implements IReadOnlyDictionary(Of String, String).ContainsKey
            Return _data.Any(Function(i) i.Key = key)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of String, String)) Implements IReadOnlyDictionary(Of String, String).GetEnumerator
            Return _data.Select(Function(i) New KeyValuePair(Of String, String)(i.Key, i.Value)).GetEnumerator()
        End Function

        Public Function TryGetValue(key As String, ByRef value As String) As Boolean Implements IReadOnlyDictionary(Of String, String).TryGetValue
            For Each a In _data
                If a.Key = key Then
                    value = a.Value
                    Return True
                End If
            Next

            value = Nothing
            Return False
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function
    End Class
End Namespace