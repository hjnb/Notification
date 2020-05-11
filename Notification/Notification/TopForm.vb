Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Thread
Imports Windows.Data
Imports Windows.UI.Notifications

Public Class TopForm

    'データベースのパス
    Private dbFilePath As String = My.Application.Info.DirectoryPath & "\Notification.mdb"
    Private DB_Notification As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFilePath

    'トースト通知用xmlファイルパス
    Private xmlFilePath As String = My.Application.Info.DirectoryPath & "\toast.xml"

    Private firstFlg As Boolean = True

    Private notificationDt As New DataTable

    Private cnn As New ADODB.Connection

    'スレッドタイマ
    Private mtimer As System.Threading.Timer

    Private Delegate Sub LoadNotificationDelegate()

    ''' <summary>
    ''' loadイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TopForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'データベースファイルの存在チェック
        If Not System.IO.File.Exists(dbFilePath) Then
            MsgBox("データベースファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        If Not System.IO.File.Exists(xmlFilePath) Then
            MsgBox("xmlファイルが存在しません。ファイルを配置して下さい。")
            Me.Close()
            Exit Sub
        End If

        Me.MaximizeBox = False

        'データベース接続
        cnn.Open(DB_Notification)

        'タイマコールバック関数
        Dim timerDelegate As TimerCallback = New TimerCallback(AddressOf TimerEvent)
        'タイマ生成（コールバック関数の設定）
        mtimer = New Timer(timerDelegate, Nothing, 0, 10000)
    End Sub

    ''' <summary>
    ''' 処方箋通知データ読み込み
    ''' </summary>
    Private Sub LoadNotification()
        Dim rs As New ADODB.Recordset
        Dim sql As String = "select * from Noti where Flg = 1"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, rs, "Noti")
        Dim dt As DataTable = ds.Tables("Noti")

        If firstFlg Then '初回の処理
            '保持テーブルデータ更新
            notificationDt = dt.Copy()
            '通知処理
            NoticeNewData(notificationDt)
            firstFlg = False
        Else
            '取得データと前回データの比較
            Dim aaa = dt.AsEnumerable().Except(notificationDt.AsEnumerable(), DataRowComparer.Default)
            If aaa.Any() Then
                '差分が有る場合通知処理
                Dim diffDt As DataTable = aaa.CopyToDataTable()
                NoticeNewData(diffDt)
                notificationDt = diffDt
            End If
        End If
    End Sub

    ''' <summary>
    ''' 新規データをトースト通知
    ''' </summary>
    ''' <param name="dt">表示対象データテーブル</param>
    Private Sub NoticeNewData(dt As DataTable)
        For Each row As DataRow In dt.Rows
            Dim displayStr As String = row("Cod").ToString() & "様の処方箋通知"

            Dim xmlStr As String = File.ReadAllText(xmlFilePath, Encoding.GetEncoding("Shift_JIS"))
            Dim content As Xml.Dom.XmlDocument = New Xml.Dom.XmlDocument()
            content.LoadXml(xmlStr)
            content.GetElementsByTagName("text").First().AppendChild(content.CreateTextNode(displayStr))
            Dim notifier As ToastNotifier = ToastNotificationManager.CreateToastNotifier("Microsoft.Windows.Computer")
            Dim tn As New ToastNotification(content)

            notifier.Show(tn)
        Next
    End Sub

    ''' <summary>
    ''' System.Threading.Timer からの呼び出しを処理するメソッド
    ''' </summary>
    ''' <param name="state">このデリゲートで呼び出されたメソッドに関連するアプリケーション固有の情報を格納するオブジェクト</param>
    Private Sub TimerEvent(ByVal state As Object)
        '
        Invoke(New LoadNotificationDelegate(AddressOf LoadNotification), New Object() {})
    End Sub

    Private Sub test()
        ''テンプレート利用ver
        'Dim type = ToastTemplateType.ToastImageAndText01
        'Dim content = ToastNotificationManager.GetTemplateContent(type)
        'Dim text = content.GetElementsByTagName("text").First()
        'text.AppendChild(content.CreateTextNode("Toast"))
        'Dim notifier = ToastNotificationManager.CreateToastNotifier("Microsoft.Windows.Computer")
        'notifier.Show(New ToastNotification(content))

        ''自作xml利用ver
        'Dim xmlStr As String = File.ReadAllText(xmlFilePath, Encoding.GetEncoding("Shift_JIS"))
        'Dim content As Xml.Dom.XmlDocument = New Xml.Dom.XmlDocument()
        'content.LoadXml(xmlStr)
        'Dim notifier As ToastNotifier = ToastNotificationManager.CreateToastNotifier("Microsoft.Windows.Computer")
        'notifier.Show(New ToastNotification(content))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        NoticeNewData(notificationDt)
    End Sub
End Class
