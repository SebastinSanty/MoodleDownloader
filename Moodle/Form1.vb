Imports System.Net
Imports System.String
Imports System.Text.RegularExpressions
Imports HtmlAgilityPack

Public Class Form1
    Public completed As Boolean
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate("http://10.1.1.242/moodle/login/index.php")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        WebBrowser1.Document.GetElementById("username").SetAttribute("value", "f2015357")
        WebBrowser1.Document.GetElementById("password").SetAttribute("value", "Sebastian@16")
        WebBrowser1.Document.GetElementById("loginbtn").InvokeMember("click")

    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        'If WebBrowser1.Url = New Uri("http://10.1.1.242/moodle/my/") Then
        '    If WebBrowser1.ReadyState = WebBrowser.ReadyState Then
        '        Dim str As String = WebBrowser1.Document.GetElementById("course_id").InnerText

        '        Label2.Text = str
        '    End If
        'End If
        completed = True
    End Sub

    Private Sub WebBrowser1_Navigated(sender As Object, e As WebBrowserNavigatedEventArgs) Handles WebBrowser1.Navigated

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim str As String = WebBrowser1.Document.Body.InnerText
        Label2.Text = ""
        Dim ind As Integer = str.IndexOf("My coursesMy courses")
        Dim indi As Integer = str.IndexOf("All courses")
        Dim retstr As String = str.Substring(ind + 20, indi - ind - 20)
        Dim strarr() As String = retstr.Split(")")
        For Each s As String In strarr

        Next
        For i As Integer = 0 To strarr.Length - 2
            Dim newButton = New Button()
            newButton.Name = "btn" + i.ToString()
            newButton.Location = New Point(100, i * 35 + 100)
            newButton.Width = 300
            newButton.Text = (strarr(i).ToString() + ")")
            AddHandler newButton.Click, AddressOf dynabtn_Click
            Me.Controls.Add(newButton)
        Next
        ''''''''''''''''''
        Dim strh As String = WebBrowser1.Document.Body.InnerHtml
        Dim temp As String = ""
        For i As Integer = 0 To 3
            temp += strh(strh.IndexOf(strarr(2), strh.IndexOf(strarr(2)) + 1) - 6 + i)
        Next

        '''''''''''''''''''
    End Sub
    Private Sub dynabtn_Click(sender As Object, e As EventArgs)
        Dim find As String
        find = sender.Text
        Static strh As String = WebBrowser1.Document.Body.InnerHtml
        Dim temp As String

        temp = ""
        For i As Integer = 0 To 3
            temp += strh(strh.IndexOf(find, strh.IndexOf(find) + 1) - 6 + i)
        Next
        WebBrowser1.Navigate("http://10.1.1.242/moodle/course/view.php?id=" & temp)
        Label3.Text = WebBrowser1.Document.Body.InnerText
        Dim webClient As New System.Net.WebClient
        Dim WebSource As String = webClient.DownloadString("http://10.1.1.242/moodle/course/view.php?id=" & temp)
        Dim links As New List(Of String)()
        Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
        htmlDoc.LoadHtml(WebSource)

        For Each link As HtmlNode In htmlDoc.DocumentNode.SelectNodes("//a[starts-with(@href,'http: //10.1.1.242/moodle/Mod/resource/view.php?id=']")

            MsgBox(link.InnerText)

        Next


    End Sub
    Public Function ExtractLinks(ByVal url As String, ByVal txt As String) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("LinkText")
        dt.Columns.Add("LinkUrl")

        Dim wc As New WebClient
        Dim html As String = txt

        Dim links As MatchCollection = Regex.Matches(html, "<a.*?href=""(.*?)"".*?>(.*?)</a>")

        For Each match As Match In links
            Dim dr As DataRow = dt.NewRow
            Dim matchUrl As String = match.Groups(1).Value
            'Ignore all anchor links
            If matchUrl.StartsWith("#") Then
                Continue For
            End If
            'Ignore all javascript calls
            If matchUrl.ToLower.StartsWith("javascript:") Then
                Continue For
            End If
            'Ignore all email links
            If matchUrl.ToLower.StartsWith("mailto:") Then
                Continue For
            End If
            'For internal links, build the url mapped to the base address
            If Not matchUrl.StartsWith("http://") And Not matchUrl.StartsWith("https://") Then
                matchUrl = MapUrl(url, matchUrl)
            End If
            'Add the link data to datatable
            dr("LinkUrl") = matchUrl
            dr("LinkText") = match.Groups(2).Value
            dt.Rows.Add(dr)
        Next

        Return dt
    End Function

    Public Function MapUrl(ByVal baseAddress As String, ByVal relativePath As String) As String

        Dim u As New System.Uri(baseAddress)

        If relativePath = "./" Then
            relativePath = "/"
        End If

        If relativePath.StartsWith("/") Then
            Return u.Scheme + Uri.SchemeDelimiter + u.Authority + relativePath
        Else
            Dim pathAndQuery As String = u.AbsolutePath
            ' If the baseAddress contains a file name, like ..../Something.aspx
            ' Trim off the file name
            pathAndQuery = pathAndQuery.Split("?")(0).TrimEnd("/")
            If pathAndQuery.Split("/")(pathAndQuery.Split("/").Count - 1).Contains(".") Then
                pathAndQuery = pathAndQuery.Substring(0, pathAndQuery.LastIndexOf("/"))
            End If
            baseAddress = u.Scheme + Uri.SchemeDelimiter + u.Authority + pathAndQuery

            'If the relativePath contains ../ then
            ' adjust the baseAddress accordingly

            While relativePath.StartsWith("../")
                relativePath = relativePath.Substring(3)
                If baseAddress.LastIndexOf("/") > baseAddress.IndexOf("//" + 2) Then
                    baseAddress = baseAddress.Substring(0, baseAddress.LastIndexOf("/")).TrimEnd("/")
                End If
            End While

            Return baseAddress + "/" + relativePath
        End If

    End Function
End Class
