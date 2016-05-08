Public Class Form1

    Dim user As String = ""
    Dim pass As String = ""

    Dim courseCode As String = ""
    Dim codes As ArrayList = Nothing
    Dim currentDex As Integer = -1

    Dim rnd As New Random
    Dim loopsLeft As Integer = rnd.Next(20, 26)

    Private Sub fillLogIn()
        Dim userBox As HtmlElement = Nothing
        Dim passBox As HtmlElement = Nothing
        Dim loginButton As HtmlElement = Nothing

        For Each h As HtmlElement In WebBrowser1.Document.All
            If h <> Nothing And h.Name <> Nothing Then
                If h.Name.Equals("mli") Then
                    userBox = h
                ElseIf h.Name.Equals("password") Then
                    passBox = h
                ElseIf h.Name.Equals("dologin") Then
                    loginButton = h
                End If
            End If
        Next

        If userBox <> Nothing And passBox <> Nothing And loginButton <> Nothing Then
            userBox.InnerText = user
            passBox.InnerText = pass
            loginButton.InvokeMember("click")
        End If
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Dim url As String = e.Url.ToString
        If url.Equals("https://passportyork.yorku.ca/ppylogin/ppylogin") And user <> Nothing Then
            fillLogIn()
        ElseIf url.Equals("https://wrem.sis.yorku.ca/Apps/WebObjects/REM.woa/wa/DirectAction/rem") Then
            Dim selectEle As HtmlElement = Nothing
            Dim continueEle As HtmlElement = Nothing
            For Each h As HtmlElement In WebBrowser1.Document.All
                If h <> Nothing And h.Name <> Nothing Then
                    If h.Name.Equals("3.5.1.27.1.11") Then
                        selectEle = h
                    ElseIf h.Name.Equals("3.5.1.27.1.13") Then
                        continueEle = h
                    End If
                End If
            Next
            If selectEle <> Nothing And continueEle <> Nothing Then
                selectEle.SetAttribute("value", "1")
                continueEle.InvokeMember("click")
            End If
        ElseIf url.Contains("/wo/") Then
            If WebBrowser1.DocumentTitle.Equals("REM - Summary") Then
                Threading.Thread.Sleep(rnd.Next(5000, 8001))
                Dim addEle As HtmlElement = Nothing
                For Each h As HtmlElement In WebBrowser1.Document.All
                    If h <> Nothing And h.Name <> Nothing Then
                        If h.Name.Equals("3.1.27.1.17") Then
                            addEle = h
                        End If
                    End If
                Next
                If addEle <> Nothing Then
                    addEle.InvokeMember("click")
                End If
            ElseIf WebBrowser1.DocumentText.Contains("Please key in the 6 digit catalogue number") Then
                Dim courseBox As HtmlElement = Nothing
                Dim addButton As HtmlElement = Nothing
                For Each h As HtmlElement In WebBrowser1.Document.All
                    If h <> Nothing And h.Name <> Nothing Then
                        If h.Name.Equals("3.1.27.5.7") Then
                            courseBox = h
                        ElseIf h.Name.Equals("3.1.27.5.9") Then
                            addButton = h
                        End If
                    End If
                Next
                If courseBox <> Nothing And addButton <> Nothing Then
                    If currentDex <> -1 Then
                        courseCode = codes.Item(currentDex)
                        If currentDex < (codes.Count - 1) Then
                            currentDex += 1
                        Else
                            currentDex = 0
                        End If
                    End If
                    courseBox.InnerText = courseCode
                    addButton.InvokeMember("click")
                End If
            ElseIf WebBrowser1.DocumentText.Contains("Please confirm that you want to:") Then
                Dim yesButton As HtmlElement = Nothing
                For Each h As HtmlElement In WebBrowser1.Document.All
                    If h <> Nothing And h.Name <> Nothing Then
                        If h.Name.Equals("3.1.27.7.9") Then
                            yesButton = h
                        End If
                    End If
                Next
                If yesButton <> Nothing Then
                    yesButton.InvokeMember("click")
                End If
            ElseIf WebBrowser1.DocumentText.Contains("Result:") Then
                If WebBrowser1.DocumentText.Contains("The course has not been added.") Then
                    Dim continueEle As HtmlElement = Nothing
                    For Each h As HtmlElement In WebBrowser1.Document.All
                        If h <> Nothing And h.Name <> Nothing Then
                            If h.Name.Equals("3.1.27.15.9") Then
                                continueEle = h
                            End If
                        End If
                    Next
                    If continueEle <> Nothing Then
                        loopsLeft -= 1
                        If loopsLeft = 0 Then
                            Dim logoutEle As HtmlElement = Nothing
                            For Each h As HtmlElement In WebBrowser1.Document.All
                                If h <> Nothing And h.Name <> Nothing Then
                                    If h.Name.Equals("3.1.3.1.1") Then
                                        logoutEle = h
                                    End If
                                End If
                            Next
                            If logoutEle <> Nothing Then
                                logoutEle.InvokeMember("click")
                            End If
                            loopsLeft = rnd.Next(3, 5)
                        Else
                            continueEle.InvokeMember("click")
                        End If
                    Else
                        Me.Text = Me.Text + " | " + courseCode
                        If currentDex <> -1 Then
                            codes.Remove(courseCode)
                            If codes.Count = 0 Then
                                MsgBox("Operation Complete")
                            End If
                        Else
                            MsgBox("Operation Complete")
                        End If
                    End If
                End If
            End If
        ElseIf url.Equals("https://passportyork.yorku.ca/ppylogin/ppylogout") Then
            WebBrowser1.Navigate("http://google.com")
        ElseIf url.Equals("http://www.google.ca/") Then
            Threading.Thread.Sleep(rnd.Next(60000, 120001))
            WebBrowser1.Navigate("https://wrem.sis.yorku.ca/Apps/WebObjects/REM")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        user = TextBox1.Text
        pass = TextBox2.Text
        courseCode = TextBox3.Text
        If courseCode.Contains(",") Then
            codes = New ArrayList(courseCode.Split(","))
            currentDex = 0
        End If
        Panel1.Dispose()
        fillLogIn()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
