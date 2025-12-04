<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="rrr.aspx.vb" Inherits="MgmSP.rrr" %>

<!DOCTYPE html>

<script runat="server">

    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
        Const OUTPUT_FILENAME As String = “renderedpage.html”
        Dim renderedOutput As StringBuilder = Nothing
        Dim strWriter As IO.StringWriter = Nothing
        Dim tWriter As HtmlTextWriter = Nothing
        Dim outputStream As IO.FileStream = Nothing
        Dim sWriter As IO.StreamWriter = Nothing
        Dim filename As String
        Dim nextPage As String

        Try
            'create a HtmlTextWriter to use for rendering the page
            renderedOutput = New StringBuilder
            strWriter = New IO.StringWriter(renderedOutput)
            tWriter = New HtmlTextWriter(strWriter)

            MyBase.Render(tWriter)

            'save the rendered output to a file
            'filename = Server.MapPath(“.”) & “\” & OUTPUT_FILENAME
            filename = "c:\data\" & OUTPUT_FILENAME
            outputStream = New IO.FileStream(filename,
                                  IO.FileMode.Create)
            sWriter = New IO.StreamWriter(outputStream)
            sWriter.Write(renderedOutput.ToString())
            sWriter.Flush()

            ' redirect to another page
            '  NOTE: Continuing with the display of this page will result in the
            '       page being rendered a second time which will cause an exception
            '        to be thrown
            nextPage = “DisplayMessage.aspx?” &
                       “PageHeader=Information” & “&” &
                       “Message1=HTML Output Saved To ” & OUTPUT_FILENAME
            'Response.Redirect(nextPage)
            'Response.Write(renderedOutput.ToString())

            writer.Write(renderedOutput.ToString())
        Finally

            'clean up
            If (Not IsNothing(outputStream)) Then
                outputStream.Close()
            End If

            If (Not IsNothing(tWriter)) Then
                tWriter.Close()
            End If

            If (Not IsNothing(strWriter)) Then
                strWriter.Close()
            End If
        End Try
    End Sub

</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Capture Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:dropdownlist id="DropDownList1" runat="server">
            <asp:listitem>red</asp:listitem>
            <asp:listitem>blue</asp:listitem>
            <asp:listitem>green</asp:listitem>
        </asp:dropdownlist></div>
    </form>
</body>
</html>
