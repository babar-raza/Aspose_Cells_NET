<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="other-drawing-objects.aspx.cs" Inherits="Workbooks_DrawingObjects_OtherDrawingObjects"
    Title="Other Drawing Objects - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Other Drawing Objects - Aspose.Cells</h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo describes how to <b>Add Drawing objects</b> to the worksheet using
            <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Aspose.Cells component provides the way to add and customize the controls (<b>TextBox</b>,
            <b>CheckBox</b> etc.) to a worksheet. The component also allows to insert embedded
            OLE objects. The demo creates an excel file and adds a text box control to <b>B2</b>
            cell of the first worksheet. It then, adds an image and an excel file data to stream
            and embeds it to become an embedded <b>OLE object</b> in the worksheet. It also
            sets the border around the object.You can either open the resulting excel file into
            <b>MS Excel</b> or save directly to your disk.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/Designer/book1.xls">book1.xls</asp:HyperLink>
            and
            <asp:HyperLink ID="lnkImg" runat="server" NavigateUrl="~/Image/School.jpg">School.jpg</asp:HyperLink>
            used in this demo.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
