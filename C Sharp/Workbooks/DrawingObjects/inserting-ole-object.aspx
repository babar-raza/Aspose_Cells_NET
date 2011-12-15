<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="inserting-ole-object.aspx.cs" Inherits="Workbooks_DrawingObjects_InsertingOleObject"
    Title="Insert Ole Objects - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Inserting Ole Objects - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo exhibits how to <b>Insert an Ole Object</b> in a worksheet using
            <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            OLE (Object Linking and Embedding) is Microsoft's framework for a compound document
            technology. Aspose.Cells supports to add / manipulate Ole Objects into your worksheets
            and can give more value to your workbooks. Aspose.Cells provides some important
            API related the task. In this demo we will <b>insert an Image into an Excel File</b>
            as an ole object. Both Image and Excel file data is read using the <b>Stream</b>.
            You can either open the resulting excel file into <b>MS Excel</b> or save directly
            to your disk.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="/designer/book1.xls">book1.xls</asp:HyperLink>
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
