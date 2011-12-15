<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="setting-page-option.aspx.cs" Inherits="Workbooks_PageSetup_SettingPageOption"
    Title="Setting Page Options - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Setting Page Options - Aspose.Cells
                    </h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo describes how to implement Page Setup Settings using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            Many a times, you would like to configure page setup settings for your worksheets
            to control the printing process. For example, you may need to set Page Orientation,
            Scaling Factor, Paper Size, Print Quality and First Page Number etc. These page
            setup settings offer various options to be configured. All the page setup options
            are completely supported in Aspose.Cells.
        </p>
        <p>
            The demo makes use of an existing excel file, opens it and implements some page
            setup settings for the first worksheet of the workbook. It sets landscape orientation
            with zoom factor value equal to 10, A4 paper size, print quality value equal to
            200 dpi (dot per inches) with 1 as first page number. You can either open the resultant
            excel file into MS Excel or save directly to your disk to check the results.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/book1.xls">book1.xls</asp:HyperLink>
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
