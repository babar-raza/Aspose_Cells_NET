<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="summary-by-year-form.aspx.cs"
    Inherits="Aspose.Cells.Demos.Northwind.SummaryByYearForm" MasterPageFile="~/tpl/Demo.Master"
    Title="Summary of Sales by Year - Apose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Summary of Sales by Year - Aspose.Cells</h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo illustrates how to create a well formatted Summary of Sales by
            Year report using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Aspose.Cells component gives you the agility to report your data in a variety of
            ways. Aspose.Cells component is fully functional for creating all types of reports.
            You may customize the size and appearance of everything on a report. You can display
            the information the way you want to see it.
        </p>
        <p>
            The demo generates a summary report showing quarterly sales for each year. ADO.NET
            is used to retrieve the data from the Orders table of Northwind database, to generate
            the report. You can either open the resultant excel file into MS Excel or save directly
            to your disk to check the results.</p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/Northwind.xls">Northwind.xls</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
    </div>
</asp:Content>
