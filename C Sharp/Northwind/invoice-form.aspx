<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="invoice-form.aspx.cs" Inherits="Aspose.Cells.Demos.Northwind.InvoiceForm"
    MasterPageFile="~/tpl/Demo.Master" Title="Invoice - Apose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Invoice - Aspose.Cells</h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo illustrates how to create a well formatted Customer Invoice report
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Aspose.Cells component gives you the agility to report your data in a variety of
            ways. Aspose.Cells component is fully functional for creating all types of reports.
            You may customize the size and appearance of everything on a report. You can display
            the information the way you want to see it.
        </p>
        <p>
            The demo generates a printed report displaying the customer invoice against the
            order. A separate sheet is generated for each customer Invoice. ADO.NET is used
            to retrieve the data from the Invoices and Orders tables of Northwind database,
            to generate the report. You can either open the resultant excel file into MS Excel
            or save directly to your disk to check the results.</p>
        <p>
            Click Process to see how example Prints each invoice on a separate page creating
            the Excel spreadsheet completely via API.
        </p>
        <p>
            Please download
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/image/logo.jpg">logo.jpg</asp:HyperLink>
            and
            <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/image/namelogo.jpg">namelogo.jpg</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
