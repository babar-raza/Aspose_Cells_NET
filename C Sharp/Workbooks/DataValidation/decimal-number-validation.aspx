<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="decimal-number-validation.aspx.cs" Inherits="DecimalNumberValidation"
    Title="Applying Decimal Number Validation - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Applying Decimal Number Validation - Aspose.Cells</h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This demo shows how to apply decimal number <b>Data Validation</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            Data validation is a strong feature by Aspose.Cells that helps developers to <b>Validate
                Information</b> that is entered in their worksheets. With data validation, developers
            can provide users with a list of choices, restrict data entries to a specific type
            or size etc. With Decimal Number Data Validation, you can allow the user enter decimal
            numbers into the related cells within a specified range or crieteria. Following
            demo shows how to implment <b>Decimal ValidationType</b>.</p>
        <p>
            Click <b>Process </b>to see how example creates an excel file with decimal number
            validation applied to range of Cells (<b>A1:A10</b>). You can either open the resulting
            excel file into <b>MS Excel</b> or save directly to your disk.
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
