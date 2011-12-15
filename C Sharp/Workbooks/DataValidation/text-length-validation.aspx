<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="text-length-validation.aspx.cs" Inherits="TextLengthValidation" Title="Applying Text Length Validation - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Applying Text Length Validation - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This demo shows how to apply text <b>Length Validation</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            Data validation is a strong feature by Aspose.Cells that helps developers to validate
            information that is entered in their worksheets. With data validation, developers
            can provide users with a list of choices, restrict data entries to a specific type
            or size etc. With this type of validation, you can allow the user enter text values
            of specified length into the related cells. Following is the demo, which shows how
            to implement <b>TextLength ValidationType</b>. In the demo, the user is restricted
            to enter string values with the <b>length lesser than or equal to 5</b>. The validation
            area is <b>B1</b> cell.</p>
        <p>
            Click <b>Process </b>to see how example creates an excel file with text length validation
            applied to cell <b>B1</b>. You can either open the resulting excel file into <b>MS Excel</b>
            or save directly to your disk.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
    </div>
</asp:Content>
