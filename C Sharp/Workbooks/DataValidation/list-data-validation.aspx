<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="list-data-validation.aspx.cs" Inherits="ListDataValidation" Title="Applying List Data Validation - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Applying List Data Validation - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This demo shows how to apply list <b>Data Validation</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            <b>Data Validation</b> is a strong feature by Aspose.Cells that helps developers
            to <b>Validate Information</b> that is entered in their worksheets. With data validation,
            developers can provide users with a list of choices, restrict data entries to a
            specific type or size etc. List data validation allows the user enter values from
            a drop down list. It provides a list, a series of rows that contains related data.
            Following demo shows how to implment <b>List ValidationType</b>. In the demo, a
            second worksheet is added, which represents the <b>Source of List</b>, the user
            is restricted to select values form the list only, the validation area is <b>A1:A5</b>
            in the first worksheet. It is important here that you set <b>Validation.InCellDropDown</b>
            property to true.</p>
        <p>
            Click <b>Process </b>to see how example creates an excel file with list data validation
            applied to range of Cells (<b>A1:A5</b>). You can either open the resulting excel
            file into <b>MS Excel</b> or save directly to your disk.
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
