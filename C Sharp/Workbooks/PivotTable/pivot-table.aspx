<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="pivot-table.aspx.cs" Inherits="Aspose.Cells.Demos.Pivot_Table" Title="Adding Pivot Table - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Adding Pivot Table - Aspose.Cells
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
            This online demo exhibits <b>how to add a Pivot Table</b> in a worksheet using <a
                href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET</p>
        <p>
            <b>PivotTables</b> can be added to the spreadsheets using Aspose.Cells. Aspose.Cells
            provides some special set of classes that are used to create and set the PivotTables.
            These classes are used to create and set <b>PivotTable Objects</b>, which act as
            the building blocks of a <b>PivotTable</b>.</p>
        <p>
            Click <b>Process </b>to see how demo can set the appearance properties of a pivot
            table. You can either open the resulting excel file into <b>MS Excel</b> or save
            directly to your disk.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
    </div>
</asp:Content>
