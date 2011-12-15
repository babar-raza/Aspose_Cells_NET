<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="named-ranges.aspx.cs" Inherits="Workbooks_Data_NamedRanges" Title="Named Ranges - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tbody>
            <tr>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%;">
                    <h2 class="demos-heading-bg">
                        Named Ranges - Aspose.Cells
                    </h2>
                </td>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This demo explains how to create and access <b>Named Ranges</b> of cells in your
            worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Aspose.Cells allows you to manipulate Named Ranges . You can use the labels of columns
            and rows on a worksheet to refer to the cells within those columns and rows. You
            can create descriptive names to represent <b>cells</b>, <b>ranges of cells</b>,
            <b>formulas</b>, or <b>constant values</b>. Labels can be used in formulas that
            refer to data on the same worksheet; if you want to represent a range on another
            worksheet, use a name. The demo creates a named range (<b>B1:E5</b>) of cells first.
            It then, names the range. It then, accesses it by its name and puts some value into
            its left and bottom most cells with some formatting. You can either open the resulting
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
