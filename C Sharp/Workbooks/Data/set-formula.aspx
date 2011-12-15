<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="set-formula.aspx.cs" Inherits="Workbooks_Data_SetFormula" Title="Using Formulas/Functions to Process Data - Aspose.Cells Demos" %>

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
                        Formulas - Aspose.Cells
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
            This demo exhibits an exercise of different types of <b>Worksheet Formulas / functions</b>
            to process data in the spreadsheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Aspose.Cells component supports all the commonly used <b>Functions</b> related to
            different categories: <b>Mathematical</b>, <b>String</b>, <b>Statistical</b>, <b>DateTime</b>,
            <b>Logical</b>, <b>Database</b> and <b>Financial</b> etc.
        </p>
        <p>
            The demo makes use of a template excel file which contains a list of all the formulas
            / functions string of all the categories mentioned. The file also contains some
            static data used in different formulas. The demo retrieves the formulas / functions
            string and calculates the formulas / functions to fill the column <b>D</b> with
            resultant data. You can either open the resulting excel file into <b>MS Excel</b>
            or save directly to your disk.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~\designer\Workbooks\Formula.xls">Formula.xls</asp:HyperLink>
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
