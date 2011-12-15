<%@ Page Language="c#" CodeBehind="inserting-and-deleting-rows-and-columns.aspx.cs"
    AutoEventWireup="false" Inherits="Aspose.Cells.Demos.InsertingAndDeletingRowsAndColumns"
    MasterPageFile="~/tpl/Demo.Master" Title="Insert and Delete Rows & Columns - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Insert and Delete Rows & Columns - Aspose.Cells
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
            This online demo demonstrates the exercise to manipulate (insert, delete) rows and
            columns in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Aspose.Cells component allows you to manipulate rows and columns in the worksheet.
            You can insert blank cells, rows, and columns and fill them with data. Moreover
            you can remove any column or row in the worksheet. In the demo you are provided
            two command buttons, Insert and Delete for practice.
        </p>
        <p>
            When you click <b>Insert</b> button. It creates a workbook and inputs some data into cells,
            then it inserts 10 rows starting from 3rd row and inserts 3rd column</p>
        <p>
            When you click <b>Delete</b> button. It creates a workbook and inputs some data into cells,
            then it deletes 10 rows starting from 3rd row and deletes 3rd column</p>
        <p>
            You can either open the resultant excel files into <b>MS Excel</b> or save directly to
            your disk.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="Button1" runat="server" Text="Insert"></asp:Button>&nbsp;&nbsp;
            <asp:Button ID="Button2" runat="server" Text="Delete"></asp:Button>
        </p>
    </div>
</asp:Content>
