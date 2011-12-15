<%@ Page Language="c#" CodeBehind="hiding-rows-and-columns.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.HidingRowsAndColumns" MasterPageFile="~/tpl/Demo.Master"
    Title="Hiding Rows and Columns - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Hiding Rows and Columns - Aspose.Cells
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
            two command buttons, Insert and Delete for practice. When you click Insert button,
            It creates a workbook and inputs some data into different cells (A1, A2, A3 and
            B1) of the first worksheet. It then inserts 10 rows at 4<sup>th</sup> row and inserts
            a column at the second column index (B). When you click Delete button, It creates
            a workbook and inputs some data into different cells (A1, A2, B3, B13 and C1) of
            the first worksheet. It then removes 10 rows starting from 3<sup>rd</sup> row and
            deletes the 3<sup>rd</sup> column (C) from the worksheet. You can either open the
            resultant excel files into MS Excel or save directly to your disk.
        </p>
        <p>
            <asp:Button runat="server" ID="Button1" Text="Display" />&nbsp;&nbsp;
            <asp:Button runat="server" ID="Button2" Text=" Hide " />
        </p>
    </div>
</asp:Content>
