<%@ Page Language="c#" CodeBehind="merge-unmerge-cells.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.MergeUnMergeCells" MasterPageFile="~/tpl/Demo.Master"
    Title="How to Merge and Unmerge the Cells - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Merge/UnMerge Cells - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo explains how to <b>Merge</b> / <b>UnMerge</b> cells in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Sometimes, while working with worksheets, you don't always want the same number
            of cells in every row or column. You might want to put a title in a single cell
            that spans the top of your worksheet. You might be creating an <b>invoice</b>, and
            want a fewer columns for the total. When you want to make one cell from two or more
            cells, you <b>Merge the Cells</b>. Aspose.Cells has the ability to merge the cells
            in the worksheet. You may unmerge the merged cells too. A Merged Cell is basically
            a single cell that is created by combining two or more selected cells. The demo
            shows how to merge by pressing <b>Merge</b> some cells (<b>C6:E7</b>) in the worksheet.
            Also, by pressing <b>UnMerge</b>, un-merge the merged cells (<b>C6</b>) in the worksheet.
            You can either open the resulting excel file into <b>MS Excel</b> or save directly
            to your disk.</p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/Workbooks/MergeCells.xls">MergeCells.xls</asp:HyperLink>
            used in this demo.
        </p>
        <table class="genericTable" style="font-size: 10pt; font-family: Arial">
            <tr>
                <td>
                    <asp:Button runat="server" ID="btnMerge" Text="Merge" />&nbsp;&nbsp;
                    <asp:Button runat="server" ID="btnUnMerge" Text=" UnMerge " />
                </td>
            </tr>
        </table>
    </div>
</asp:Content>
