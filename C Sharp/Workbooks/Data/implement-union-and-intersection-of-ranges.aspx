<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    Inherits="Union_Intersection" Title="Implement Union & Intersection of Ranges - Aspose.Cells Demos"
    CodeBehind="implement-union-and-intersection-of-ranges.aspx.cs" %>

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
                        Implement Union &amp; Intersection of Ranges - Aspose.Cells
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
            This online demo explains how to take <b>Union and Intersection of Ranges</b> using
            <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            Aspose.Cells allows you to intersect the two ranges. You can also take the union
            of the ranges. The demo uses an excel spreadsheet named "<b>BKRanges.xls</b>" and
            performs the union and intersection operations on the specified ranges in the workbook.
            In the generated file, the cells shaded with green color is the resultant <b>range</b>
            for intersection operation. Similarly, the cells shaded with yellow color are the
            resultant ranges for union operation.
        </p>
        <p>
            Click <b>Process </b>to see how example takes union and intersection of ranges in
            the template file's sheet. You can either open the output excel file into your <b>MS
                Excel</b> or save directly to your disk.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~\designer\Workbooks\BKRanges.xls">BKRanges.xls</asp:HyperLink>
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
