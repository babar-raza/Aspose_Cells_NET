<%@ Page Language="c#" CodeBehind="grouping-rows-and-columns.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.GroupingRowsAndColumns" MasterPageFile="~/tpl/Demo.Master"
    Title="Group and Ungroup Rows & Columns - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Group and Ungroup Rows & Columns - Aspose.Cells
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
            This online demo specifies how to <b>Group / Ungroup</b> rows and column in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Aspose.Cells component allows you to group / ungroup rows and columns in the worksheet.
            A show detail symbol <b>+</b> and a hide detail symbol <b>-</b> attached to the
            rows and columns headers specifying the groups in the worksheet. Aspose.Cells component
            also provides you the feasibility to specify the outline setting of the summary
            row / column according to your requirements. In the demo, you are provided two command
            buttons ,<b>Group</b> and <b>Ungroup</b> for practice. When you click <b>Group</b>
            button, It uses a template excel file and <b>groups first 10 rows</b> of the first
            worksheet. It also <b>groups the first two columns</b> of the worksheet. When you
            click <b>Ungroup</b> button, It uses a template excel file and un-groups the grouped
            rows and columns in the worksheet. You can either open the resultant excel files
            into <b>MS Excel </b>or save directly to your disk.</p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFileGroup" runat="server" NavigateUrl="~/designer/Workbooks/GroupingRowsAndColumns.xls">GroupingRowsAndColumns.xls</asp:HyperLink>
            and
            <asp:HyperLink ID="lnkFileUngroup" runat="server" NavigateUrl="~/designer/Workbooks/UnGroupingRowsAndColumns.xls">UnGroupingRowsAndColumns.xls</asp:HyperLink>
            used in this demo.
        </p>
        <table>
            <tr>
                <td>
                    <asp:Button ID="Button1" Text="Group " runat="server"></asp:Button>&nbsp;
                    <asp:Button ID="Button2" Text="UnGroup" runat="server" />
                </td>
            </tr>
        </table>
    </div>
</asp:Content>
