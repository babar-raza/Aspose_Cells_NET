<%@ Page Language="c#" CodeBehind="export-data.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.ExportData"
    MasterPageFile="~/tpl/Demo.Master" Title="Exporting data from excel files - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tbody>
            <tr>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%;">
                    <h2 class="demos-heading-bg">
                        Exporting Data - Aspose.Cells
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
            This online demo shows how to retrieve data from a worksheet and <b>Export</b> to
            fill a DataTable object using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET
        </p>
        <p>
            The demo gets data from a worksheet of an excel file and exports it to fill a DataTable.
            The example also binds the <b>DataTable object</b> to a <b>DataGrid control</b>.
            This is achieved through a single API of Aspose.Cells. All this is done with so
            much ease and it minimizes your coding lines which may extend while performing this
            operation manually.
        </p>
        <p>
            Click <b>Export Data</b> to see how example exports part of the excel file into
            a <b>DataTable</b> showing the table in a grid as a result.
        </p>
        <p>
            Please download
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~\designer\book1.xls">book1.xls</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:Button runat="server" ID="btnExportData" Text="Export Data" />
        </p>
        <p>
            <asp:DataGrid ID="dgExportData" runat="server" CellPadding="4" CellSpacing="0" Width="90%"
                BorderColor="#aaaaaa">
            </asp:DataGrid>
        </p>
    </div>
</asp:Content>
