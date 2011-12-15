<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="cost-pareto.aspx.cs" Inherits="Aspose.Cells.Demos.CostPareto" Title="Cost Pareto - Aspose.Cells Demos" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tbody>
            <tr>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%;">
                    <h2 class="demos-heading-bg">
                        Cost Pareto - Aspose.Cells
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
            A <b>Pareto chart</b>, is a type of chart that contains both bars and a line graph,
            where individual values are represented in descending order by bars, and the cumulative
            total is represented by the line.
        </p>
        <p>
            This demo utilizes an xml file, reads the xml data from the file to generate two
            reports named Cost Data and Pareto Chart each representing a worksheet in the workbook.
            The first report calculates the annual cost of different cost centers while the
            second worksheet presents a column chart based on the first worksheet data.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/Database/CostPareto.xml">CostPareto.xml</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnProcess" runat="server" Text="Process" OnClick="btnProcess_Click" />
        </p>
    </div>
</asp:Content>
