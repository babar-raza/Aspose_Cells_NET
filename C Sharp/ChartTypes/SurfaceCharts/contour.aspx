<%@ Page Language="c#" Codebehind="Contour.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.Contour"
    MasterPageFile="~/tpl/Demo.Master" 
    Title="Contour - Aspose.Cells Demos" %>

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
                        Contour - Aspose.Cells
                    </h2>
                </td>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
            </tr>
        </tbody>
    </table>
    
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
    <p>
        This online demo demonstrates how to create a Surface Contour chart in a worksheet using
        <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                        Aspose.Cells</a> for .NET.
    </p>
    <p>
        Contour is a surface chart viewed from above, where colors represent specific ranges of values. 
        Displayed without color, this chart type is called a Wireframe Contour. Aspose.Cells is a powerful component, 
        which supports all the standard and custom charts to help you display data in more meaningful ways. 
        You may create many kinds of charts including Surface Contour.
    </p>
    <p>
        The demo creates a workbook first and inputs some chart related data into the first six columns 
        (A, B, C, D, E and F) of the first worksheet named Data. The first column denotes a category, time (in seconds), 
        related to different ranges (0.2 - 1.0) where as the second, third, fourth, fifth and sixth columns contain 
        values of temperature series (10, 20, 30, 40 and 50). The demo creates a surface contour chart titled 
        "Tensile strength Measurements" into the second worksheet named "Chart" based on time and temperature 
        values in the first worksheet. In the demo, you are provided a sample snapshot of the chart, a drop down list 
        that represents chart type (SurfaceContour and SurfaceContourWirframe) and a command button labeled "Create Report" 
        to create and exercise the chart based on your selection from the drop down list. You may either open the resultant 
        excel file into MS Excel or save directly to your disk.
    </p>
    
    <p>
        Click <b>Create Report</b> to see how demo can  set the appearance properties of a contour chart.
    </p>
    </div>
   
    <table class="genericTable" style="text-align: left; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" align="right">
                <img alt="" src="../../Image/Contour.jpg"/></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Chart Type:</td>
                        <td>
                            <asp:DropDownList runat="server" ID="ChartTypeList">
                                <asp:ListItem Value="0">SurfaceContour</asp:ListItem>
                                <asp:ListItem Value="1">SurfaceContourWireframe</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Save Format:
                        </td>
                        <td>
                           <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                            <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                            <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
                           </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <asp:Button ID="btnProcess" runat="server" Text="Create Report"></asp:Button></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
