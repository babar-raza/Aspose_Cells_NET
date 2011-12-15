<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="time-validation.aspx.cs" Inherits="TimeDataValidation" Title="Applying Time Validation - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Applying Time Validation - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This demo shows how to apply <b>Time Validation</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            Data validation is a strong feature by Aspose.Cells that helps developers to validate
            information that is entered in their worksheets. With data validation, developers
            can provide users with a list of choices, restrict data entries to a specific type
            or size etc. With this type of validation, you can allow the user enter Time values
            into the related cells within a specified range or crieteria. Following is the example,
            which shows how to implment Time ValidationType. In the demo, the user is restricted
            to enter <b>Time values between 09:00 to 11:30 AM</b> only. Here, the validation
            area is <b>B1</b> cell.</p>
        <p>
            Click <b>Process </b>to see how example creates an excel file with time data validation
            applied to Cell "<b>B1</b>". You can either open the resulting excel file into <b>MS
                Excel</b> or save directly to your disk.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <%--<asp:ListItem Value="XLSX">XLSX</asp:ListItem>--%>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
