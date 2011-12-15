<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="adding-document-properties.aspx.cs" Inherits="Workbooks_Worksheets_AddingDocumentProperties"
    Title="Adding Custom Document Properties - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Adding Custom Document Properties - Aspose.Cells</h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo describes how to add custom document properties to a workbook using
            <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET
        </p>
        <p>
            <b>Microsoft Excel</b> provides a feature to add some properties to the Excel files.
            These document properties allow some useful information to be stored along with
            the documents (Excel files). There are two kinds of document properties, <b>System Defined
                (Built-in) Properties</b> and <b>User Defined (Custom) Properties</b>. Built-in
            properties contain general information about the document like document title, author's
            name, document statistics and so on. Custom properties are those ones, which are
            defined by the users as Name/Value pairs, where both name and value are defined
            by the user. The most important point to know about the built-in and custom properties
            is that built-in properties can be accessed and modified only but not created or
            removed because these properties are system defined. However, custom properties
            can be created and managed freely by the developers because these properties are
            defined by the users.
        </p>
        <p>
            In this demo, we will add custom document properties using Aspose.Cells.</p>
        <p>
            Click <b>Execute</b> to see how example creates an excel file and add some custom
            document properties to a workbook. You can either open the resulting excel file
            into <b>MS Excel</b> or save directly to your disk.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/book1.xls">book1.xls</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <%--<asp:ListItem Value="XLSX">XLSX</asp:ListItem>--%>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Execute" OnClick="btnExecute_Click" />
        </p>
        <p>
            <asp:Image ID="imgCustomDocumentProperties" runat="server" ImageUrl="~/Image/CustomDocumentProperties.jpg" />
        </p>
    </div>
</asp:Content>
