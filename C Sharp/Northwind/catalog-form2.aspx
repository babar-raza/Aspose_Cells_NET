<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="catalog-form.aspx.cs" Inherits="Aspose.Cells.Demos.Northwind.CatalogForm" 
MasterPageFile="~/tpl/DemoEmpty.Master" Title="Catalog - Apose.Cells Demos"%>
<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <p class="componentDescriptionTxt">
    	Click <b>Execute</b> to see how example 
        Prints a catalog of products. The document will contain two-page report header, photos for each category;
		the demo starts each category on a new page; keeps all records for a category on same page;
        prints an order form in the report footer on a separate page.
        <br/>You can either open the resulting excel file into MS Excel
        or save directly to your disk.
    </p>
    <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click"/>
</asp:Content>