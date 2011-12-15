<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="hello-world.aspx.vb" Inherits="Workbooks_Data_HelloWorld" Title="Hello World - Aspose.Cells Demos" %>

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
						Hello World - Aspose.Cells
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
			This online demo demonstrates the ability of <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET to add and retrieve data from the cells of a worksheet.
		</p>
		<p>
			The demo is a perfect example for any <b>beginner</b> who wants to learn the APIs
			of Aspose.Cells and create his/her first Hello World conventional example and quickly
			put into action to make the start. This demo utilizes a template file and then enlightens
			how to insert data of <b>different format</b> and <b>data types</b> (like String,
			Numeric, Boolean and DateTime etc.) into different cells of a worksheet. You can
			either open the resulting excel file into <b>MS Excel</b> or save directly to your
			disk.
		</p>
		<p>
			Please download
			<asp:HyperLink ID="TemplateLink" runat="server" NavigateUrl="~\designer\Workbooks\HelloWorld.xls">HelloWorld.xls</asp:HyperLink>
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
