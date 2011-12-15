<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="using-icustom-function.aspx.vb" Inherits="Workbooks_Data_UsingICustomFunction"
	Title="Using ICustomFunction - Aspose.Cells Demos" %>

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
						Using ICustomFunction - Aspose.Cells
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
			<b>ICustomFunction</b> is very usefull feature when user has defined his own funtions
			in template file. Using <b>ICustomFunction</b> feature, the same user defined function
			can be defined in user's application and Aspose.Cells will execute it like other
			default <b>Excel</b> functions and evaluate there values. This demo exhibits the
			use of <b>ICustomFunction</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			In this demo we create a new workbook with some sample data. We create a cutom function
			with two parameters, a cell with decimal value and a range of cells. In Custom Function
			(MyFunc) we add the values of second parameter (<b>C1:C5 range</b>) and divide it
			with first parameter value (<b>B1 decimal value</b>). We assign this function to
			cell <b>A1</b> and use CalculateFormula API to see how it evaluates the <b>ICustomFunction</b>.
			You can either open the resulting excel file into <b>MS Excel</b> or save directly
			to your disk.
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
