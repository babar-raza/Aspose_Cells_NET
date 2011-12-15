<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="financial-plan.aspx.vb" Inherits="Aspose.Cells.Demos.FinancialPlan"
	Title="Financial Plan - Aspose.Cells Demos" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tbody>
			<tr>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%;">
					<h2 class="demos-heading-bg" style="font-family: Arial; font-size: large;">
						&nbsp;Financial plan - Aspose.Cells</h2>
				</td>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			The Five Year Plan (Service Industry) template has been designed to help a financial
			services company --- such as a small bank, mortgage broker, or savings-and-loan
			company. The high-level financial plan defines your financial model and pricing
			assumptions with all other important aspects. This plan also includes expected annual
			sales and profits for the next five years.
		</p>
		<p>
			The demo uses a template file which has five worksheets named Model Inputs, Profit
			and Loss, Balance Sheet, Cash Flow and Loan Payment Calculator. All the sheets are
			filled with data with all its formatting using APIs of Aspose.Cells component to
			produce a complete 5 years Financial Plan for any Corporate.
		</p>
		<p>
			You can either open the resultant excel file into MS Excel or save directly to your
			disk to check the results.
		</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/FinancialPlan.xls">FinancialPlan.xls</asp:HyperLink>
			used in this demo.</p>
		<p class="componentDescriptionTxt">
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnProcess" runat="server" Text="Process" OnClick="btnProcess_Click" /></p>
	</div>
</asp:Content>
