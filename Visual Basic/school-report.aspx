<%@ Page Language="vb" Codebehind="school-report.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.SchoolReport" MasterPageFile="~/tpl/Demo.Master"
	Title="School Report - Aspose.Cells Demos" %>

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
						School Report - Aspose.Cells
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
			Please choose a student and click <b>Generate Report</b> to see student report cards
			generated and sent to user in a resulting Excel file.</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/SchoolData.xls">SchoolData.xls</asp:HyperLink>
			and
			<asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/Image/School.jpg">SchoollData.jpg</asp:HyperLink>
			used in this demo.</p>
		<p align="center">
			<img src="image/school.jpg">
		</p>
		<p align="center">
			<strong>Please Choose a Student:</strong>
		</p>
		<p align="center">
			<asp:ListBox ID="ListBox1" runat="server" Width="157px"></asp:ListBox>
		</p>
		<p align="center">
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="Button1" runat="server" Text="Generate Report"></asp:Button></p>
	</div>
</asp:Content>
