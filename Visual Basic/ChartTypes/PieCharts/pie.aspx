<%@ Page Language="vb" Codebehind="Pie.aspx.vb" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.Pie"
	MasterPageFile="~/tpl/Demo.Master" Title="Pie - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
				font-size: large;">
				<h2 class="demos-heading-bg">
					Pie - Aspose.Cells</h2>
			</td>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo exhibits how to create a <b>Pie chart</b> with <b>2-D</b> and <b>3-D</b>
			visual effects in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			This type of chart displays the contribution of each value to a total. It is also
			available with a 3-D visual effect. Aspose.Cells is a powerful component, which
			supports all the standard and custom charts to help you display data in more meaningful
			ways. The demo creates a workbook first and inputs the chart source data into the
			first two columns (A and B) of the first worksheet. The first column represents
			the category data (Region) where as the second column represents the sales data
			representing values.
		</p>
		<p>
			The demo creates a pie chart representing Sales By Region into the worksheet named
			Pie based on the different sale values related to different regions. In the demo,
			you are provided a sample snapshot of the chart, two drop down lists which represent
			first slice angle (90, 80, 180 and 360) and data label position (Center, InsideBase,
			InsideEnd and OutSideEnd), a check box that represents whether you want to create
			the chart with 3-D flavor and a command button labeled Create Report to create and
			exercise the chart using your desired inputs. You can either open the resultant
			excel file into <b>MS Excel</b> or save directly to your disk to check the results.
		</p>
		<p>
			Click <b>Create Report</b> to see how demo can set the appearance properties of
			a pie chart.</p>
			</div>
		<table class="genericTable" style="font-family: Arial; font-size: small;">
			<tr>
				<td align="right">
					<img alt="" src="../../Image/Pie.jpg" /></td>
				<td valign="top" align="left">
					<table class="genericTable">
						<tr>
							<td>
								FirstSliceAngle:
							</td>
							<td>
								<asp:DropDownList ID="FirstSliceAngle" runat="server">
									<asp:ListItem Value="0">0</asp:ListItem>
									<asp:ListItem Value="1">90</asp:ListItem>
									<asp:ListItem Value="2">180</asp:ListItem>
									<asp:ListItem Value="3">360</asp:ListItem>
								</asp:DropDownList></td>
						</tr>
						<tr>
							<td>
								DataLabels Postion:</td>
							<td>
								<asp:DropDownList ID="LabelsPostionList" runat="server">
									<asp:ListItem Value="0">Center</asp:ListItem>
									<asp:ListItem Value="1">InsideBase</asp:ListItem>
									<asp:ListItem Value="2">InsideEnd</asp:ListItem>
									<asp:ListItem Value="3" Selected="True">OutsideEnd</asp:ListItem>
								</asp:DropDownList></td>
						</tr>
						<tr>
							<td>
								Show as 3D:</td>
							<td>
								<asp:CheckBox runat="server" ID="CheckBoxShow3D" /></td>
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
