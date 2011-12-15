<%@ Page Language="vb" Codebehind="chart-demos.aspx.vb" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.ChartDemos"
	MasterPageFile="~/tpl/Demo.Master" 
	Title="Creating chart demos - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<p class="productTitle">Chart Demos - <a href="default.aspx">Aspose.Cells Demos</a></p>
	<p class="componentDescriptionTxt">
		For an overview of some standard chart types and their subtypes, please click the
		following list:
	</p>
	<div style="text-align: left">
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_1')">
				<img alt="Show" src="Image/blueup.gif" border="0">Column Charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_1">
			<p class="componentDescriptionTxt">
				A column chart shows data changes over a period of time or illustrates comparisons
				among items. Column charts have the following chart sub-types:</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<!--
						<a href="Aspose.Cells.Charts.ChartTypes/columncharts/clustered-column.aspx">Clustered Column</a></p>
						-->
						<a href="ChartTypes/columncharts/clustered-column.aspx">Clustered Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to set the appearance properties of a Clustered Column chart. This Demo
						exhibits how to create a <b>Clustered Column</b> chart in a worksheet. This type
						of chart compares values across categories. Normally the categories are organized
						horizontally, and values vertically, to emphasize variation over time. <b>Aspose.Cells</b>
						component is a powerful component, which supports all the standard and custom charts
						to help you display data in more meaningful ways. You may create many kinds of charts
						including <b>Clustered Column</b>. The component can insert a chart into the worksheet
						in a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs some chart related data into the first two columns (<b>A</b> and <b>B</b>)
						of the first worksheet. The first column represents the category data (<b>Region</b>)
						and the second column represents values (<b>Marketing Costs</b>). The demo creates
						the chart representing <strong>Marketing Costs By Region</strong> based on the different
						costs involving different regions. In the example, you have been provided a sample
						snapshot of the chart and a few controls that represent the related list of data
						including two text boxes which represent category axis and value axis titles, five
						drop down lists which represent major unit, minor unit, minimum, maximum values
						of value axis and gap width that represents the space b/w the column clusters in
						the chart and a command button labeled <b>Create Report</b> to create and exercise
						the chart using your desired inputs. You can either open the resultant excel file(s)
						into MS Excel or save directly to your disk to check the results.</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/columncharts/stacked-column.aspx">Stacked Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to set the appearance properties of a Stacked Column chart. This Demo
						describes how to create a <b>Stacked Column</b> chart with simple and 3-D visual
						effects in a worksheet. This type of chart shows the relationship of individual
						items to the whole, comparing the contribution of each value to a total across categories.
						<b>Aspose.Cells</b> is a powerful component, which supports all the standard and
						custom charts to help you display data in more meaningful ways. You may create many
						kinds of charts including <b>Stacked Column</b>. The component can create the chart
						into the worksheet in a workbook using the simplest APIs with ease. The demo creates
						a workbook first and inputs some chart related data into the first three columns
						(<b>A</b>, <b>B </b>and<b> C</b>) of the first worksheet named <b>Data</b>. The
						first column represents the category data (<b>Year</b>) where as the second and
						third column represent values for <b>Product1</b> and <b>Product2</b>. The demo
						creates a stacked column chart representing <strong>Product Sales</strong> into
						the second worksheet named <b>Chart</b> based on the different product values related
						to different years (2004-2006) in the first worksheet. In the example, you have
						been provided a sample snapshot of the chart, a check box that represents whether
						you want to create a 3-D stacked column chart and a command button labeled <b>Create
							Report</b> to create and exercise the chart using your desired inputs. You can
						either open the resultant excel file(s) into MS Excel or save directly to your disk
						to check the results.</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/columncharts/percent-stacked-column.aspx">100% Stacked Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to set the appearance properties of a 100% Stacked Column chart. This
						Demo demonstrates how to create a <b>100% Stacked Column</b> chart with simple and
						3-D visual effects in a worksheet. This type of chart compares the percentage each
						value contributes to a total across categories. <b>Aspose.Cells</b> is a powerful
						component, which supports all the standard and custom charts to help you display
						data in more meaningful ways. You may create many kinds of charts including <b>100%
							Stacked Column</b>. The component can create the chart into the worksheet in
						a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs some chart related data into the first five columns (<b>A</b>, <b>B</b>,<b>
							C</b>,<b> D </b>and<b> E</b>) of the first worksheet named <b>Data</b>. The
						first column represents the product names (<b>Product1</b>, <b>Product2</b> and
						<b>Product3</b>) where as the second, third, fourth and fifth columns represent
						percentage values involving different quarters (<b>Qurarter1</b>, <b>Quarter2</b>,
						<b>Quarter3</b> and <b>Quarter4</b>) which represent category data. The demo creates
						a 100% stacked column chart representing <strong>Product Contribution to Total Sales</strong>
						into the second worksheet named <b>Chart</b> based on the different product values
						related to different quarters in the first worksheet. In the example, you are provided
						a sample snapshot of the chart, a check box that represents whether you want to
						create the chart with 3-D flavor and a command button labeled <b>Create Report</b>
						to create and exercise the chart using your desired inputs. You can either open
						the resultant excel file(s) into MS Excel or save directly to your disk to check
						the results.</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/columncharts/column-3d.aspx">3D Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to set the appearance properties of a 3D Column chart. This Demo exhibits
						how to create a <b>3-D Column</b> chart with simple, clustered and stacked flavors
						in a worksheet. This type of chart compares data points along two axes. <b>Aspose.Cells</b>
						component is a powerful component, which supports all the standard and custom charts
						to help you display data in more meaningful ways. You may create many kinds of charts
						including <b>3-D Column</b>. The component can create the chart into the worksheet
						in a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs some chart related data into the first two columns (<b>A</b> and <b>B</b>)
						of the first worksheet. The first column represents the category data (<b>Region</b>)
						and the second column represents values (<b>Marketing Costs</b>). The demo creates
						a 3-D column chart representing <strong>Marketing Costs By Region</strong> based
						on the different costs involving different regions. In the example, you have been
						provided a sample snapshot of the chart and a few controls that represent the related
						list of chart data including a text box which represents chart type, five drop down
						lists which represent wall color, floor color, rotation angle, elevation angle and
						depth in percentage and a command button labeled <b>Create Report</b> to create
						and exercise the chart using your desired inputs. You can either open the resultant
						excel file(s) into MS Excel or save directly to your disk to check the results.</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_2')">
				<img alt="Show" src="Image/blueup.gif" border="0">Bar Charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_2">
			<p class="componentDescriptionTxt">
				A bar chart illustrates comparisons among individual items. Bar charts have the
				following chart sub-types:</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/barcharts/clustered-bar.aspx">Clustered Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to set the appearance properties of a Clustered Bar chart as 2D or 3D.
						This Demo exhibits how to create a <b>Clustered Bar</b> chart with 2-D and 3-D visual
						effects in a worksheet. This type of chart compares values across categories. Normally,
						categories are organized vertically, and values horizontally, to place focus on
						comparing the values. <b>Aspose.Cells</b> is a powerful component, which supports
						all the standard and custom charts to help you display data in more meaningful ways.
						You may create many kinds of charts including <b>Clustered Bar</b>. The component
						can create the chart into the worksheet in a workbook using the simplest APIs with
						ease. The demo creates a workbook first and inputs the source data related chart
						into the first three columns (<b>A</b>, <b>B </b>and<b> C</b>) of the first worksheet.
						The first column represents the category data (<b>Region</b>) where as the second
						and third columns represent the sales data representing values related to the products
						(<b>Apple</b> and <b>Orange</b>). The demo creates a clustered bar chart representing
						<b>Fruit Sales By</b> <b>Region</b> into the first worksheet named <b>Clustered Bar</b>
						based on the different product values related to different regions. In the example,
						you are provided a sample snapshot of the chart, a check box that represents whether
						you want to create the chart with 3-D flavor and a command button labeled <b>Create
							Report</b> to create and exercise the chart using your desired inputs. You can
						either open the resultant excel file(s) into MS Excel or save directly to your disk
						to check the results.</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/barcharts/stacked-bar.aspx">Stacked Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to set the appearance properties of a Stacked Bar chart as 2D or 3D. This
						Demo exhibits how to create a <b>Stacked Bar</b> chart with 2-D and 3-D visual effects
						in a worksheet. This type of chart shows the relationship of individual items to
						the whole. It is also available with a 3-D visual effect. <b>Aspose.Cells</b> is
						a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>Stacked Bar</b>. The component can create the chart into the worksheet in a workbook
						using the simplest APIs with ease. The demo creates a workbook first and inputs
						the source data related chart into the first four columns (<b>A</b>, <b>B</b>,<b> C
						</b>and<b> D</b>) of the first worksheet named <b>Data</b>. The first column represents
						the category data (<b>Region</b>) where as the second, third and fourth columns
						represent the sales data representing values related to different products (<b>Apple</b>,
						<b>Orange</b> <b>and Banana</b>). The demo creates a Stacked Bar chart representing
						<b>Fruit Sales By</b> <b>Region</b> into the second worksheet named <b>Chart</b>
						based on the different product values related to different regions in the first
						worksheet. In the example, you are provided a sample snapshot of the chart, a check
						box that represents whether you want to create the chart with 3-D visual effect
						and a command button labeled <b>Create Report</b> to create and exercise the chart
						using your desired inputs. You can either open the resultant excel file(s) into
						MS Excel or save directly to your disk to check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/barcharts/percent-stacked-bar.aspx">100 % Stacked Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to set the appearance properties of a 100% Stacked Bar. This Demo demonstrates
						how to create a <b>100% Stacked Bar</b> chart with simple and 3-D visual effects
						in a worksheet. This type of chart compares the percentage each value contributes
						to a total across categories. <b>Aspose.Cells</b> is a powerful component, which
						supports all the standard and custom charts to help you display data in more meaningful
						ways. You may create many kinds of charts including <b>100% Stacked Bar</b>. The
						component can create the chart into the worksheet in a workbook using the simplest
						APIs with ease. The demo creates a workbook first and inputs some chart related
						data into the first five columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D </b>and<b> E</b>)
						of the first worksheet named <b>Data</b>. The first column represents the product
						names (<b>Product1</b>, <b>Product2</b> and <b>Product3</b>) where as the second,
						third, fourth and fifth columns represent percentage values involving different
						quarters (<b>Qurarter1</b>, <b>Quarter2</b>, <b>Quarter3</b> and <b>Quarter4</b>)
						which represent category data. The demo creates a 100% stacked bar chart representing
						<b>Product Contribution to Total Sales</b> into the second worksheet named <b>Chart</b>
						based on the different product values related to different quarters in the first
						worksheet. In the example, you are provided a sample snapshot of the chart, a check
						box that represents whether you want to create the chart with 3-D visual effect
						and a command button labeled <b>Create Report</b> to create and exercise the chart
						using your desired inputs. You can either open the resultant excel file(s) into
						MS Excel or save directly to your disk to check the results.</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_3')">
				<img alt="Show" src="Image/blueup.gif" border="0">Line Charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_3">
			<p class="componentDescriptionTxt">
				A line chart shows trends in data at equal intervals. Line charts have the following
				chart sub-types:</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/linecharts/line.aspx">Line</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Line chart that displays trends over time or categories. This
						Demo exhibits how to create a <b>Line</b> chart in a worksheet. This type of chart
						displays trends over time or categories. It is also available with markers displayed
						at each data value. <b>Aspose.Cells</b> component supports all the standard and
						custom charts including <b>Line</b> chart to help you display data in more meaningful
						ways. The component can create the chart into the worksheet in a workbook using
						the simplest APIs with ease. The demo creates a workbook first and inputs the source
						data related chart into the first six columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D</b>,<b>
							E </b>and<b> F</b>) of the first worksheet named <b>Line</b>. The first column
						presents different regions where as the second, third, fourth, fifth and sixth columns
						represent the sales data representing values involving different years (<b>2002</b>
						<b>-</b> <b>2006</b>). The demo creates a Line chart representing <b>Sales By Region
							For Years</b> into the worksheet based on the different sales values of different
						regions in different years. In the example, you are provided a sample snapshot of
						the chart, a few controls including five drop down lists which represent chart type
						(Line and LineWithDataMarker), marker style (Square, Triangle, Diamond, Circle,
						Dash, Dot, None etc.), marker background color, marker foreground color, marker
						size and a command button labeled <b>Create Report</b> to create and exercise the
						chart based on your selection from the drop down lists. You are allowed to either
						open the resultant excel file(s) into MS Excel or save directly to your disk to
						check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/linecharts/stacked-line.aspx">Stacked Line</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Stacked Line chart that displays the trend of the contribution
						of each value over time or categories. This Demo demonstrates how to create a <b>Stacked</b>
						<b>Line</b> chart in a worksheet. This type of chart displays the trend of the contribution
						of each value over time or categories. It is also available with markers displayed
						at each data value. <b>Aspose.Cells</b> component supports all the standard and
						custom charts including <b>Stacked</b> <b>Line</b> chart to help you display data
						in more meaningful ways. The component can create the chart into the worksheet in
						a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs the source data related chart into the first six columns (<b>A</b>, <b>B</b>,<b>
							C</b>,<b> D</b>,<b> E </b>and<b> F</b>) of the first worksheet named <b>Data</b>.
						The first column represents different regions where as the second, third, fourth,
						fifth and sixth columns represent the sales data representing values involving different
						years (<b>2002</b> <b>-</b> <b>2006</b>). The demo creates a stacked line chart
						representing <b>Sales By Region For Years</b> into the second worksheet named <b>Chart</b>
						based on the different sales values of different regions in different years in the
						first worksheet. In the example, you are provided a sample snapshot of the chart,
						a drop down list which represents whether you want to create the chart with data
						markers and a command button labeled <b>Create Report</b> to create and exercise
						the chart based on your selection from the drop down list. You can either open the
						resultant excel file(s) into MS Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/linecharts/percent-stacked-line.aspx">100% Stacked Line</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						shows how to create 100% Stacked Line chart displays the trend of the percentage
						each value contributes over time or categories. This Demo demonstrates how to create
						a <b>100% Stacked Line</b> chart in a worksheet. This type of chart displays the
						trend of the percentage each value contributes over time or categories. It is also
						available with markers displayed at each data value. <b>Aspose.Cells</b> is a powerful
						component, which supports all the standard and custom charts to help you display
						data in more meaningful ways. You may create many kinds of charts including <b>100%
							Stacked Line</b>. The component can create the chart into the worksheet in a
						workbook using the simplest APIs with ease. The demo creates a workbook first and
						inputs some chart related data into the first five columns (<b>A</b>, <b>B</b>,<b> C</b>,<b>
							D </b>and<b> E</b>) of the first worksheet named <b>Data</b>. The first column
						represents the product names (<b>Product1</b>, <b>Product2</b> and <b>Product3</b>)
						where as the second, third, fourth and fifth columns represent percentage values
						involving different quarters (<b>Qurarter1</b>, <b>Quarter2</b>, <b>Quarter3</b>
						and <b>Quarter4</b>). The demo creates a 100% stacked line chart representing <b>Product
							contribution to total sales</b> into the second worksheet named <b>Chart</b>
						based on the different product values related to different quarters in the first
						worksheet. In the example, you have been provided a sample snapshot of the chart,
						a drop down list that represents whether you want to create the chart with data
						markers and a command button labeled <b>Create Report</b> to create and exercise
						the chart based on the selection from the drop down list. You can either open the
						resultant excel file(s) into MS Excel or save directly to your disk to check the
						results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/linecharts/line-3d.aspx">3D Line</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create a 3-D Line chart. This Demo exhibits how to create a <b>Line</b>
						chart with a <strong>3-D</strong> visual effect in a worksheet. <b>Aspose.Cells</b>
						component supports all the standard and custom charts including <b>3-D</b> <b>Line</b>
						chart to help you display data in more meaningful ways. The component can create
						the chart into the worksheet in a workbook using the simplest APIs with ease. The
						demo creates a workbook first and inputs the chart source data into the first six
						columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D</b>,<b> E </b>and<b> F</b>) of the first
						worksheet named <b>3D Line</b>. The first column presents different regions where
						as the second, third, fourth, fifth and sixth columns represent the sales data representing
						values involving different years (<b>2002</b> <b>-</b> <b>2006</b>). The demo creates
						a 3-D Line chart representing <b>Sales By Region </b>into the worksheet based on
						the different sales values of different regions in different years. In the example,
						you are provided a sample snapshot of the chart, a few controls including four drop
						down lists which represent major tick mark type (None, Inside, Outside and Cross)
						and minor tick mark type (None, Inside, Outside and Cross) for values, value labels
						rotation angle and category labels rotation angle and a command button labeled <b>Create
							Report</b> to create and exercise the chart based on your selection from the
						drop down lists. You are allowed to either open the resultant excel file(s) into
						MS Excel or save directly to your disk to check the results.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_4')">
				<img alt="Show" src="Image/blueup.gif" border="0">Pie Charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_4">
			<p class="componentDescriptionTxt">
				A pie chart shows the size of items that make up a data series ,proportional to
				the sum of the items. It always shows only one data series and is useful when you
				want to emphasize a significant element in the data. Pie charts have the following
				chart sub-types:</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/piecharts/pie.aspx">Pie</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create a Pie chart that displays the contribution of each value to
						a total. This Demo exhibits how to create a <b>Pie</b> chart with 2-D and 3-D visual
						effects in a worksheet. This type of chart displays the contribution of each value
						to a total. It is also available with a 3-D visual effect. <b>Aspose.Cells</b> is
						a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>Pie</b>. The component can create the chart into the worksheet in a workbook
						using the simplest APIs with ease. The demo creates a workbook first and inputs
						the chart source data into the first two columns (<b>A </b>and<b> B</b>) of the
						first worksheet. The first column represents the category data (<b>Region</b>) where
						as the second column represents the sales data representing values. The demo creates
						a pie chart representing <b>Sales By</b> <b>Region</b> into the worksheet named
						<b>Pie </b>based on the different sale values related to different regions. In the
						example, you are provided a sample snapshot of the chart, two drop down lists which
						represent first slice angle (90, 80, 180 and 360) and data label position (Center,
						InsideBase, InsideEnd and OutSideEnd), a check box that represents whether you want
						to create the chart with 3-D flavor and a command button labeled <b>Create Report</b>
						to create and exercise the chart using your desired inputs. You can either open
						the resultant excel file(s) into MS Excel or save directly to your disk to check
						the results.</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/piecharts/exploded-pie.aspx">Exploded Pie</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create an Exploded Pie chart that displays the contribution of each
						value to a total while emphasizing individual values. This Demo exhibits how to
						create an <b>Exploded</b> <b>Pie</b> chart with 2-D and 3-D visual effects in a
						worksheet. This type of chart displays the contribution of each value to a total
						while emphasizing individual values. It is also available with a 3-D visual effect.
						<b>Aspose.Cells</b> is a powerful component, which supports all the standard and
						custom charts to help you display data in more meaningful ways. You may create many
						kinds of charts including <b>Exploded</b> <b>Pie</b>. The component can create the
						chart into the worksheet in a workbook using the simplest APIs with ease. The demo
						creates a workbook first and inputs the chart source data into the first two columns
						(<b>A </b>and<b> B</b>) of the first worksheet named <b>Data</b>. The first column
						represents the category data (<b>Region</b>) where as the second column represents
						the sales data representing values. The demo creates an exploded pie chart representing
						<b>Sales By</b> <b>Region</b> into the second worksheet named <b>Chart </b>based
						on the different sale values related to different regions in the first worksheet.
						In the example, you are provided a sample snapshot of the chart, a check box that
						represents whether you want to create the chart with 3-D flavor and a command button
						labeled <b>Create Report</b> to create and exercise the chart using your desired
						inputs. You can either open the resultant excel file(s) into MS Excel or save directly
						to your disk to check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/piecharts/pie-of-pie.aspx">Pie of Pie</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Pie of Pie chart with user-defined values extracted and combined
						into a second pie. This Demo exhibits how to create a <b>Pie of Pie</b> chart in
						a worksheet. This is a pie chart with user-defined values extracted and combined
						into a second pie. For example, to make small slices easier to see, you can group
						them together as one item in a pie chart and then break down that item in a smaller
						pie next to the main chart. <b>Aspose.Cells</b> is a powerful component, which supports
						all the standard and custom charts to help you display data in more meaningful ways.
						You may create many kinds of charts including <b>Pie of Pie</b>. The component can
						create the chart into the worksheet in a workbook using the simplest APIs with ease.
						The demo creates a workbook first and inputs the chart source data into the first
						two columns (<b>A </b>and<b> B</b>) of the first worksheet named <b>Data</b>. The
						first column represents the category data (<b>Region</b>) where as the second column
						represents the sales data representing values. The demo creates a pie of pie chart
						representing <b>Sales By</b> <b>Region</b> into the second worksheet named <b>Chart
						</b>based on the different sale values related to different regions in the first
						worksheet. In the example, you are provided a sample snapshot of the chart and a
						command button labeled <b>Create Report</b> to create the chart. You can either
						open the resultant excel file into MS Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/piecharts/bar-of-pie.aspx">Bar of Pie</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Bar of Pie chart with user-defined values extracted and combined
						into a stacked bar. This Demo exhibits how to create a <b>Bar of Pie</b> chart in
						a worksheet. This is a pie chart with user-defined values extracted and combined
						into a stacked bar. For example, to make small slices easier to see, you can group
						them together as one item in a pie chart and then break down that item in a smaller
						bar next to the main chart. <b>Aspose.Cells</b> is a powerful component, which supports
						all the standard and custom charts to help you display data in more meaningful ways.
						You may create many kinds of charts including <b>Bar of Pie</b>. The component can
						create the chart into the worksheet in a workbook using the simplest APIs with ease.
						The demo creates a workbook first and inputs the chart source data into the first
						two columns (<b>A </b>and<b> B</b>) of the first worksheet named <b>BarofPie</b>.
						The first column represents the category data (<b>Region</b>) where as the second
						column represents the sales data representing values. The demo creates a bar of
						pie chart representing <b>Sales By</b> <b>Region</b> into the worksheet based on
						the different sale values related to different regions. In the example, you are
						provided a sample snapshot of the chart and a command button labeled <b>Create Report</b>
						to create the chart. You can either open the resultant excel file into MS Excel
						or save directly to your disk.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_5')">
				<img alt="Show" src="Image/blueup.gif" border="0">XY (Scatter) charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_5">
			<p class="componentDescriptionTxt">
				An xy (scatter) chart shows the relationships among the numeric values in several,or
				plots two groups of numbers as one series of xy coordinates.</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/scattercharts/scatter.aspx">Scatter</a>
					</p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Scatter chart based on uneven intervals (or clusters) of two
						sets of data. This Demo demonstrates how to create a <b>Scatter</b> chart in a worksheet.
						This type of chart compares pairs of values. The example creates a scatter chart
						which shows uneven intervals (or clusters) of two sets of data. When you arrange
						your data for a scatter chart, place x values in one row or column, and then enter
						corresponding y values in the adjacent rows or columns. <b>Aspose.Cells</b> is a
						powerful component, which supports all the standard and custom charts to help you
						display data in more meaningful ways. You may create many kinds of charts including
						<b>Scatter</b> chart. The component can create the chart into the worksheet in a
						workbook using the simplest APIs with ease. The demo creates a workbook first and
						inputs some chart related data into the first two columns (<b>A </b>and<b> B</b>)
						of the first worksheet named <b>Scatter</b>. The first column provides <b>Daily Rainfall</b>
						that represents the x values where as the second column denotes <b>Particulate</b>
						that represents the y values. The demo creates a scatter chart representing <b>Particulate
							Levels in Rainfall</b> into the worksheet based on the x and y values. In the
						example, you are provided a sample snapshot of the chart and a command button labeled
						<b>Create Report</b> to create and exercise the chart. You can either open the resultant
						excel file into MS Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/scattercharts/scatter-connected-by-lines.aspx">Scatter with Data Points
							Connected by Lines</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Scatter chart with or without straight or smoothed connecting
						lines between data points. This Demo exhibits how to create a <b>Scatter</b> chart
						connected by lines / curves with or without data markers in a worksheet. mso-bidi-font-family:
						'Times New Roman'">This type of chart compares pairs of values. This type of chart
						can be displayed with or without straight or smoothed connecting lines between data
						points. These lines can be displayed with or without markers. mso-fareast-font-family:
						'Times New Roman'; mso-bidi-font-family: 'Times New Roman'">The example creates
						a scatter chart with connected lines / curves with or without data markers which
						shows uneven intervals (or clusters) of two sets of data. When you arrange your
						data for a scatter chart, place x values in one row or column, and then enter corresponding
						y values in the adjacent rows or columns. <b>Aspose.Cells</b> is a powerful component,
						which supports all the standard and custom charts to help you display data in more
						meaningful ways. You may create many kinds of charts including all types of <b>Scatter</b>
						chart. The component can create the chart into the worksheet in a workbook using
						the simplest APIs with ease. The demo creates a workbook first and inputs chart
						related source data into the first two columns (<b>A </b>and<b> B</b>) of the first
						worksheet named <b>Data</b>. The first column provides <b>Daily Rainfall</b> that
						represents the x values where as the second column denotes <b>Particulate</b> that
						represents the y values. The demo creates a scatter chart representing <b>Particulate
							Levels in Rainfall</b> into second worksheet named <b>Chart</b> based on the
						x and y values in the first worksheet. In the example, you are provided a sample
						snapshot of the chart, a drop down list which represents different types of scatter
						chart and a command button labeled <b>Create Report</b> to create and exercise the
						chart based on the type you selected. You are allowed to either open the resultant
						excel file(s) into MS Excel or save directly to your disk.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_6')">
				<img alt="Show" src="Image/blueup.gif" border="0">Area charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_6">
			<p class="componentDescriptionTxt">
				An area chart emphasizes the magnitude of change over time. Area charts have the
				following chart sub-types:
			</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/areacharts/area.aspx">Area</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Area chart that displays the trend of values over time or categories.
						This Demo exhibits how to create <b>Area</b> chart with 2-D and 3-D flavors in a
						worksheet. This type of chart displays the trend of values over time or categories.
						It is also available with a 3-D visual effect. By displaying the sum of the plotted
						values, an area chart also shows the relationship of parts to a whole. For example,
						the following area chart emphasizes increased sales in different regions and illustrates
						the contribution of each country to total sales. <b>Aspose.Cells</b> component supports
						all the standard and custom charts including <b>Area</b> chart to help you display
						data in more meaningful ways. The component can create the chart into the worksheet
						in a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs the source data related chart into the first six columns (<b>A</b>, <b>B</b>,<b>
							C</b>,<b> D</b>,<b> E </b>and<b> F</b>) of the first worksheet named <b>Area</b>.
						The first column represents the category data (<b>Region</b>) where as the second,
						third, fourth, fifth and sixth columns represent the sales data representing values
						involving different years (<b>2002</b> <b>-</b> <b>2006</b>). The demo creates an
						area chart representing <b>Sales By</b> <b>Region</b> into the worksheet based on
						the different sales values related to different regions. In the example, you are
						provided a sample snapshot of the chart, a check box that represents whether you
						want to create the chart with 3-D visual effect and a command button labeled <b>Create
							Report</b> to create and exercise the chart. You can either open the resultant
						excel file(s) into MS Excel or save directly to your disk to check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/areacharts/stacked-area.aspx">Stacked Area</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Stacked Area chart that displays the trend of the contribution
						of each value over time or categories. This Demo exhibits how to create a <b>Stacked</b>
						<b>Area</b> chart with 2-D and 3-D flavors in a worksheet. This type of chart displays
						the trend of the contribution of each value over time or categories. It is also
						available with a 3-D visual effect. <b>Aspose.Cells</b> component supports all the
						standard and custom charts including <b>Stacked</b> <b>Area</b> chart to help you
						display data in more meaningful ways. The component can create the chart into the
						worksheet in a workbook using the simplest APIs with ease. The demo creates a workbook
						first and inputs the source data related chart into the first six columns (<b>A</b>,
						<b>B</b>,<b> C</b>,<b> D</b>,<b> E </b>and<b> F</b>) of the first worksheet named
						<b>Data</b>. The first column represents the category data (<b>Region</b>) where
						as the second, third, fourth, fifth and sixth columns represent the sales data representing
						values involving different years (<b>2002</b> <b>-</b> <b>2006</b>). The demo creates
						a stacked area chart representing <b>Total Sales</b> into the second worksheet named
						<b>Chart</b> based on the different sales values related to different regions in
						the first worksheet. In the example, you are provided a sample snapshot of the chart,
						a check box that represents whether you want to create the chart with 3-D visual
						effect and a command button labeled <b>Create Report</b> to create and exercise
						the chart. You can either open the resultant excel file(s) into MS Excel or save
						directly to your disk to check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/areacharts/percent-stacked-area.aspx">100% Stacked Area</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to creates 100% Stacked Area chart that displays the trend of the percentage
						each value contributes over time or categories. This Demo demonstrates how to create
						a <b>100% Stacked Area</b> chart with simple and 3-D visual effects in a worksheet.
						This chart type displays the trend of the percentage each value contributes over
						time or categories. It is also available with a 3-D visual effect. <b>Aspose.Cells</b>
						is a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>100% Stacked Area</b>. The component can create the chart into the worksheet
						in a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs some chart related data into the first five columns (<b>A</b>, <b>B</b>,<b>
							C</b>,<b> D </b>and<b> E</b>) of the first worksheet named <b>Data</b>. The
						first column represents the product names (<b>Product1</b>, <b>Product2</b> and
						<b>Product3</b>) where as the second, third, fourth and fifth columns represent
						percentage values involving different quarters (<b>Qurarter1</b>, <b>Quarter2</b>,
						<b>Quarter3</b> and <b>Quarter4</b>) which represent category data. The demo creates
						a 100% stacked area chart representing <b>Product contribution to total sales</b>
						into the second worksheet named <b>Chart</b> based on the different product values
						related to different quarters in the first worksheet. In the example, you are provided
						a sample snapshot of the chart, a check box that represents whether you want to
						create the chart with 3-D visual effect and a command button labeled <b>Create Report</b>
						to create and exercise the chart using your desired inputs. You can either open
						the resultant excel file(s) into MS Excel or save directly to your disk to check
						the results.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_7')">
				<img alt="Show" src="Image/blueup.gif" border="0">Doughnut charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_7">
			<p class="componentDescriptionTxt">
				Like a pie chart, a doughnut chart shows the relationship of parts to a whole; however,
				it can contain more than one data series. Doughnut charts have the following chart
				sub-types:
			</p>
			<ul class="GenericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/doughnutcharts/doughnut.aspx">Doughnut</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create a Doughnut chart that displays data in rings. This Demo describes
						how to create a <b>Doughnut</b> chart in a worksheet. This type of chart displays
						data in rings, where each ring represents a data series. For example, in the following
						chart, the inner ring represents 2005 fruit sales, and the outer ring represents
						2006 fruit sales. <b>Aspose.Cells</b> is a powerful component, which supports all
						the standard and custom charts to help you display data in more meaningful ways.
						You may create many kinds of charts including <b>Doughnut</b>. The component can
						create the chart into the worksheet in a workbook using the simplest APIs with ease.
						The demo creates a workbook first and inputs the chart source data into the first
						three columns (<b>A</b>,<b> B </b>and<b> C</b>) of the first worksheet named <b>Doughnut</b>.
						The first column represents the products (<b>Apple</b> and <b>Orange</b> ) where
						as the second and third columns represent the sales data involving different years
						representing values. The demo creates a doughnut chart representing <b>Fruit</b>
						<b>Sales by</b> <b>Region For Years</b> into the worksheet<b> </b>based on the different
						fruit sale values related to different years. In the example, you are provided a
						sample snapshot of the chart and a command button labeled <b>Create Report</b> to
						create the chart. You can either open the resultant excel file into MS Excel or
						save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/doughnutcharts/exploded-doughnut.aspx">Exploded Doughnut</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create an Exploded Doughnut chart which is like an Exploded Pie chart
						but it can contain more than one data series. This Demo describes how to create
						an <b>Exploded</b> <b>Doughnut</b> chart in a worksheet. This chart type is like
						an exploded pie chart, but it can contain more than one data series. This type of
						chart displays the contribution of each value to a total while emphasizing individual
						values. <b>Aspose.Cells</b> is a powerful component, which supports all the standard
						and custom charts to help you display data in more meaningful ways. You may create
						many kinds of charts including <b>Exploded</b> <b>Doughnut</b>. The component can
						create the chart into the worksheet in a workbook using the simplest APIs with ease.
						The demo creates a workbook first and inputs the chart source data into the first
						two columns (<b>A </b>and<b> B</b>) of the first worksheet named <b>Data</b>. The
						first column denotes the category data that represents the products (<b>Apple</b>
						and <b>Orange</b> ) where as the second column represents the yearly sales data
						that mentions values. The demo creates an exploded doughnut chart representing <b>Fruit</b>
						<b>Sales by</b> <b>Region For Years</b> into the second worksheet named<b> Chart </b>
						based on the different fruit sale values in the first worksheet. In the example,
						you are provided a sample snapshot of the chart and a command button labeled <b>Create
							Report</b> to create the chart. You can either open the resultant excel file
						into MS Excel or save directly to your disk.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_8')">
				<img alt="Show" src="Image/blueup.gif" border="0">Radar charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_8">
			<p class="componentDescriptionTxt">
				A radar chart compares the aggregate values of a number of data series. Radar charts
				have the following chart sub-types:
			</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/radarcharts/radar.aspx">Radar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Radar chart that displays changes in values relative to a center
						point. This Demo demonstrates how to create a <b>Radar</b> chart in a worksheet.
						This type of chart displays changes in values relative to a center point. It can
						be displayed with markers for each data point. For example, in the following radar
						chart, the data series that covers the most area, <b>Brand A</b>, represents the
						brand with the highest vitamin content. <b>Aspose.Cells</b> is a powerful component,
						which supports all the standard and custom charts to help you display data in more
						meaningful ways. You may create many kinds of charts including <b>Radar</b>. The
						component can create the chart into the worksheet in a workbook using the simplest
						APIs with ease. The demo creates a workbook first and inputs some chart related
						data into the first seven columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D</b>,<b> E</b>,<b>
							F </b>and<b> G</b>) of the first worksheet named <b>Radar</b>. The first column
						represents the different brands (<b>Brand A</b>, <b>Brand B</b> and <b>Brand C</b>)
						where as the second, third, fourth, fifth, sixth and seventh columns represent percentage
						values related to vitamins (<b>Vitamin A</b>, <b>Vitamin B1</b>, <b>Vitamin B2</b>,
						<b>Vitamin C</b>, <b>Vitamin D</b> and <b>Vitamin E</b>). The demo creates a radar
						chart titled <b>Nutritional Analysis</b> into the worksheet based on the different
						vitamin content values of different brands. In the example, you are provided a sample
						snapshot of the chart, a drop down list that represents the chart type (Radar, RadarWithDataMarkers)
						and a command button labeled <b>Create Report</b> to create and exercise the chart
						based on your selection from the drop down list. You can either open the resultant
						excel file(s) into MS Excel or save directly to your disk to check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/radarcharts/filled-radar.aspx">Filled Radar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create a Filled Radar chart. This Demo demonstrates how to create a
						<b>Filled</b> <b>Radar</b> chart in a worksheet. This type of chart displays changes
						in values relative to a center point. In this type of chart, the area covered by
						a data series is filled with a color. For example, in the following radar chart,
						the data series that covers the most area, <b>Brand A</b>, represents the brand
						with the highest vitamin content. <b>Aspose.Cells</b> is a powerful component, which
						supports all the standard and custom charts to help you display data in more meaningful
						ways. You may create many kinds of charts including <b>Filled</b> <b>Radar</b>.
						The component can create the chart into the worksheet in a workbook using the simplest
						APIs with ease. The demo creates a workbook first and inputs some chart related
						data into the first seven columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D</b>,<b> E</b>,<b>
							F </b>and<b> G</b>) of the first worksheet named <b>Data</b>. The first column
						represents the different brands (<b>Brand A</b>, <b>Brand B</b> and <b>Brand C</b>)
						where as the second, third, fourth, fifth, sixth and seventh columns represent percentage
						values related to vitamins (<b>Vitamin A</b>, <b>Vitamin B1</b>, <b>Vitamin B2</b>,
						<b>Vitamin C</b>, <b>Vitamin D</b> and <b>Vitamin E</b>). The demo creates a filled
						radar chart titled <b>Nutritional Analysis</b> into the second worksheet named <b>Chart</b>
						based on the different vitamin contents of different brands in the first worksheet.
						In the example, you are provided a sample snapshot of the chart and a command button
						labeled <b>Create Report</b> to create and exercise the chart. You can either open
						the resultant excel file into MS Excel or save directly to your disk.</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_9')">
				<img alt="Show" src="Image/blueup.gif" border="0">Surface charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_9">
			<p class="componentDescriptionTxt">
				A surface chart is useful when you want to find optimum combinations between two
				sets of data. As in a topographic map, colors and patterns indicate areas that are
				in the same range of values. Surface charts have the following chart sub-types:</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/surfacecharts/surface-3d.aspx">3D Surface</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create a 3-D Surface chart that displays trends in values across two
						dimensions in a continuous curve. This Demo demonstrates how to create a <b>3-D Surface</b>
						chart in a worksheet. This type of chart shows trends in values across two dimensions
						in a continuous curve. For example, the following surface chart shows the various
						combinations of temperature and time that result in the same measure of tensile
						strength. The colors in this chart represent specific ranges of values. Displayed
						without color, a 3-D surface chart is called a wireframe 3-D surface chart. <b>Aspose.Cells</b>
						is a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>3-D Surface</b>. The component can create the chart into the worksheet in a workbook
						using the simplest APIs with ease. The demo creates a workbook first and inputs
						some chart related data into the first six columns (<b>A</b>, <b>B</b>,<b> C</b>,<b>
							D</b>,<b> E </b>and<b> F</b>) of the first worksheet named <b>3D Surface</b>.
						The first column denotes a category, time (in seconds), related to different ranges
						(<b>0.2</b> - <b>1.0</b>) where as the second, third, fourth, fifth and sixth columns
						contain values of temperature series (<b>10</b>, <b>20</b>, <b>30</b>, <b>40</b>
						and <b>50</b>). The demo creates a 3-D surface chart titled <b>Tensile strength Measurements</b>
						into the worksheet based on time and temperature values. In the example, you are
						provided a sample snapshot of the chart, a drop down list that represents chart
						type (Surface3D and SurfaceWirframe3D) and a command button labeled <b>Create Report</b>
						to create and exercise the chart based on your selection from the drop down list.
						You may either open the resultant excel file(s) into MS Excel or save directly to
						your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/surfacecharts/contour.aspx">Contour</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Surface Contour chart. This Demo demonstrates how to create
						a <b>Surface Contour</b> chart in a worksheet. mso-fareast-font-family: 'Times New
						Roman'; This is a surface chart viewed from above, where colors represent specific
						ranges of values. Displayed without color, this chart type is called a Wireframe
						Contour. <b>Aspose.Cells</b> is a powerful component, which supports all the standard
						and custom charts to help you display data in more meaningful ways. You may create
						many kinds of charts including <b>Surface Contour</b>. The component can create
						the chart into the worksheet in a workbook using the simplest APIs with ease. The
						demo creates a workbook first and inputs some chart related data into the first
						six columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D</b>,<b> E </b>and<b> F</b>) of the
						first worksheet named <b>Data</b>. The first column denotes a category, time (in
						seconds), related to different ranges (<b>0.2</b> - <b>1.0</b>) where as the second,
						third, fourth, fifth and sixth columns contain values of temperature series (<b>10</b>,
						<b>20</b>, <b>30</b>, <b>40</b> and <b>50</b>). The demo creates a surface contour
						chart titled <b>Tensile strength Measurements</b> into the second worksheet named
						<b>Chart</b> based on time and temperature values in the first worksheet. In the
						example, you are provided a sample snapshot of the chart, a drop down list that
						represents chart type (SurfaceContour and SurfaceContourWirframe) and a command
						button labeled <b>Create Report</b> to create and exercise the chart based on your
						selection from the drop down list. You may either open the resultant excel file(s)
						into MS Excel or save directly to your disk.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_10')">
				<img alt="Show" src="Image/blueup.gif" border="0">Stock charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_10">
			<p class="componentDescriptionTxt">
				This chart type is most often used for stock price data, but can also be used for
				scientific data (for example, to indicate temperature changes). You must organize
				your data in the correct order to create stock charts. Stock charts have the following
				chart sub-types:
			</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/stockcharts/high-low-close.aspx">High-Low-Close</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create a High-Low-Close chart that displays stock prices. This Demo
						exhibits how to create a <b>High-Low-Close Stock</b> chart in a worksheet. mso-fareast-font-family:
						'Times New Roman'; The high-low-close chart is often used to illustrate stock prices.
						It requires three series of values in the following order (high, low, and then close).
						<b>Aspose.Cells</b> is a powerful component, which supports all the standard and
						custom charts to help you display data in more meaningful ways. You may create many
						kinds of charts including <b>High-Low-Close</b>. The component can create the chart
						into the worksheet in a workbook using the simplest APIs with ease. The demo creates
						a workbook first and inputs some chart related data into the first four columns
						(<b>A</b>, <b>B</b>,<b> C </b>and<b> D</b>) of the first worksheet named <b>HighLowClose</b>.
						The first column represents the companies (<b>Microsoft</b>, <b>Mutual Fund 1</b>
						and <b>Mutual Fund 2</b>), which denotes category data where as the second, third
						and fourth columns represent stock price values related to the scenarios (<b>High</b>,
						<b>Low</b> and <b>Close</b>). The demo creates a high-low-close stock chart representing
						<b>Stock chart</b> into the worksheet based on the different stock price values
						of the three states mentioned above. In the example, you have been provided a sample
						snapshot of the chart and a command button labeled <b>Create Report</b> to create
						the chart. You can either open the resultant excel file into MS Excel or save directly
						to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/stockcharts/open-high-low-close.aspx">Open-High-Low-Close</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Open-High-Low-Close chart that displays stock prices. This Demo
						exhibits how to create an <b>Open-High-Low-Close Stock</b> chart in a worksheet.
						The open-high-low-close chart is often used to illustrate stock prices. This type
						of chart requires four series of values in the correct order (open, high, low, and
						then close). <b>Aspose.Cells</b> is a powerful component, which supports all the
						standard and custom charts to help you display data in more meaningful ways. You
						may create many kinds of charts including <b>Open-High-Low-Close</b>. The component
						can create the chart into the worksheet in a workbook using the simplest APIs with
						ease. The demo creates a workbook first and inputs some chart related data into
						the first five columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D </b>and<b> E</b>) of
						the first worksheet named <b>Data</b>. The first column represents the companies
						(<b>Microsoft</b>, <b>Mutual Fund 1</b> and <b>Mutual Fund 2</b>), which denotes
						category data where as the second, third, fourth and fifth columns represent stock
						price values related to the scenarios (<b>Open</b>, <b>High</b>, <b>Low</b> and
						<b>Close</b>). The demo creates an open-high-low-close stock chart representing
						<b>Stock chart</b> into the second worksheet named <b>Chart</b> based on the different
						stock price values of the four states (mentioned above) in the first worksheet.
						In the example, you have been provided a sample snapshot of the chart and a command
						button labeled <b>Create Report</b> to create the chart. You are allowed to either
						open the resultant excel file into MS Excel or save directly to your disk.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_11')">
				<img alt="Show" src="Image/blueup.gif" border="0">Cylinder charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_11"> <!--border="0"-->
			<p class="componentDescriptionTxt">
				These chart types use cylinder data markers to lend a dramatic effect to column,
				bar, and 3-D column charts. Much like column and bar charts, cylinder charts have
				the following chart sub-types:</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/cylindercharts/cylinder-column.aspx">Column, Stacked Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Cylinder Column chart. This Demo exhibits how to create a <b>Cylinder
							Column</b> chart in a worksheet. The columns in these types of chart are represented
						by cylindrical shapes. You may create it with stacked flavor too. <b>Aspose.Cells</b>
						is a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>Cylinder Column</b>. The component can create the chart into the worksheet in
						a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs the chart source data into the first two columns (<b>A </b>and<b> B</b>)
						of the first worksheet named <b>Cylinder Column</b>. The first column represents
						the category data (<b>Year </b>spanned over 1996 - 2006) where as the second column
						represents the number of employees which denotes values in the chart. The demo creates
						a cylinder column chart representing <b>Number of Employees</b> into the worksheet
						based on the employee values related to different years. In the example, you are
						provided a sample snapshot of the chart, a drop down list which represents the chart
						type (Cylinder and CylinderStacked) and a command button labeled <b>Create Report</b>
						to create the chart based on your selection from the drop down list. The chart is
						created with a particular elevation and rotation angle, a specified depth and gap
						width in percentage b/w clusters as well. You can either open the resultant excel
						file(s) into MS Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/cylindercharts/cylinder-bar.aspx">Bar, Stacked Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Cylinder Bar chart. This Demo exhibits how to create a <b>Cylinder
							Bar</b> chart in a worksheet. The bars in these types of chart are represented
						by cylindrical shapes. You may create it with stacked flavor too. <b>Aspose.Cells</b>
						is a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>Cylinder Bar</b>. The component can create the chart into the worksheet in a
						workbook using the simplest APIs with ease. The demo creates a workbook first and
						inputs the chart source data into the first two columns (<b>A </b>and<b> B</b>)
						of the first worksheet named <b>Data</b>. The first column represents the category
						data (<b>Year </b>spanned over 1996 - 2006) where as the second column represents
						the number of employees which denotes values in the chart. The demo creates a cylinder
						bar chart titled <b>Number of Employees</b> into the second worksheet named <b>Chart</b>
						based on the employee values related to different years in the first worksheet.
						In the example, you are provided a sample snapshot of the chart, a drop down list
						which represents the chart type (CylindericalBar and CylindericalBarStacked) and
						a command button labeled <b>Create Report</b> to create the chart based on your
						selection from the drop down list. You can either open the resultant excel file(s)
						into MS Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/cylindercharts/cylinder-percent-stacked.aspx">100% Stacked Column
							or 100% Stacked Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Cylinder 100% Stacked Bar/Column chart that displays the trend
						of the percentage each value contributes over time or categories. This Demo demonstrates
						how to create a <b>Cylinder</b> <b>100% Stacked Bar/Column</b> chart in a worksheet.
						This type of chart compares the percentage each value contributes to a total across
						categories. The bars/columns in these types of chart are represented by cylindrical
						shapes.<b> Aspose.Cells</b> is a powerful component, which supports all the standard
						and custom charts to help you display data in more meaningful ways. You may create
						many kinds of charts including <b>Cylinder</b> <b>100% Stacked</b>. The component
						can create the chart into the worksheet in a workbook using the simplest APIs with
						ease. The demo creates a workbook first and inputs some chart related data into
						the first five columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D </b>and<b> E</b>) of
						the first worksheet named <b>Data</b>. The first column represents the product names
						(<b>Product1</b>, <b>Product2</b> and <b>Product3</b>) where as the second, third,
						fourth and fifth columns represent percentage values involving different quarters
						(<b>Qurarter1</b>, <b>Quarter2</b>, <b>Quarter3</b> and <b>Quarter4</b>). The demo
						creates a cylinder 100% stacked chart titled <b>Product contribution to total sales</b>
						into the second worksheet named <b>Chart</b> based on the different product values
						related to different quarters in the first worksheet. In the example, you are provided
						a sample snapshot of the chart, a drop down list that represents the chart type
						(Cylinder100PercentStacked and CylindericalBar100PercentStacked) and a command button
						labeled <b>Create Report</b> to create and exercise the chart based on your selection
						from the drop down list. You can either open the resultant excel file(s) into MS
						Excel or save directly to your disk to check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/cylindercharts/cylinderical-column-3d.aspx">3-D Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Cylinder 3-D Column chart. This Demo exhibits how to create
						a <b>Cylinder 3-D Column</b> chart in a worksheet. mso-fareast-font-family: 'Times
						New Roman'; The 3-D columns in these types of chart are represented by cylindrical
						shapes. <b>Aspose.Cells</b> is a powerful component, which supports all the standard
						and custom charts to help you display data in more meaningful ways. You may create
						many kinds of charts including <b>Cylinder 3-D Column</b>. The component can create
						the chart into the worksheet in a workbook using the simplest APIs with ease. The
						demo creates a workbook first and inputs the chart source data into the first two
						columns (<b>A </b>and<b> B</b>) of the first worksheet named <b>Cylinderical Column3D</b>.
						The first column represents the category data (<b>Year </b>spanned over 1996 - 2006)
						where as the second column represents the number of employees which denotes values
						in the chart. The demo creates a cylinder 3-D column chart representing <b>Number of
							Employees</b> into the worksheet based on the employee values related to different
						years. In the example, you are provided a sample snapshot of the chart and a command
						button labeled <b>Create Report</b> to create the chart. The chart is created with
						a particular elevation and rotation angle, a specified depth and gap width in percentage
						b/w clusters as well. You can either open the resultant excel file into MS Excel
						or save directly to your disk.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_12')">
				<img alt="Show" src="Image/blueup.gif" border="0">Cone charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_12"> <!--border="0"-->
			<p class="componentDescriptionTxt">
				These chart types use cone data markers to lend a dramatic effect to column, bar,
				and 3-D column charts. Much like column and bar charts, cone charts have the following
				chart sub-types:</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/conecharts/cone-column.aspx">Column,Stacked Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Cone Column chart. This Demo exhibits how to create a <b>Cone Column</b>
						chart in a worksheet. The columns in these types of chart are represented by conical
						shapes. You may create it with stacked flavor too. <b>Aspose.Cells</b> is a powerful
						component, which supports all the standard and custom charts to help you display
						data in more meaningful ways. You may create many kinds of charts including <b>Cone
							Column</b>. The component can create the chart into the worksheet in a workbook
						using the simplest APIs with ease. The demo creates a workbook first and inputs
						the chart source data into the first two columns (<b>A </b>and<b> B</b>) of the
						first worksheet named <b>Cone Column</b>. The first column represents the category
						data (<b>Year </b>spanned over 1996 - 2006) where as the second column represents
						the number of employees which denotes values in the chart. The demo creates a cone
						column chart representing <b>Number of Employees</b> into the worksheet based on
						the employee values related to different years. In the example, you are provided
						a sample snapshot of the chart, a drop down list which represents the chart type
						(Cone and ConeStacked) and a command button labeled <b>Create Report</b> to create
						the chart based on your selection from the drop down list. The chart is created
						with a particular elevation and rotation angle with a specified depth and gap width
						in percentage b/w clusters. You can either open the resultant excel file(s) into
						MS Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/conecharts/cone-bar.aspx">Bar, Stacked Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Cone Bar chart. This Demo exhibits how to create a <b>Cone Bar</b>
						chart in a worksheet. The bars in these types of chart are represented by conical
						shapes. You may create it with stacked flavor too. <b>Aspose.Cells</b> is a powerful
						component, which supports all the standard and custom charts to help you display
						data in more meaningful ways. You may create many kinds of charts including <b>Cone
							Bar</b>. The component can create the chart into the worksheet in a workbook
						using the simplest APIs with ease. The demo creates a workbook first and inputs
						the chart source data into the first two columns (<b>A </b>and<b> B</b>) of the
						first worksheet named <b>Data</b>. The first column represents the category data
						(<b>Year </b>spanned 1996 - 2006) where as the second column represents the number
						of employees which denotes values in the chart. The demo creates a cone bar chart
						representing <b>Number of Employees</b> into the second worksheet named <b>Chart</b>
						based on the employee values related to different years in the first worksheet.
						In the example, you are provided a sample snapshot of the chart, a drop down list
						which represents the chart type (ConicalBar and ConicalBarStacked) and a command
						button labeled <b>Create Report</b> to create the chart based on your selection
						from the drop down list. You can either open the resultant excel file(s) into MS
						Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/conecharts/cone-percent-stacked.aspx">100% Stacked Column or 100%
							Stacked Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Cone 100% Stacked Bar/Column chart that displays the trend of
						the percentage each value contributes over time or categories. This Demo demonstrates
						how to create a <b>Cone</b> <b>100% Stacked Bar/Column</b> chart in a worksheet.
						This type of chart compares the percentage each value contributes to a total across
						categories. The bars/columns in these types of chart are represented by conical
						shapes. <b>Aspose.Cells</b> is a powerful component, which supports all the standard
						and custom charts to help you display data in more meaningful ways. You may create
						many kinds of charts including <b>Cone</b> <b>100% Stacked</b>. The component can
						create the chart into the worksheet in a workbook using the simplest APIs with ease.
						The demo creates a workbook first and inputs some chart related data into the first
						five columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D </b>and<b> E</b>) of the first
						worksheet named <b>Data</b>. The first column represents the product names (<b>Product1</b>,
						<b>Product2</b> and <b>Product3</b>) where as the second, third, fourth and fifth
						columns represent percentage values involving different quarters (<b>Qurarter1</b>,
						<b>Quarter2</b>, <b>Quarter3</b> and <b>Quarter4</b>). The demo creates a cone 100%
						stacked chart titled <b>Product contribution to total sales</b> into the second
						worksheet named <b>Chart</b> based on the different product values related to different
						quarters in the first worksheet. In the example, you are provided a sample snapshot
						of the chart, a drop down list that represents the chart type (Cone100PercentStacked
						and ConicalBar100PercentStacked) and a command button labeled <b>Create Report</b>
						to create and exercise the chart based on your selection from the drop down list.
						You can either open the resultant excel file(s) into MS Excel or save directly to
						your disk to check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/conecharts/cone-column-3d.aspx">3-D Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Cone 3-D Column chart. This Demo exhibits how to create a <b>Cone
							3-D Column</b> chart in a worksheet. mso-fareast-font-family: 'Times New Roman';
						The 3-D columns in these types of chart are represented by conical shapes. <b>Aspose.Cells</b>
						is a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>Cone 3-D Column</b>. The component can create the chart into the worksheet in
						a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs the chart source data into the first two columns (<b>A </b>and<b> B</b>)
						of the first worksheet named <b>Cone Column3D</b>. The first column represents the
						category data (<b>Year </b>spanned over 1996 - 2006) where as the second column
						represents the number of employees which denotes values in the chart. The demo creates
						a cone 3-D column chart representing <b>Number of Employees</b> into the worksheet
						based on the employee values related to different years. In the example, you are
						provided a sample snapshot of the chart and a command button labeled <b>Create Report</b>
						to create the chart. The chart is created with a particular elevation and rotation
						angle with a specified depth and gap width in percentage b/w clusters. You can either
						open the resultant excel file into MS Excel or save directly to your disk.
					</p>
				</li>
			</ul>
		</div>
		<p>
			<a class="DropDown" href="javascript:ToggleDiv('divExpCollAsst_13')">
				<img alt="Show" src="Image/blueup.gif" border="0">Pyramid charts</a></p>
		<div class="demoDivPanelCollapsed" id="divExpCollAsst_13"> <!--border="0"-->
			<p class="componentDescriptionTxt">
				These chart types use pyramid data markers to lend a dramatic effect to column,
				bar, and 3-D column charts. Much like column and bar charts, pyramid charts have
				the following chart sub-types:</p>
			<ul class="genericList">
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/pyramidcharts/pyramid-column.aspx">Column,Stacked Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Pyramid Column chart. This Demo exhibits how to create a <b>Pyramid
							Column</b> chart in a worksheet. The columns in these types of chart are represented
						by pyramid shapes. You may create it with stacked flavor too. <b>Aspose.Cells</b>
						is a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>Pyramid Column</b>. The component can create the chart into the worksheet in
						a workbook using the simplest APIs with ease. The demo creates a workbook first
						and inputs the chart source data into the first two columns (<b>A </b>and<b> B</b>)
						of the first worksheet named <b>Pyramid Column</b>. The first column represents
						the category data (<b>Year </b>spanned over 1996 - 2006) where as the second column
						represents the number of employees which denotes values in the chart. The demo creates
						a pyramid column chart representing <b>Number of Employees</b> into the worksheet
						based on the employee values related to different years. In the example, you are
						provided a sample snapshot of the chart, a drop down list which represents the chart
						type (Pyramid and PyramidStacked) and a command button labeled <b>Create Report</b>
						to create the chart based on your selection from the drop down list. The chart is
						created with a particular elevation and rotation angle, a specified depth and gap
						width in percentage b/w clusters as well. You can either open the resultant excel
						file(s) into MS Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/pyramidcharts/pyramid-bar.aspx">Bar, Stacked Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Pyramid Bar chart. This Demo exhibits how to create a <b>Pyramid
							Bar</b> chart in a worksheet. The bars in these types of chart are represented
						by pyramid shapes. You may create it with stacked flavor too. <b>Aspose.Cells</b>
						is a powerful component, which supports all the standard and custom charts to help
						you display data in more meaningful ways. You may create many kinds of charts including
						<b>Pyramid Bar</b>. The component can create the chart into the worksheet in a workbook
						using the simplest APIs with ease. The demo creates a workbook first and inputs
						the chart source data into the first two columns (<b>A </b>and<b> B</b>) of the
						first worksheet named <b>Data</b>. The first column represents the category data
						(<b>Year </b>spanned over 1996 - 2006) where as the second column represents the
						number of employees which denotes values in the chart. The demo creates a pyramid
						bar chart representing <b>Number of Employees</b> into the second worksheet named
						<b>Chart</b> based on the employee values related to different years in the first
						worksheet. In the example, you are provided a sample snapshot of the chart, a drop
						down list which represents the chart type (PyramidBar and PyramidBarStacked) and
						a command button labeled <b>Create Report</b> to create the chart based on your
						selection from the drop down list. You can either open the resultant excel file(s)
						into MS Excel or save directly to your disk.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/pyramidcharts/pyramid-percent-stacked.aspx">100% Stacked Column or
							100% Stacked Bar</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Pyramid 100% Stacked Column/Bar chart that displays the trend
						of the percentage each value contributes over time or categories. This Demo exhibits
						how to create a <b>Pyramid</b> <b>100% Stacked Bar/Column</b> chart in a worksheet.
						This type of chart compares the percentage each value contributes to a total across
						categories. The bars/columns in these types of chart are represented by pyramid
						shapes. <b>Aspose.Cells</b> is a powerful component, which supports all the standard
						and custom charts to help you display data in more meaningful ways. You may create
						many kinds of charts including <b>Pyramid</b> <b>100% Stacked</b>. The component
						can create the chart into the worksheet in a workbook using the simplest APIs with
						ease. The demo creates a workbook first and inputs some chart related data into
						the first five columns (<b>A</b>, <b>B</b>,<b> C</b>,<b> D </b>and<b> E</b>) of
						the first worksheet named <b>Data</b>. The first column represents the product names
						(<b>Product1</b>, <b>Product2</b> and <b>Product3</b>) where as the second, third,
						fourth and fifth columns represent percentage values involving different quarters
						(<b>Qurarter1</b>, <b>Quarter2</b>, <b>Quarter3</b> and <b>Quarter4</b>). The demo
						creates a pyramid 100% stacked chart titled <b>Product contribution to total sales</b>
						into the second worksheet named <b>Chart</b> based on the different product values
						related to different quarters in the first worksheet. In the example, you are provided
						a sample snapshot of the chart, a drop down list that represents the chart type
						(Pyramid100PercentStacked and PyramidBar100PercentStacked) and a command button
						labeled <b>Create Report</b> to create and exercise the chart based on your selection
						from the drop down list. You can either open the resultant excel file(s) into MS
						Excel or save directly to your disk to check the results.
					</p>
				</li>
				<li class="genericList">
					<p class="productTitle">
						<a href="ChartTypes/pyramidcharts/pyramid-column-3d.aspx">3-D Column</a></p>
					<p class="componentDescriptionCaption">
						Description</p>
					<p class="componentDescriptionTxt">
						Shows how to create Pyramid 3-D Column chart. This Demo exhibits how to create a
						<b>Pyramid 3-D Column</b> chart in a worksheet. mso-fareast-font-family: 'Times
						New Roman'; The 3-D columns in these types of chart are represented by pyramid shapes.
						<b>Aspose.Cells</b> is a powerful component, which supports all the standard and
						custom charts to help you display data in more meaningful ways. You may create many
						kinds of charts including <b>Pyramid 3-D Column</b>. The component can create the
						chart into the worksheet in a workbook using the simplest APIs with ease. The demo
						creates a workbook first and inputs the chart source data into the first two columns
						(<b>A </b>and<b> B</b>) of the first worksheet named <b>Pyramid Column3D</b>. The
						first column represents the category data (<b>Year </b>spanned over 1996 - 2006)
						where as the second column represents the number of employees which denotes values
						in the chart. The demo creates a pyramid 3-D column chart titled <b>Number of Employees</b>
						into the worksheet based on the employee values related to different years. In the
						example, you are provided a sample snapshot of the chart and a command button labeled
						<b>Create Report</b> to create the chart. The chart is created with a particular
						elevation and rotation angle, a specified depth and gap width in percentage b/w
						clusters as well. You can either open the resultant excel file into MS Excel or
						save directly to your disk.
					</p>
				</li>
			</ul>
		</div>
	</div>
</asp:Content>
