<%@ Page Language="vb" CodeBehind="Northwind.aspx.vb" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.NorthwindPage"
	MasterPageFile="~/tpl/Demo.Master" Title="Northwind Demos - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<p class="productTitle">
		Northwind Demos - <a href="default.aspx">Aspose.Cells Demos</a></p>
	<p class="componentDescriptionTxt">
		The demos illustrate how to create the most diverse reports similar to those contained
		in Microsoft Access <b>Northwind</b> Sample Database using <b>Aspose.Cells</b> component.</p>
	<p class="componentDescriptionTxt">
		<b>Report</b> is an effective way to present your data in a printed format. You
		may refer it to a printed output of the data in the database. <b>Aspose.Cells</b>
		component gives you the agility to report your data in a variety of ways. A report
		can show all or only some of the data of a record, and it can be based on either
		a table or a query. The flexibility to customize reports and to organize the data,
		is provided by the component. <b>Aspose.Cells</b> component is fully functional
		for creating all types of reports. You may customize the size and appearance of
		everything on a report. You can display the information the way you want to see
		it. Normally, we use ADO .Net components (Connection, Command, DataAdapter, DataSet,
		DataTable etc.) to establish a connection with the data source, retrieve data from
		the database and fill a dataset or datatable(s) with it, and then generate a report
		in the worksheet. <b>Aspose.Cells</b> component offers some rich APIs for creating,
		formatting and managing reports. For example, you may furnish the cells in your
		report sheet. You can style your sheets like changing background and foreground
		color of the cells, shaping borders around your cells and adjusting data alignment.
		Sometimes, you might want to put a title in a single cell that spans the top of
		your report. You can easily merge the cells into a single cell within a specified
		range of the cells in your report. Moreover you can set a font to a range of cells,
		you may change the color of the font too. Additionally, you can input any type of
		formula to the cells and calculate the results.
	</p>
	<p class="componentDescriptionTxt">
		The demos create <b>14</b> reports based on "Northwind.mdb" database stored in the
		"Database" Folder in your Application Directory. The examples also utilize a template
		excel file. The <b>Northwind</b> database contains the sales data for a fictitious
		company called <b>Northwind Traders</b>, which imports and exports specialty foods
		from around the world. By viewing the database objects like tables, queries and
		reports, you can develop ideas for your own database application. <b>Aspose.Cells</b>
		component makes use of <b>Northwind</b> data to demonstrate the practice session
		in designing queries to produce the most diverse reports, since it contains enough
		records to produce meaningful results.Each report is well organized, formatted and
		presents data in an efficient manner. Click on the different hyperlinks below, each
		hyperlink represents a report. You are allowed to either open the resultant reports
		into your browser or save directly to your disk.</p>
	<ul class="genericList">
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/product-list-form.aspx">Alphabetical List of Products</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints products grouped by first letter of name.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/catalog-form.aspx">Catalog</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints a catalog of products. Has two-page report header; uses photos for each category;
				starts each category on a new page; keeps all records for a category on same page;
				prints an order form in the report footer on a separate page.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/catalog-subreport-form.aspx">Catalog Subreport</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints subreport for Catalog report.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/customer-labels-form.aspx">Customer Labels</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints customers' company names and addresses on 3-up labels.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/employee-sales-form.aspx">Employee Sales by Country</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints sales grouped by country and employee. Calculates subtotals, grand total,
				percents; prints range on report; prints message when employee's total reaches goal;
				a sheet per country.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/invoice-form.aspx">Invoice</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints each invoice on a separate page. This is the only demo to create the Excel
				spreadsheet completely via API.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/products-by-category-form.aspx">Products by Category</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints products by category. Has 3 columns per page; starts each category in new
				column.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/sales-by-category-form.aspx">Sales by Category</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints sales for 1994 by category. Shows sales in a subreport and in a chart on
				the main report.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/sales-by-category-subreport-form.aspx">Sales by Category Subreport</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints subreport for Sales by Category report.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/sales-by-year-form.aspx">Sales by Year</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints quarter totals in subreport in group header; optionally prints detail records.
				Displays page header on pages that don't have a group header.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/sales-by-year-subreport-form.aspx">Sales by Year Subreport</a>
			</p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints subreport for Sales by Year report.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/sales-totals-form.aspx">Sales Totals by Amount</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints sales in descending order by amount. Prints top 10 customers on first page;
				prints page total in page footer.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/summary-by-quarter-form.aspx">Summary of Sales by Quarter</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints summary report showing sales from multiple years for each quarter.
			</p>
		</li>
		<li class="genericList">
			<p class="productTitle">
				<a href="northwind/summary-by-year-form.aspx">Summary of Sales by Year</a></p>
			<p class="componentDescriptionCaption">
				Description</p>
			<p class="componentDescriptionTxt">
				Prints summary report showing quarterly sales for each year.
			</p>
		</li>
	</ul>
</asp:Content>
