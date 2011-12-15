<%@ Page Language="c#" CodeBehind="business-reports.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.BusinessReports" MasterPageFile="~/tpl/Demo.Master"
    Title="Business Reports - Aspose.Cells Demos" %>

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
                        Bisiness Reports - <a href="default.aspx">Aspose.Cells Demos</a>
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
            The demo illustrates how to create <strong>Financial Plan</strong> and <strong>Cost
                Pareto</strong> chart reports similar to business reports template files contained
            in <strong>Microsoft</strong> website.</p>
        <p>
            The <b>Business Plan</b> template can help your company assess a particular business
            opportunity and its potential impact within your industry. This template is structured
            as a formal plan that makes it easy for you to review such data as general industry
            information, your company's organizational structure, direct competition, and potential
            customers.</p>
        <p>
            The Five Year Plan (Service Industry) template has been designed to help a financial
            services company --- such as a small bank, mortgage broker, or savings-and-loan
            company. The high-level financial plan defines your financial model and pricing
            assumptions with all other important aspects. This plan also includes expected annual
            sales and profits for the next five years.
        </p>
        <p>
            When you click the <b>Financial Plan</b> hyperlink, a template file is used which
            has five worksheets named <b>Model Inputs</b>, <b>Profit and Loss</b>, <b>Balance Sheet</b>,
            <b>Cash Flow </b>and<b> Loan Payment Calculator</b>. All the sheets are filled with
            data with all its formatting using APIs of <b>Aspose.Cells</b> component to produce
            a complete 5 years Financial Plan for any Corporate. When you click the <b>Cost Pareto</b>
            hyperlink, here the demo utilizes an xml file, reads the xml data from the file
            to generate two reports named <b>Cost Data</b> and <b>Pareto Chart </b>each representing
            a worksheet in the workbook. The first report calculates the annual cost of different
            cost centers while the second worksheet presents a column chart based on the first
            worksheet data. You are allowed to either open the resultant excel file(s) into
            your browser or save directly to your disk.</p>
        <ul>
            <li class="genericList">
                <p class="productTitle">
                    <a href="businessreports/financial-plan.aspx">Financial Plan</a></p>
                <p class="componentDescriptionCaption">
                    Description</p>
                <p class="componentDescriptionTxt">
                    Prints a 5-Year Financial Plan.
                </p>
            </li>
            <li class="genericList">
                <p class="productTitle">
                    <a href="businessreports/cost-pareto.aspx">Cost Pareto</a></p>
                <p class="componentDescriptionCaption">
                    Description</p>
                <p class="componentDescriptionTxt">
                    Prints a cost pareto chart report.
                </p>
            </li>
        </ul>
    </div>
</asp:Content>
