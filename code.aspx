<%@ Register TagPrefix="DM" Namespace="LibraryDM"%>
<%@ Register TagPrefix="DM" TagName="Top" Src="~/custom/bHTMLTop.ascx" %>
<%@ Register TagPrefix="DM" TagName="Bottom" Src="~/custom/HTMLBottom.ascx" %>
<%@ Page Debug = True %>
<script runat="server">
	Dim Sales, SalesPersons, Categories, FiscalMonths, SalesPersonOptions As System.Data.DataTable
	Dim cmSdate, cySdate, eDate, pymSdate, pySdate, pyeDate, currentMonth, currentMonthYearStartDate, currentMonthID, SalesPersonIDs, CategoryIDs
	Dim TodaySale, ISTodaySale, MTDSaleCY, ISMTDSaleCY, MTDMarginCY, ISMTDMarginCY, MTDSalePY, ISMTDSalePY, MTDMarginPY, ISMTDMarginPY, YTDSaleCY, ISYTDSaleCY, YTDMarginCY, ISYTDMarginCY, YTDSalePY, ISYTDSalePY, YTDMarginPY, ISYTDMarginPY, UnInvoicedSales, ISUnInvoicedSales, CurrentMonthBudget, ISCurrentMonthBudget, CurrentYearBudget, ISCurrentYearBudget As Double
	Dim TodaySaleTotal, MTDSaleCYTotal, MTDMarginCYTotal,	MTDSalePYTotal, MTDMarginPYTotal, YTDSaleCYTotal, YTDMarginCYTotal, YTDSalePYTotal, YTDMarginPYTotal, UnInvoicedSalesTotal, CurrentMonthBudgetTotal, CurrentYearBudgetTotal As Double
	Public Sub page_load()
		DMUser.CheckIn()
		IF ""& Request("q2") <> "" Then
			eDate = Request("q2")
			cmSdate = FirstDayOfCurrentYearMonth(eDate)
			cySdate = Sql.ReadField("SELECT YearStart FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'")
			pymSdate = FirstDayOfPreviousYearMonth(eDate)
			pySdate = Sql.ReadField("SELECT YearStart FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'")
			pySdate = FirstDayOfPreviousYear(pySdate)
			pyeDate = PreviousYearDateofSourceDate(eDate)
			q2.value = eDate
			currentMonth = DateTime.Parse(eDate).ToString("MMM")
			FiscalMonths = Sql.DataTable("SELECT row_number() over (order by DATEADD(MONTH, x.number, '10/01/2022')) AS MonthID, DATENAME(MONTH, DATEADD(MONTH, x.number, '"& Sql.ReadField("SELECT YearStart FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'") &"')) AS MonthName FROM master.dbo.spt_values x WHERE x.type = 'P' AND x.number <= DATEDIFF(MONTH, '"& Sql.ReadField("SELECT YearStart FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'") &"', '"& Sql.ReadField("SELECT YearEnd FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'") &"')")
			currentMonthYearStartDate = Sql.ReadField("SELECT YEAR(YearStart) FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'")
			For Each FiscalMonth IN FiscalMonths.rows
				IF DateTime.Parse(eDate).ToString("MMMM") = FiscalMonth("MonthName") Then
					currentMonthID = FiscalMonth("MonthID")
				End IF
			Next
		Else
			q2.value = Date.Today()
			eDate = Date.Today()
			cmSdate = FirstDayOfCurrentYearMonth(Date.Today())
			cySdate = Sql.ReadField("SELECT YearStart FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'")
			pymSdate = FirstDayOfPreviousYearMonth(Date.Today())
			pySdate = Sql.ReadField("SELECT YearStart FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'")
			pySdate = FirstDayOfPreviousYear(pySdate)
			pyeDate = PreviousYearDateofSourceDate(Date.Today())
			currentMonth = DateTime.Parse(eDate).ToString("MMM")
			FiscalMonths = Sql.DataTable("SELECT row_number() over (order by DATEADD(MONTH, x.number, '10/01/2022')) AS MonthID, DATENAME(MONTH, DATEADD(MONTH, x.number, '"& Sql.ReadField("SELECT YearStart FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'") &"')) AS MonthName FROM master.dbo.spt_values x WHERE x.type = 'P' AND x.number <= DATEDIFF(MONTH, '"& Sql.ReadField("SELECT YearStart FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'") &"', '"& Sql.ReadField("SELECT YearEnd FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'") &"')")
			currentMonthYearStartDate = Sql.ReadField("SELECT YEAR(YearStart) FROM YearEnd WHERE YearStart <= '" & eDate &"' AND YearEnd >= '" & eDate &"'")
			For Each FiscalMonth IN FiscalMonths.rows
				IF DateTime.Parse(eDate).ToString("MMMM") = FiscalMonth("MonthName") Then
					currentMonthID = FiscalMonth("MonthID")
				End IF
			Next
		End IF
		SalesPersonIDs = ""
		CategoryIDs = ""
		IF ""& Request("q5") <> "" AND  ""& Request("Type") = "" Then
			SalesPersonIDs = Request("q5")
		Else IF ""& Request("q5") <> "" AND  ""& Request("Type") <> "" Then
			CategoryIDs = Request("q5")
		End IF
		Sales = Sql.DataTable("SELECT * FROM R_MonthlySalesSummary WHERE TransactionDate BETWEEN '"& pySdate &"' AND '"& eDate &"'")
		IF ""& Request("Type") = "" Then
			SalesPersons = Sql.DataTable("SELECT DISTINCT ISNULL(Name,'') SalesPerson, ISNULL(InActive,'') InActive, ID FROM Party WHERE ISNULL(IsEmployee,'') = 'on' AND ISNULL(IsSalesRep,'') = 'on' ORDER BY ISNULL(Name,'')")
		Else
			Categories = Sql.DataTable("SELECT DISTINCT ISNULL(Name,'') Category, ID FROM Category WHERE ParentCategoryID IS NULL")
		End IF
	End Sub

	Public Function FirstDayOfCurrentYearMonth(ByVal sourceDate As DateTime) As DateTime    
		Return New DateTime(sourceDate.Year, sourceDate.Month, 1)
	End Function

	Public Function FirstDayOfPreviousYearMonth(ByVal sourceDate As DateTime) As DateTime    
		Return New DateTime(sourceDate.Year-1, sourceDate.Month, 1)
	End Function

	Public Function FirstDayOfPreviousYear(ByVal sourceDate As DateTime) As DateTime    
		Return New DateTime(sourceDate.Year-1, sourceDate.Month, 1)
	End Function

	Public Function PreviousYearDateofSourceDate(ByVal sourceDate As DateTime) As DateTime    
		Return New DateTime(sourceDate.Year-1, sourceDate.Month, sourceDate.Day)
	End Function
</script>
<DM:Top runat=server id="top" Title="SPS - Sales Comparison"/>
<link href='https://fonts.googleapis.com/css?family=Roboto' rel='stylesheet'>
<table width="100%">
	<tr>
		<td class='TitleTD' style="border-right: none;" colspan=22>
			<table>
				<tr>
					<td style='font-size: 20px; font-family: roboto; padding-right: 10px;'><%= IIF(""& Request("Type") = "","Sales Summary By Sales Rep as of","Sales Summary By Product Category as of") %></td>
					<td><form runat="server" style="height: 33px;"><DM:input ID='q2' runat="server" type='closeddate' IsPopupCalendar="True" width="100" style="height: 25px; width: 100px; font-size: 17px; font-family: roboto; padding-left: 5px; border: 1px solid #aaa;" onchange="changeDate(this.value)"/></form></td>
					<td style="padding-left: 20px">
						<% IF ""& Request("Type") = "" Then%>
							<select id="SalesPerson" multiple="multiple" style="width: 500px !important;">
								<% For Each SalesPerson In SalesPersons.rows
									DM.Write("<option "& IIF(SalesPersonIDs.contains(SalesPerson("ID")), "selected", "") &" value="& SalesPerson("ID") &">"& SalesPerson("SalesPerson") &"</option>")
								Next %>
							</select>
						<% Else %>
							<select id="Category" multiple="multiple" style="width: 500px !important;">
								<% For Each Category In Categories.rows
									DM.Write("<option "& IIF(CategoryIDs.contains(Category("ID")), "selected", "") &" value="& Category("ID") &">"& Category("Category") &"</option>")
								Next %>
							</select>
						<% End IF %>
					</td>
					<% IF ""& Request("Type") = "" Then%>
						<td style="padding-left: 12px"><button class="btn btn-inverse" style="font-family: roboto; height: 33px;" onclick="search(2)">Show Selected Sales People</button></td>
						<td style="padding-left: 12px"><button class="btn btn-inverse" style="font-family: roboto; height: 33px;" onclick="search(0)">Hide Selected Sales People</button></td>
					<% Else %>
						<td style="padding-left: 12px"><button class="btn btn-inverse" style="font-family: roboto; height: 33px;" onclick="search(2)">Show Selected Categories</button></td>
						<td style="padding-left: 12px"><button class="btn btn-inverse" style="font-family: roboto; height: 33px;" onclick="search(0)">Hide Selected Categories</button></td>
					<% End IF %>
					<td style="padding-left: 12px"><button class="btn btn-inverse" style="font-family: roboto; height: 33px;" onclick="search(1)">Show All</button></td>
					<td style="padding-left: 12px"><button class="btn btn-inverse" style="font-family: roboto; height: 33px;" onclick="exportTableToExcel('Sales')"/>Export to Excel</button></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<table width="100%" ID="Sales">
	<tr>
		<td class='HeaderTD' style='text-align: left; background-color: #d9e2f3; padding-left: 7px;' nowrap><%= IIF(""& Request("Type") = "", "Sales Rep", "Category") %></td>
		<td class='HeaderTD' style='background-color: #d9e2f3;' nowrap>Sales as on <%=q2.value%></td>
		<td class='HeaderTD' style='background-color: #d9e2f3;' nowrap>MTD Sale (CY)</td>
		<td class='HeaderTD' style='background-color: #d9e2f3;' nowrap>MTD Margin (CY)</td>
		<td class='HeaderTD' style='background-color: #d9e2f3;' nowrap>MTD Margin (CY) %</td>
		<td class='HeaderTD' style='background-color: #ececec;' nowrap><%=currentMonth%> Sales Budget (CY)</td>
		<td class='HeaderTD' style='background-color: #ececec;' nowrap>Target Ach MTD (CY) %</td>
		<td class='HeaderTD' style='background-color: #f7caac;' nowrap>MTD Sale (PY)</td>
		<td class='HeaderTD' style='background-color: #f7caac;' nowrap>MTD Margin (PY)</td>
		<td class='HeaderTD' style='background-color: #f7caac;' nowrap>MTD Margin (PY) %</td>
		<td class='HeaderTD' style='background-color: #ececec;' nowrap>MTD Growth %</td>
		<td class='HeaderTD' style='background-color: #fef2cb;' nowrap>Uninvoiced</td>
		<td class='HeaderTD' style='background-color: #fef2cb;' nowrap>MTD Sale (CY) + Uninv</td>
		<td class='HeaderTD' style='background-color: #d9e2f3;' nowrap>YTD Sale (CY)</td>
		<td class='HeaderTD' style='background-color: #d9e2f3;' nowrap>YTD Margin (CY)</td>
		<td class='HeaderTD' style='background-color: #d9e2f3;' nowrap>YTD Margin (CY) %</td>
		<td class='HeaderTD' style='background-color: #ececec;' nowrap>YTD Sales Budget (CY)</td>
		<td class='HeaderTD' style='background-color: #ececec;' nowrap>Target Ach YTD (CY) %</td>
		<td class='HeaderTD' style='background-color: #f7caac;' nowrap>YTD Sale (PY)</td>
		<td class='HeaderTD' style='background-color: #f7caac;' nowrap>YTD Margin (PY)</td>
		<td class='HeaderTD' style='background-color: #f7caac;' nowrap>YTD Margin (PY) %</td>
		<td class='HeaderTD' style='background-color: #f7caac;' nowrap>YTD Growth (PY) %</td>
	</tr>
	<% IF ""& Request("Type") = "" Then %>
		<% For Each SalesPerson In SalesPersons.rows
			TodaySale = IIF(""& Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate = '"& eDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate = '"& eDate &"'"))
			MTDSaleCY = IIF(""& Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& cmSdate &"' AND TransactionDate <= '"& eDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& cmSdate &"' AND TransactionDate <= '"& eDate &"'"))
			MTDMarginCY = IIF(""& Sales.Compute("SUM(Margin)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& cmSdate &"' AND TransactionDate <= '"& eDate &"'") = "", "0", Sales.Compute("SUM(Margin)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& cmSdate &"' AND TransactionDate <= '"& eDate &"'"))
			MTDSalePY = IIF(""& Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& pymSdate &"' AND TransactionDate <= '"& pyeDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& pymSdate &"' AND TransactionDate <= '"& pyeDate &"'"))
			MTDMarginPY = IIF(""& Sales.Compute("SUM(Margin)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& pymSdate &"' AND TransactionDate <= '"& pyeDate &"'") = "", "0", Sales.Compute("SUM(Margin)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& pymSdate &"' AND TransactionDate <= '"& pyeDate &"'"))
			YTDSaleCY = IIF(""& Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& cySdate &"' AND TransactionDate <= '"& eDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& cySdate &"' AND TransactionDate <= '"& eDate &"'"))
			YTDMarginCY = IIF(""& Sales.Compute("SUM(Margin)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& cySdate &"' AND TransactionDate <= '"& eDate &"'") = "", "0", Sales.Compute("SUM(Margin)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& cySdate &"' AND TransactionDate <= '"& eDate &"'"))
			YTDSalePY = IIF(""& Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& pySdate &"' AND TransactionDate <= '"& pyeDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& pySdate &"' AND TransactionDate <= '"& pyeDate &"'"))
			YTDMarginPY = IIF(""& Sales.Compute("SUM(Margin)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& pySdate &"' AND TransactionDate <= '"& pyeDate &"'") = "", "0", Sales.Compute("SUM(Margin)", "SalesRepID = '"& SalesPerson("ID") &"' AND TransactionDate >= '"& pySdate &"' AND TransactionDate <= '"& pyeDate &"'"))
			UnInvoicedSales = Sql.ReadField("SELECT ISNULL(SUM(Extended),0) Extended FROM R_UninvoicedSales WHERE SalesRepID = "& SalesPerson("ID"))
			CurrentMonthBudget = Sql.ReadField("SELECT cast(ISNULL(Month"& currentMonthID &"Amt,0) as decimal(18,2)) FROM CustomBudgetingSalesRep WHERE SalesRepID = "& SalesPerson("ID") &" AND BudgetYear = '"& currentMonthYearStartDate &"'")
			CurrentYearBudget = Sql.ReadField("SELECT ISNULL(cast(ISNULL(Month1Amt,0) as decimal(18,2)) + cast(ISNULL(Month2Amt,0) as decimal(18,2)) + cast(ISNULL(Month3Amt,0) as decimal(18,2)) + cast(ISNULL(Month4Amt,0) as decimal(18,2)) + cast(ISNULL(Month5Amt,0) as decimal(18,2)) + cast(ISNULL(Month6Amt,0) as decimal(18,2)) + cast(ISNULL(Month7Amt,0) as decimal(18,2)) + cast(ISNULL(Month8Amt,0) as decimal(18,2)) + cast(ISNULL(Month9Amt,0) as decimal(18,2)) + cast(ISNULL(Month10Amt,0) as decimal(18,2)) + cast(ISNULL(Month11Amt,0) as decimal(18,2)) + cast(ISNULL(Month12Amt,0) as decimal(18,2)),0) AS CurrentYearBudget FROM CustomBudgetingSalesRep WHERE SalesRepID = "& SalesPerson("ID") &" AND BudgetYear = '"& currentMonthYearStartDate &"'")
			IF SalesPerson("InActive") = "" OR CInt(TodaySale + MTDSaleCY + MTDSalePY + YTDSaleCY + YTDSalePY + UnInvoicedSales) <> 0 Then
				DM.Write("<tr class='TR "& IIF(""& Request("filter") <> "", IIF(""& Request("filter") = "hide", IIF(SalesPersonIDs.Contains(SalesPerson("ID")), "hide", ""), IIF(SalesPersonIDs.Contains(SalesPerson("ID")), "", "hide")), "") &"' "& IIF(""& SalesPerson("InActive") = "", "", "style='background-color: #fde5e5'") &">")
					DM.Write("<td class='DetailTD' style='text-align: left; padding-right: 7px;' nowrap>"& SalesPerson("SalesPerson") &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(TodaySale) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(MTDSaleCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(MTDMarginCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((MTDMarginCY/MTDSaleCY)*100 > 0 OR (MTDMarginCY/MTDSaleCY)*100 < 0) And (MTDSaleCY > 0 OR MTDSaleCY < 0), (MTDMarginCY/MTDSaleCY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(CurrentMonthBudget) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((MTDSaleCY/CurrentMonthBudget)*100 > 0 OR (MTDSaleCY/CurrentMonthBudget)*100 < 0) And (CurrentMonthBudget > 0 OR CurrentMonthBudget < 0), (MTDSaleCY/CurrentMonthBudget)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(MTDSalePY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(MTDMarginPY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((MTDMarginPY/MTDSalePY)*100 > 0 OR (MTDMarginPY/MTDSalePY)*100 < 0) AND (MTDSalePY > 0 OR MTDSalePY < 0), (MTDMarginPY/MTDSalePY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((MTDSaleCY/MTDSalePY)*100 > 0 OR (MTDSaleCY/MTDSalePY)*100 < 0) AND (MTDSalePY > 0 OR MTDSalePY < 0), (MTDSaleCY/MTDSalePY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(UnInvoicedSales) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(UnInvoicedSales+MTDSaleCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(YTDSaleCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(YTDMarginCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((YTDMarginCY/YTDSaleCY)*100 > 0 OR (YTDMarginCY/YTDSaleCY)*100 < 0) And (YTDSaleCY > 0 OR YTDSaleCY < 0), (YTDMarginCY/YTDSaleCY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(CurrentYearBudget) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((YTDSaleCY/CurrentYearBudget)*100 > 0 OR (YTDSaleCY/CurrentYearBudget)*100 < 0) And (CurrentYearBudget > 0 OR CurrentYearBudget < 0), (YTDSaleCY/CurrentYearBudget)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(YTDSalePY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(YTDMarginPY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((YTDMarginPY/YTDSalePY)*100 > 0 OR (YTDMarginPY/YTDSalePY)*100 < 0) AND (YTDSalePY > 0 OR YTDSalePY < 0), (YTDMarginPY/YTDSalePY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((YTDSaleCY/YTDSalePY)*100 > 0 OR (YTDSaleCY/YTDSalePY)*100 < 0) AND (YTDSalePY > 0 OR YTDSalePY < 0), (YTDSaleCY/YTDSalePY)*100, 0)) &"%</td>")
				DM.Write("</tr>")
			End IF
			TodaySaleTotal = TodaySaleTotal + TodaySale
			MTDSaleCYTotal = MTDSaleCYTotal + MTDSaleCY
			MTDMarginCYTotal = MTDMarginCYTotal + MTDMarginCY
			MTDSalePYTotal = MTDSalePYTotal + MTDSalePY
			MTDMarginPYTotal = MTDMarginPYTotal + MTDMarginPY
			YTDSaleCYTotal = YTDSaleCYTotal + YTDSaleCY
			YTDMarginCYTotal = YTDMarginCYTotal + YTDMarginCY
			YTDSalePYTotal = YTDSalePYTotal + YTDSalePY
			YTDMarginPYTotal = YTDMarginPYTotal + YTDMarginPY
			UnInvoicedSalesTotal = UnInvoicedSalesTotal + UnInvoicedSales
			CurrentMonthBudgetTotal = CurrentMonthBudgetTotal + CurrentMonthBudget
			CurrentYearBudgetTotal = CurrentYearBudgetTotal + CurrentYearBudget
			IF (("4233,4661,4662").Contains(SalesPerson("ID"))) Then
				ISTodaySale = ISTodaySale + TodaySale
				ISMTDSaleCY = ISMTDSaleCY + MTDSaleCY
				ISMTDMarginCY = ISMTDMarginCY + MTDMarginCY
				ISMTDSalePY = ISMTDSalePY + MTDSalePY
				ISMTDMarginPY = ISMTDMarginPY + MTDMarginPY
				ISYTDSaleCY = ISYTDSaleCY + YTDSaleCY
				ISYTDMarginCY = ISYTDMarginCY + YTDMarginCY
				ISYTDSalePY = ISYTDSalePY + YTDSalePY
				ISYTDMarginPY = ISYTDMarginPY + YTDMarginPY
				ISUnInvoicedSales = ISUnInvoicedSales + UnInvoicedSales
				ISCurrentMonthBudget = ISCurrentMonthBudget + CurrentMonthBudget
				ISCurrentYearBudget = ISCurrentYearBudget + CurrentYearBudget
			End IF
		Next %>
	<% Else %>
		<% For Each Category In Categories.rows
			TodaySale = IIF(""& Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate = '"& eDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate = '"& eDate &"'"))
			MTDSaleCY = IIF(""& Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& cmSdate &"' AND TransactionDate <= '"& eDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& cmSdate &"' AND TransactionDate <= '"& eDate &"'"))
			MTDMarginCY = IIF(""& Sales.Compute("SUM(Margin)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& cmSdate &"' AND TransactionDate <= '"& eDate &"'") = "", "0", Sales.Compute("SUM(Margin)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& cmSdate &"' AND TransactionDate <= '"& eDate &"'"))
			MTDSalePY = IIF(""& Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& pymSdate &"' AND TransactionDate <= '"& pyeDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& pymSdate &"' AND TransactionDate <= '"& pyeDate &"'"))
			MTDMarginPY = IIF(""& Sales.Compute("SUM(Margin)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& pymSdate &"' AND TransactionDate <= '"& pyeDate &"'") = "", "0", Sales.Compute("SUM(Margin)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& pymSdate &"' AND TransactionDate <= '"& pyeDate &"'"))
			YTDSaleCY = IIF(""& Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& cySdate &"' AND TransactionDate <= '"& eDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& cySdate &"' AND TransactionDate <= '"& eDate &"'"))
			YTDMarginCY = IIF(""& Sales.Compute("SUM(Margin)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& cySdate &"' AND TransactionDate <= '"& eDate &"'") = "", "0", Sales.Compute("SUM(Margin)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& cySdate &"' AND TransactionDate <= '"& eDate &"'"))
			YTDSalePY = IIF(""& Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& pySdate &"' AND TransactionDate <= '"& pyeDate &"'") = "", "0", Sales.Compute("SUM(Extended)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& pySdate &"' AND TransactionDate <= '"& pyeDate &"'"))
			YTDMarginPY = IIF(""& Sales.Compute("SUM(Margin)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& pySdate &"' AND TransactionDate <= '"& pyeDate &"'") = "", "0", Sales.Compute("SUM(Margin)", "ItemCategoryID = '"& Category("ID") &"' AND TransactionDate >= '"& pySdate &"' AND TransactionDate <= '"& pyeDate &"'"))
			UnInvoicedSales = Sql.ReadField("SELECT ISNULL(SUM(Extended),0) Extended FROM R_UninvoicedSales WHERE ItemCategoryID = "& Category("ID"))
			CurrentMonthBudget = Sql.ReadField("SELECT cast(ISNULL(Month"& currentMonthID &"Amt,0) as decimal(18,2)) FROM CustomBudgetingProductCategory WHERE CategoryID = "& Category("ID") &" AND BudgetYear = '"& currentMonthYearStartDate &"'")
			CurrentYearBudget = Sql.ReadField("SELECT ISNULL(cast(ISNULL(Month1Amt,0) as decimal(18,2)) + cast(ISNULL(Month2Amt,0) as decimal(18,2)) + cast(ISNULL(Month3Amt,0) as decimal(18,2)) + cast(ISNULL(Month4Amt,0) as decimal(18,2)) + cast(ISNULL(Month5Amt,0) as decimal(18,2)) + cast(ISNULL(Month6Amt,0) as decimal(18,2)) + cast(ISNULL(Month7Amt,0) as decimal(18,2)) + cast(ISNULL(Month8Amt,0) as decimal(18,2)) + cast(ISNULL(Month9Amt,0) as decimal(18,2)) + cast(ISNULL(Month10Amt,0) as decimal(18,2)) + cast(ISNULL(Month11Amt,0) as decimal(18,2)) + cast(ISNULL(Month12Amt,0) as decimal(18,2)),0) AS CurrentYearBudget FROM CustomBudgetingProductCategory WHERE CategoryID = "& Category("ID") &" AND BudgetYear = '"& currentMonthYearStartDate &"'")
			IF CInt(TodaySale + MTDSaleCY + MTDSalePY + YTDSaleCY + YTDSalePY + UnInvoicedSales) <> 0 Then
				DM.Write("<tr class='TR "& IIF(""& Request("filter") <> "", IIF(""& Request("filter") = "hide", IIF(CategoryIDs.Contains(Category("ID")), "hide", ""), IIF(""& CategoryIDs <> "" And CategoryIDs.Contains(Category("ID")), "", "hide")), "") &"'>")
					DM.Write("<td class='DetailTD' style='text-align: left; padding-right: 7px;' nowrap>"& Category("Category") &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(TodaySale) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(MTDSaleCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(MTDMarginCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((MTDMarginCY/MTDSaleCY)*100 > 0 OR (MTDMarginCY/MTDSaleCY)*100 < 0) And (MTDSaleCY > 0 OR MTDSaleCY < 0), (MTDMarginCY/MTDSaleCY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(CurrentMonthBudget) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((MTDSaleCY/CurrentMonthBudget)*100 > 0 OR (MTDSaleCY/CurrentMonthBudget)*100 < 0) And (CurrentMonthBudget > 0 OR CurrentMonthBudget < 0), (MTDSaleCY/CurrentMonthBudget)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(MTDSalePY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(MTDMarginPY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((MTDMarginPY/MTDSalePY)*100 > 0 OR (MTDMarginPY/MTDSalePY)*100 < 0) AND (MTDSalePY > 0 OR MTDSalePY < 0), (MTDMarginPY/MTDSalePY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((MTDSaleCY/MTDSalePY)*100 > 0 OR (MTDSaleCY/MTDSalePY)*100 < 0) AND (MTDSalePY > 0 OR MTDSalePY < 0), (MTDSaleCY/MTDSalePY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(UnInvoicedSales) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(UnInvoicedSales+MTDSaleCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(YTDSaleCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(YTDMarginCY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((YTDMarginCY/YTDSaleCY)*100 > 0 OR (YTDMarginCY/YTDSaleCY)*100 < 0) And (YTDSaleCY > 0 OR YTDSaleCY < 0), (YTDMarginCY/YTDSaleCY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(CurrentYearBudget) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((YTDSaleCY/CurrentYearBudget)*100 > 0 OR (YTDSaleCY/CurrentYearBudget)*100 < 0) And (CurrentYearBudget > 0 OR CurrentYearBudget < 0), (YTDSaleCY/CurrentYearBudget)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(YTDSalePY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.FormatCurrency(YTDMarginPY) &"</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((YTDMarginPY/YTDSalePY)*100 > 0 OR (YTDMarginPY/YTDSalePY)*100 < 0) AND (YTDSalePY > 0 OR YTDSalePY < 0), (YTDMarginPY/YTDSalePY)*100, 0)) &"%</td>")
					DM.Write("<td class='DetailTD'>"& DM.SPSFormatNumber(IIF(((YTDSaleCY/YTDSalePY)*100 > 0 OR (YTDSaleCY/YTDSalePY)*100 < 0) AND (YTDSalePY > 0 OR YTDSalePY < 0), (YTDSaleCY/YTDSalePY)*100, 0)) &"%</td>")
				DM.Write("</tr>")
			End IF
			TodaySaleTotal = TodaySaleTotal + TodaySale
			MTDSaleCYTotal = MTDSaleCYTotal + MTDSaleCY
			MTDMarginCYTotal = MTDMarginCYTotal + MTDMarginCY
			MTDSalePYTotal = MTDSalePYTotal + MTDSalePY
			MTDMarginPYTotal = MTDMarginPYTotal + MTDMarginPY
			YTDSaleCYTotal = YTDSaleCYTotal + YTDSaleCY
			YTDMarginCYTotal = YTDMarginCYTotal + YTDMarginCY
			YTDSalePYTotal = YTDSalePYTotal + YTDSalePY
			YTDMarginPYTotal = YTDMarginPYTotal + YTDMarginPY
			UnInvoicedSalesTotal = UnInvoicedSalesTotal + UnInvoicedSales
			CurrentMonthBudgetTotal = CurrentMonthBudgetTotal + CurrentMonthBudget
			CurrentYearBudgetTotal = CurrentYearBudgetTotal + CurrentYearBudget
			IF (("4233,4661,4662").Contains("0")) Then
				ISTodaySale = ISTodaySale + TodaySale
				ISMTDSaleCY = ISMTDSaleCY + MTDSaleCY
				ISMTDMarginCY = ISMTDMarginCY + MTDMarginCY
				ISMTDSalePY = ISMTDSalePY + MTDSalePY
				ISMTDMarginPY = ISMTDMarginPY + MTDMarginPY
				ISYTDSaleCY = ISYTDSaleCY + YTDSaleCY
				ISYTDMarginCY = ISYTDMarginCY + YTDMarginCY
				ISYTDSalePY = ISYTDSalePY + YTDSalePY
				ISYTDMarginPY = ISYTDMarginPY + YTDMarginPY
				ISUnInvoicedSales = ISUnInvoicedSales + UnInvoicedSales
				ISCurrentMonthBudget = ISCurrentMonthBudget + CurrentMonthBudget
				ISCurrentYearBudget = ISCurrentYearBudget + CurrentYearBudget
			End IF
		Next %>
	<% End IF %>
	<tr>
		<td class='HeaderTD' style='text-align: left; background-color: #fbe4d5; padding-left: 7px;' nowrap>Gross Sales</td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(TodaySaleTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(MTDSaleCYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(MTDMarginCYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.SPSFormatNumber(IIF(((MTDMarginCYTotal/MTDSaleCYTotal)*100 > 0 OR (MTDMarginCYTotal/MTDSaleCYTotal)*100 < 0) And (MTDSaleCYTotal > 0 OR MTDSaleCYTotal < 0), (MTDMarginCYTotal/MTDSaleCYTotal)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(CurrentMonthBudgetTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.SPSFormatNumber(IIF(((MTDSaleCYTotal/CurrentMonthBudgetTotal)*100 > 0 OR (MTDSaleCYTotal/CurrentMonthBudgetTotal)*100 < 0) And (CurrentMonthBudgetTotal > 0 OR CurrentMonthBudgetTotal < 0), (MTDSaleCYTotal/CurrentMonthBudgetTotal)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(MTDSalePYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(MTDMarginPYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.SPSFormatNumber(IIF(((MTDMarginPYTotal/MTDSalePYTotal)*100 > 0 OR (MTDMarginPYTotal/MTDSalePYTotal)*100 < 0) AND (MTDSalePYTotal > 0 OR MTDSalePYTotal < 0), (MTDMarginPYTotal/MTDSalePYTotal)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.SPSFormatNumber(IIF(((MTDSaleCYTotal/MTDSalePYTotal)*100 > 0 OR (MTDSaleCYTotal/MTDSalePYTotal)*100 < 0) AND (MTDSalePYTotal > 0 OR MTDSalePYTotal < 0), (MTDSaleCYTotal/MTDSalePYTotal)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(UnInvoicedSalesTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(UnInvoicedSalesTotal+MTDSaleCYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(YTDSaleCYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(YTDMarginCYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.SPSFormatNumber(IIF(((YTDMarginCYTotal/YTDSaleCYTotal)*100 > 0 OR (YTDMarginCYTotal/YTDSaleCYTotal)*100 < 0) And (YTDSaleCYTotal > 0 OR YTDSaleCYTotal < 0), (YTDMarginCYTotal/YTDSaleCYTotal)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(CurrentYearBudgetTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.SPSFormatNumber(IIF(((YTDSaleCYTotal/CurrentYearBudgetTotal)*100 > 0 OR (YTDSaleCYTotal/CurrentYearBudgetTotal)*100 < 0) And (CurrentYearBudgetTotal > 0 OR CurrentYearBudgetTotal < 0), (YTDSaleCYTotal/CurrentYearBudgetTotal)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(YTDSalePYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.FormatCurrency(YTDMarginPYTotal)%></td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.SPSFormatNumber(IIF(((YTDMarginPYTotal/YTDSalePYTotal)*100 > 0 OR (YTDMarginPYTotal/YTDSalePYTotal)*100 < 0) AND (YTDSalePYTotal > 0 OR YTDSalePYTotal < 0), (YTDMarginPYTotal/YTDSalePYTotal)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #fbe4d5;' nowrap><%=DM.SPSFormatNumber(IIF(((YTDSaleCYTotal/YTDSalePYTotal)*100 > 0 OR (YTDSaleCYTotal/YTDSalePYTotal)*100 < 0) AND (YTDSalePYTotal > 0 OR YTDSalePYTotal < 0), (YTDSaleCYTotal/YTDSalePYTotal)*100, 0))%>%</td>
	</tr>
	<tr>
		<td class='HeaderTD' style='text-align: left; background-color: #d6eaf1; padding-left: 7px;' nowrap>Sales To Internal</td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISTodaySale)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISMTDSaleCY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISMTDMarginCY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.SPSFormatNumber(IIF(((ISMTDMarginCY/ISMTDSaleCY)*100 > 0 OR (ISMTDMarginCY/ISMTDSaleCY)*100 < 0) And (ISMTDSaleCY > 0 OR ISMTDSaleCY < 0), (ISMTDMarginCY/ISMTDSaleCY)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISCurrentMonthBudget)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.SPSFormatNumber(IIF(((ISMTDSaleCY/ISCurrentMonthBudget)*100 > 0 OR (ISMTDSaleCY/ISCurrentMonthBudget)*100 < 0) And (ISCurrentMonthBudget > 0 OR ISCurrentMonthBudget < 0), (ISMTDSaleCY/ISCurrentMonthBudget)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISMTDSalePY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISMTDMarginPY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.SPSFormatNumber(IIF(((ISMTDMarginPY/ISMTDSalePY)*100 > 0 OR (ISMTDMarginPY/ISMTDSalePY)*100 < 0) AND (ISMTDSalePY > 0 OR ISMTDSalePY < 0), (ISMTDMarginPY/ISMTDSalePY)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.SPSFormatNumber(IIF(((ISMTDSaleCY/ISMTDSalePY)*100 > 0 OR (ISMTDSaleCY/ISMTDSalePY)*100 < 0) AND (ISMTDSalePY > 0 OR ISMTDSalePY < 0), (ISMTDSaleCY/ISMTDSalePY)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISUnInvoicedSales)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISUnInvoicedSales+ISMTDSaleCY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISYTDSaleCY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISYTDMarginCY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.SPSFormatNumber(IIF(((ISYTDMarginCY/ISYTDSaleCY)*100 > 0 OR (ISYTDMarginCY/ISYTDSaleCY)*100 < 0) And (ISYTDSaleCY > 0 OR ISYTDSaleCY < 0), (ISYTDMarginCY/ISYTDSaleCY)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISCurrentYearBudget)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.SPSFormatNumber(IIF(((ISYTDSaleCY/ISCurrentYearBudget)*100 > 0 OR (ISYTDSaleCY/ISCurrentYearBudget)*100 < 0) And (ISCurrentYearBudget > 0 OR ISCurrentYearBudget < 0), (ISYTDSaleCY/ISCurrentYearBudget)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISYTDSalePY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.FormatCurrency(ISYTDMarginPY)%></td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.SPSFormatNumber(IIF(((ISYTDMarginPY/ISYTDSalePY)*100 > 0 OR (ISYTDMarginPY/ISYTDSalePY)*100 < 0) AND (ISYTDSalePY > 0 OR ISYTDSalePY < 0), (ISYTDMarginPY/ISYTDSalePY)*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: #d6eaf1;' nowrap><%=DM.SPSFormatNumber(IIF(((ISYTDSaleCY/ISYTDSalePY)*100 > 0 OR (ISYTDSaleCY/ISYTDSalePY)*100 < 0) AND (ISYTDSalePY > 0 OR ISYTDSalePY < 0), (ISYTDSaleCY/ISYTDSalePY)*100, 0))%>%</td>
	</tr>
	<tr>
		<td class='HeaderTD' style='text-align: left; background-color: aliceblue; padding-left: 7px; padding-right: 7px;' nowrap>Net Sales (Less Internal Sales)</td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(TodaySaleTotal-ISTodaySale)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(MTDSaleCYTotal-ISMTDSaleCY)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(MTDMarginCYTotal-ISMTDMarginCY)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.SPSFormatNumber(IIF((((MTDMarginCYTotal-ISMTDMarginCY)/(MTDSaleCYTotal-ISMTDSaleCY))*100 > 0 OR ((MTDMarginCYTotal-ISMTDMarginCY)/(MTDSaleCYTotal-ISMTDSaleCY))*100 < 0) And ((MTDSaleCYTotal-ISMTDSaleCY) > 0 OR (MTDSaleCYTotal-ISMTDSaleCY) < 0), ((MTDMarginCYTotal-ISMTDMarginCY)/(MTDSaleCYTotal-ISMTDSaleCY))*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(CurrentMonthBudgetTotal-ISCurrentMonthBudget)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.SPSFormatNumber(IIF((((MTDSaleCYTotal-ISMTDSaleCY)/(CurrentMonthBudgetTotal-ISCurrentMonthBudget))*100 > 0 OR ((MTDSaleCYTotal-ISMTDSaleCY)/(CurrentMonthBudgetTotal-ISCurrentMonthBudget))*100 < 0) And ((CurrentMonthBudgetTotal-ISCurrentMonthBudget) > 0 OR (CurrentMonthBudgetTotal-ISCurrentMonthBudget) < 0), ((MTDSaleCYTotal-ISMTDSaleCY)/(CurrentMonthBudgetTotal-ISCurrentMonthBudget))*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency((MTDSalePYTotal-ISMTDSalePY))%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency((MTDMarginPYTotal-ISMTDMarginPY))%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.SPSFormatNumber(IIF((((MTDMarginPYTotal-ISMTDMarginPY)/(MTDSalePYTotal-ISMTDSalePY))*100 > 0 OR ((MTDMarginPYTotal-ISMTDMarginPY)/(MTDSalePYTotal-ISMTDSalePY))*100 < 0) AND ((MTDSalePYTotal-ISMTDSalePY) > 0 OR (MTDSalePYTotal-ISMTDSalePY) < 0), ((MTDMarginPYTotal-ISMTDMarginPY)/(MTDSalePYTotal-ISMTDSalePY))*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.SPSFormatNumber(IIF((((MTDSaleCYTotal-ISMTDSaleCY)/(MTDSalePYTotal-ISMTDSalePY))*100 > 0 OR ((MTDSaleCYTotal-ISMTDSaleCY)/(MTDSalePYTotal-ISMTDSalePY))*100 < 0) AND ((MTDSalePYTotal-ISMTDSalePY) > 0 OR (MTDSalePYTotal-ISMTDSalePY) < 0), ((MTDSaleCYTotal-ISMTDSaleCY)/(MTDSalePYTotal-ISMTDSalePY))*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(UnInvoicedSalesTotal-ISUnInvoicedSales)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency((UnInvoicedSalesTotal-ISUnInvoicedSales)+(MTDSaleCYTotal-ISMTDSaleCY))%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(YTDSaleCYTotal-ISYTDSaleCY)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(YTDMarginCYTotal-ISYTDMarginCY)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.SPSFormatNumber(IIF((((YTDMarginCYTotal-ISYTDMarginCY)/(YTDSaleCYTotal-ISYTDSaleCY))*100 > 0 OR ((YTDMarginCYTotal-ISYTDMarginCY)/(YTDSaleCYTotal-ISYTDSaleCY))*100 < 0) And ((YTDSaleCYTotal-ISYTDSaleCY) > 0 OR (YTDSaleCYTotal-ISYTDSaleCY) < 0), ((YTDMarginCYTotal-ISYTDMarginCY)/(YTDSaleCYTotal-ISYTDSaleCY))*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(CurrentYearBudgetTotal-ISCurrentYearBudget)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.SPSFormatNumber(IIF((((YTDSaleCYTotal-ISYTDSaleCY)/(CurrentYearBudgetTotal-ISCurrentYearBudget))*100 > 0 OR ((YTDSaleCYTotal-ISYTDSaleCY)/(CurrentYearBudgetTotal-ISCurrentYearBudget))*100 < 0) And ((CurrentYearBudgetTotal-ISCurrentYearBudget) > 0 OR (CurrentYearBudgetTotal-ISCurrentYearBudget) < 0), ((YTDSaleCYTotal-ISYTDSaleCY)/(CurrentYearBudgetTotal-ISCurrentYearBudget))*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(YTDSalePYTotal-ISYTDSalePY)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.FormatCurrency(YTDMarginPYTotal-ISYTDMarginPY)%></td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.SPSFormatNumber(IIF((((YTDMarginPYTotal-ISYTDMarginPY)/(YTDSalePYTotal-ISYTDSalePY))*100 > 0 OR ((YTDMarginPYTotal-ISYTDMarginPY)/(YTDSalePYTotal-ISYTDSalePY))*100 < 0) AND ((YTDSalePYTotal-ISYTDSalePY) > 0 OR (YTDSalePYTotal-ISYTDSalePY) < 0), ((YTDMarginPYTotal-ISYTDMarginPY)/(YTDSalePYTotal-ISYTDSalePY))*100, 0))%>%</td>
		<td class='HeaderTD' style='background-color: aliceblue;' nowrap><%=DM.SPSFormatNumber(IIF((((YTDSaleCYTotal-ISYTDSaleCY)/(YTDSalePYTotal-ISYTDSalePY))*100 > 0 OR ((YTDSaleCYTotal-ISYTDSaleCY)/(YTDSalePYTotal-ISYTDSalePY))*100 < 0) AND ((YTDSalePYTotal-ISYTDSalePY) > 0 OR (YTDSalePYTotal-ISYTDSalePY) < 0), ((YTDSaleCYTotal-ISYTDSaleCY)/(YTDSalePYTotal-ISYTDSalePY))*100, 0))%>%</td>
	</tr>
</table>
<style>
	.TitleTD {
		border: 1px solid #ddd;
		padding: 4px;
		font-family: roboto;
		background-color: rgba(0,0,0,.03);
		color: #2171ae;
		padding: 12px 7px;
		font-size: 20px;
		border-bottom: none;
	}
	.HeaderTD {
		border: 1px solid #aaa;
		padding: 7px 4px 7px 10px;
		font-family: roboto;
		text-align: right;
		font-size: 12px;
		font-weight: bold;

	}
	.DetailTD {
		border: 1px solid #ddd;
		padding: 6px 4px 6px 7px;
		font-family: roboto;
		text-align: right;
		font-size: 12px;
		border-bottom: none;
	}
	.TR:hover {
		background-color: rgba(0,0,0,.05);
	}
	.TR {
		background-color: rgba(0,0,0,.02);
	}
	.DetailTD:hover {
		background-color: rgba(0,0,0,.05);
		cursor: pointer;
	}
	.ui-datepicker-trigger {
		margin-bottom: 0px !important;
		cursor: pointer !important;
	}
	.select2-selection__clear {
		border: 1px solid #828282;
		background-color: #e4e4e4;
		border-radius: 3px;
		padding-left: 5px !important;
		padding-right: 5px !important;
		margin-right: 0px !important;
		font-size: 12px;
		font-family: 'Roboto';
	}
	.select2-selection__choice {
		padding-top: 2px !important;
		padding-bottom: 2px !important;
		font-family: 'Roboto';
		border-radius: 3px;
	}
	.select2-container--default .select2-selection--multiple {
		border-radius: 3px !important;
	}
	.select2-container {
		font-size: 12px;
		font-family: 'Roboto';
	}
</style>
<link href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" rel="stylesheet" />
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js"></script>
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
<script>
	var type = '<%=""& Request("Type")%>';
	$(document).ready(function() {
		if (type == '')
		{
			$('#SalesPerson').select2({
				placeholder: 'Select Sales People...',
				minimumResultsForSearch: Infinity,
				allowClear: true
			});
		}
		else
		{
			$('#Category').select2({
				placeholder: 'Select Categories...',
				minimumResultsForSearch: Infinity,
				allowClear: true
			});
		}
    });
	
    function search(filter) {
		var selectedValues
		if (type == '')
		{
			selectedValues = $('#SalesPerson').val();
		}
		else
		{
			selectedValues = $('#Category').val();
		}
		let url = new URL(location.href);
		let params = new URLSearchParams(url.search);
		params.delete('q5');
		params.delete('filter');
		if (filter == 1)
		{
			location.href = 'R_SalesComparison.aspx?' + params
		}
		else if (filter == 2)
		{
			location.href = 'R_SalesComparison.aspx?' + params + '&filter=show&q5='+ selectedValues
		}
		else
		{
			location.href = 'R_SalesComparison.aspx?' + params + '&filter=hide&q5='+ selectedValues
		}
    }
	function changeDate(value) {
		let url = new URL(location.href);
		let params = new URLSearchParams(url.search);
		params.delete('q2');
		location.href = "R_SalesComparison.aspx?q2="+ value + '&' + params
	}
	function exportTableToExcel() {
		var table = document.getElementById("Sales");
		var workbook = XLSX.utils.book_new();
		var worksheet = XLSX.utils.table_to_sheet(table, { range: 1 });
		XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
		XLSX.writeFile(workbook, "SalesSummarybySalesRep.xlsx");
	}
</script>
