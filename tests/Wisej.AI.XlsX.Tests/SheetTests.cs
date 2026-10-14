///////////////////////////////////////////////////////////////////////////////
//
// (C) 2026 ICE TEA GROUP LLC - ALL RIGHTS RESERVED
//
//
//
// ALL INFORMATION CONTAINED HEREIN IS, AND REMAINS
// THE PROPERTY OF ICE TEA GROUP LLC AND ITS SUPPLIERS, IF ANY.
// THE INTELLECTUAL PROPERTY AND TECHNICAL CONCEPTS CONTAINED
// HEREIN ARE PROPRIETARY TO ICE TEA GROUP LLC AND ITS SUPPLIERS
// AND MAY BE COVERED BY U.S. AND FOREIGN PATENTS, PATENT IN PROCESS, AND
// ARE PROTECTED BY TRADE SECRET OR COPYRIGHT LAW.
//
// DISSEMINATION OF THIS INFORMATION OR REPRODUCTION OF THIS MATERIAL
// IS STRICTLY FORBIDDEN UNLESS PRIOR WRITTEN PERMISSION IS OBTAINED
// FROM ICE TEA GROUP LLC.
//
///////////////////////////////////////////////////////////////////////////////

using System.Linq;
using Xunit;

namespace Wisej.AI.XlsX.Tests
{
	/// <summary>
	/// Tests the conversion of cells, ranges and tables.
	/// </summary>
	public class SheetTests : TestBase
	{
		/// <summary>
		/// Returns a workbook with a header row, shared strings and number formats.
		/// </summary>
		private static XlsxBuilder SalesWorkbook()
		{
			return new XlsxBuilder()
				.SharedStrings("Name", "Price", "Date", "Apple", "Pear | Bartlett")
				.Styles(new[] { 0, 4, 14 })
				.Sheet("Sales", Data(
					Row(1, S("A1", 0), S("B1", 1), S("C1", 2)),
					Row(2, S("A2", 3), N("B2", 1234.5, 1), N("C2", 44927, 2)),
					Row(3, S("A3", 4), N("B3", 0.9, 1), N("C3", 44928, 2))));
		}

		[Fact]
		public void Range_BecomesAPipeTableWithFormattedValues()
		{
			var text = Markdown(SalesWorkbook());

			Assert.Equal(
				"## Sales\n\n" +
				"**Range A1:C3**\n\n" +
				"| Name | Price | Date |\n" +
				"| --- | --- | --- |\n" +
				"| Apple | 1,234.50 | 2023-01-01 |\n" +
				"| Pear \\| Bartlett | 0.90 | 2023-01-02 |",
				text);
		}

		[Fact]
		public void PlainText_WritesPipeSeparatedRows()
		{
			var text = PlainText(SalesWorkbook());

			Assert.Equal(
				"Sheet: Sales\n\n" +
				"Range A1:C3\n\n" +
				"Name | Price | Date\n" +
				"Apple | 1,234.50 | 2023-01-01\n" +
				"Pear | Bartlett | 0.90 | 2023-01-02",
				text);
		}

		[Fact]
		public void PageHeaders_IntroduceEverySheet()
		{
			var builder = new XlsxBuilder()
				.Sheet("First", Data(Row(1, T("A1", "One"))))
				.Sheet("Second", Data(Row(1, T("A1", "Two"))));

			var result = builder.Extract(new XlsxExtractorOptions());

			Assert.Equal(2, result.PageCount);
			Assert.Equal("Second", result.Pages[1].SheetName);
			Assert.Equal("## Second\n\nTwo", result.Pages[1].Text);
			Assert.StartsWith("<!-- Page 1 -->\n\n## First", result.Text);
			Assert.Contains("<!-- Page 2 -->\n\n## Second", result.Text);
		}

		[Fact]
		public void Titles_BecomeParagraphsAboveTheTable()
		{
			var builder = new XlsxBuilder().Sheet("Report",
				Data(
					Row(1, T("A1", "Quarterly report")),
					Row(3, T("A3", "Region"), T("B3", "Q1")),
					Row(4, T("A4", "North"), N("B4", 10)),
					Row(5, T("A5", "South"), N("B5", 20))) +
				"<mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>");

			var text = Markdown(builder, o => o.IncludeCellRanges = false);

			Assert.Equal(
				"## Report\n\n" +
				"Quarterly report\n\n" +
				"| Region | Q1 |\n" +
				"| --- | --- |\n" +
				"| North | 10 |\n" +
				"| South | 20 |",
				text);
		}

		[Fact]
		public void TitleWithoutBlankRow_IsSeparatedFromTheTable()
		{
			var builder = new XlsxBuilder().Sheet("Report",
				Data(
					Row(1, T("A1", "Inventory")),
					Row(2, T("A2", "Item"), T("B2", "Qty")),
					Row(3, T("A3", "Bolts"), N("B3", 100))) +
				"<mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>");

			var text = Markdown(builder, o => o.IncludeCellRanges = false);

			Assert.Contains("Inventory\n\n| Item | Qty |\n| --- | --- |\n| Bolts | 100 |", text);
		}

		[Fact]
		public void MergedCells_AreRepeatedOrLeftEmpty()
		{
			var builder = new XlsxBuilder().Sheet("Data",
				Data(
					Row(1, T("A1", "Group"), T("B1", "Item")),
					Row(2, T("A2", "Fruit"), T("B2", "Apple")),
					Row(3, T("B3", "Pear"))) +
				"<mergeCells count=\"1\"><mergeCell ref=\"A2:A3\"/></mergeCells>");

			Assert.Contains("| Fruit | Pear |", Markdown(builder));
			Assert.Contains("|  | Pear |", Markdown(builder, o => o.MergedCells = MergedCellMode.Empty));
		}

		[Fact]
		public void NumericFirstRow_UsesColumnLetters()
		{
			var builder = new XlsxBuilder().Sheet("Numbers", Data(
				Row(2, N("B2", 1), N("C2", 2)),
				Row(3, N("B3", 3), N("C3", 4))));

			var text = Markdown(builder);

			Assert.Contains("**Range B2:C3**\n\n| B | C |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |", text);
		}

		[Fact]
		public void TableHeaders_CanBeForced()
		{
			var builder = new XlsxBuilder().Sheet("Numbers", Data(
				Row(1, N("A1", 2024), N("B1", 2025)),
				Row(2, N("A2", 3), N("B2", 4))));

			Assert.Contains("| 2024 | 2025 |\n| --- | --- |\n| 3 | 4 |", Markdown(builder, o => o.TableHeaders = TableHeaderMode.FirstRow));
			Assert.Contains("| A | B |\n| --- | --- |\n| 2024 | 2025 |", Markdown(builder, o => o.TableHeaders = TableHeaderMode.ColumnLetters));
		}

		[Fact]
		public void SeparateBlocks_BecomeSeparateTables()
		{
			var builder = new XlsxBuilder().Sheet("Dashboard", Data(
				Row(1, T("A1", "Name"), T("B1", "Age"), T("E1", "City"), T("F1", "Zip")),
				Row(2, T("A2", "Ann"), N("B2", 31), T("E2", "Rome"), N("F2", 100)),
				Row(3, T("E3", "Oslo"), N("F3", 150))));

			var text = Markdown(builder);

			Assert.Contains("**Range A1:B2**", text);
			Assert.Contains("**Range E1:F3**", text);
			Assert.True(text.IndexOf("A1:B2") < text.IndexOf("E1:F3"));
		}

		[Fact]
		public void LabelRow_IsTheHeaderOfTheBlocksBelow()
		{
			var builder = new XlsxBuilder()
				.Styles(new[] { 0, 7 })
				.Sheet("Prices", Data(
					Row(2, T("C2", "Retail"), T("D2", "Wholesale")),
					Row(4, T("B4", "Basic")),
					Row(5, T("B5", "1 seat"), N("C5", 100, 1), N("D5", 80, 1)),
					Row(6, T("B6", "5 seats"), N("C6", 450, 1), N("D6", 360, 1)),
					Row(8, T("B8", "Premium")),
					Row(9, T("B9", "1 seat"), N("C9", 200, 1), N("D9", 160, 1)),
					Row(11, T("B11", "Support")),
					Row(12, T("B12", "Email"), T("C12", "included"), T("D12", "included")),
					Row(13, T("B13", "Phone"), T("C13", "on request"), T("D13", "no"))));

			var text = Markdown(builder, o => o.IncludeCellRanges = false);

			Assert.Equal(
				"## Prices\n\n" +
				"Basic\n\n" +
				"| B | Retail | Wholesale |\n| --- | --- | --- |\n| 1 seat | $100.00 | $80.00 |\n| 5 seats | $450.00 | $360.00 |\n\n" +
				"Premium\n\n" +
				"| B | Retail | Wholesale |\n| --- | --- | --- |\n| 1 seat | $200.00 | $160.00 |\n\n" +
				"Support\n\n" +
				"| B | Retail | Wholesale |\n| --- | --- | --- |\n| Email | included | included |\n| Phone | on request | no |",
				text);
		}

		[Fact]
		public void SpacerColumn_KeepsOneTable()
		{
			var builder = new XlsxBuilder().Sheet("Data", Data(
				Row(1, T("A1", "Name"), T("C1", "Age")),
				Row(2, T("A2", "Ann"), N("C2", 31))));

			Assert.Contains("**Range A1:C2**\n\n| Name | Age |\n| --- | --- |\n| Ann | 31 |", Markdown(builder));
		}

		[Fact]
		public void ExcelTable_UsesItsNameAndHeader()
		{
			var builder = new XlsxBuilder()
				.Sheet("Data",
					Data(
						Row(1, T("A1", "Product"), T("B1", "Units")),
						Row(2, T("A2", "Chairs"), N("B2", 12)),
						Row(3, T("A3", "Desks"), N("B3", 4))) +
					"<tableParts count=\"1\"><tablePart r:id=\"rIdT\"/></tableParts>")
				.SheetPart(1, "rIdT", "table", "xl/tables/table1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
					$"<table {XlsxBuilder.Namespaces} id=\"1\" name=\"Table1\" displayName=\"Inventory\" ref=\"A1:B3\">" +
					"<tableColumns count=\"2\"><tableColumn id=\"1\" name=\"Product\"/><tableColumn id=\"2\" name=\"Units\"/></tableColumns></table>");

			var text = Markdown(builder);

			Assert.Contains("**Table: Inventory** (A1:B3)\n\n| Product | Units |\n| --- | --- |\n| Chairs | 12 |\n| Desks | 4 |", text);
		}

		[Fact]
		public void HiddenRowsAndColumns_AreSkipped()
		{
			var builder = new XlsxBuilder().Sheet("Data",
				"<cols><col min=\"2\" max=\"2\" hidden=\"1\"/></cols>" +
				Data(
					Row(1, T("A1", "Name"), T("B1", "Secret"), T("C1", "Age")),
					Row(2, T("A2", "Ann"), T("B2", "x"), N("C2", 31)),
					"<row r=\"3\" hidden=\"1\">" + T("A3", "Bob") + N("C3", 40) + "</row>"));

			var text = Markdown(builder);
			Assert.DoesNotContain("Secret", text);
			Assert.DoesNotContain("Bob", text);

			var all = Markdown(builder, o =>
			{
				o.IncludeHiddenRows = true;
				o.IncludeHiddenColumns = true;
			});
			Assert.Contains("Secret", all);
			Assert.Contains("Bob", all);
		}

		[Fact]
		public void HiddenSheets_AreSkippedButListed()
		{
			var builder = new XlsxBuilder()
				.Sheet("Visible", Data(Row(1, T("A1", "Shown"))))
				.Sheet("Lookup", Data(Row(1, T("A1", "Hidden"))), "hidden");

			var result = builder.Extract(new XlsxExtractorOptions());
			Assert.Equal(1, result.PageCount);
			Assert.Equal(new[] { "Visible", "Lookup" }, result.Metadata.SheetNames);

			Assert.Equal(2, builder.Extract(new XlsxExtractorOptions { IncludeHiddenSheets = true }).PageCount);
		}

		[Fact]
		public void Values_OfEveryType()
		{
			var builder = new XlsxBuilder().Sheet("Types", Data(
				Row(1, T("A1", "Kind"), T("B1", "Value")),
				Row(2, T("A2", "Bool"), "<c r=\"B2\" t=\"b\"><v>1</v></c>"),
				Row(3, T("A3", "Error"), "<c r=\"B3\" t=\"e\"><v>#DIV/0!</v></c>"),
				Row(4, T("A4", "Formula string"), "<c r=\"B4\" t=\"str\"><f>A1&amp;\"x\"</f><v>Kindx</v></c>"),
				Row(5, T("A5", "Date"), "<c r=\"B5\" t=\"d\"><v>2024-02-29T00:00:00</v></c>"),
				Row(6, T("A6", "Multi line"), T("B6", "one_x000D_\ntwo"))));

			var text = Markdown(builder);

			Assert.Contains("| Bool | TRUE |", text);
			Assert.Contains("| Error | #DIV/0! |", text);
			Assert.Contains("| Formula string | Kindx |", text);
			Assert.Contains("| Date | 2024-02-29 |", text);
			Assert.Contains("| Multi line | one<br>two |", text);
		}

		[Fact]
		public void Formulas_AreWrittenWhenRequested()
		{
			var builder = new XlsxBuilder().Sheet("Calc", Data(
				Row(1, T("A1", "Item"), T("B1", "Total")),
				Row(2, T("A2", "Sum"), "<c r=\"B2\"><f>SUM(B3:B4)</f><v>6</v></c>"),
				Row(3, T("A3", "Uncalculated"), "<c r=\"B3\"><f>B4*2</f></c>")));

			var text = Markdown(builder);
			Assert.Contains("| Sum | 6 |", text);
			Assert.Contains("| Uncalculated | =B4\\*2 |", text);

			Assert.Contains("| Sum | 6 (=SUM(B3:B4)) |", Markdown(builder, o => o.IncludeFormulas = true));
		}

		[Fact]
		public void Hyperlinks_BecomeLinks()
		{
			var builder = new XlsxBuilder()
				.Sheet("Links",
					Data(
						Row(1, T("A1", "Site"), T("B1", "Note")),
						Row(2, T("A2", "Wisej"), T("B2", "home")),
						Row(3, "<c r=\"A3\" t=\"str\"><f>HYPERLINK(\"https://example.com/a b\",\"Example\")</f><v>Example</v></c>", T("B3", "formula"))) +
					"<hyperlinks><hyperlink ref=\"A2\" r:id=\"rIdL\"/></hyperlinks>")
				.SheetHyperlink(1, "rIdL", "https://wisej.com/");

			var text = Markdown(builder);

			Assert.Contains("| [Wisej](https://wisej.com/) | home |", text);
			Assert.Contains("| [Example](https://example.com/a%20b) | formula |", text);
			Assert.Contains("\nWisej (https://wisej.com/) | home\n", PlainText(builder));
		}

		[Fact]
		public void Notes_AreListedAfterTheSheet()
		{
			var builder = new XlsxBuilder()
				.Sheet("Data", Data(
					Row(1, T("A1", "Name"), T("B1", "Score")),
					Row(2, T("A2", "Ann"), N("B2", 9))))
				.SheetPart(1, "rIdC", "comments", "xl/comments1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
					$"<comments {XlsxBuilder.Namespaces}><authors><author>Luca</author></authors><commentList>" +
					"<comment ref=\"B2\" authorId=\"0\"><text><r><t>Luca:</t></r><r><t xml:space=\"preserve\">\nChecked twice</t></r></text></comment>" +
					"</commentList></comments>");

			var text = Markdown(builder);

			Assert.EndsWith("**Comments:**\n- **B2** (Luca): Checked twice", text);
			Assert.DoesNotContain("Comments", Markdown(builder, o => o.IncludeComments = false));
		}

		[Fact]
		public void RowsPerPage_SplitsLargeSheetsAndRepeatsTheHeader()
		{
			var rows = Enumerable.Range(2, 5).Select(r => Row(r, T("A" + r, "Item " + r), N("B" + r, r))).ToList();
			rows.Insert(0, Row(1, T("A1", "Item"), T("B1", "Value")));

			var builder = new XlsxBuilder().Sheet("Long", Data(rows.ToArray()));
			var result = builder.Extract(new XlsxExtractorOptions { RowsPerPage = 2 });

			Assert.Equal(3, result.PageCount);
			Assert.Equal("## Long (continued)\n\n**Range A1:B6** (continued)\n\n| Item | Value |\n| --- | --- |\n| Item 4 | 4 |\n| Item 5 | 5 |", result.Pages[1].Text);
			Assert.All(result.Pages, p => Assert.Equal("Long", p.SheetName));
		}

		[Fact]
		public void MaxRowsPerSheet_OmitsTheRest()
		{
			var rows = Enumerable.Range(2, 5).Select(r => Row(r, T("A" + r, "Item " + r), N("B" + r, r))).ToList();
			rows.Insert(0, Row(1, T("A1", "Item"), T("B1", "Value")));

			var text = Markdown(new XlsxBuilder().Sheet("Long", Data(rows.ToArray())), o => o.MaxRowsPerSheet = 2);

			Assert.Contains("| Item 3 | 3 |", text);
			Assert.DoesNotContain("Item 4", text);
			Assert.EndsWith("*(3 more rows not included)*", text);
		}

		[Fact]
		public void RowNumbers_AreWrittenWhenRequested()
		{
			var text = Markdown(SalesWorkbook(), o => o.IncludeRowNumbers = true);

			Assert.Contains("| Row | Name | Price | Date |\n| --- | --- | --- | --- |\n| 2 | Apple |", text);
		}

		[Fact]
		public void MarkdownCharacters_AreEscaped()
		{
			var builder = new XlsxBuilder().Sheet("Text", Data(
				Row(1, T("A1", "# Not a heading")),
				Row(2, T("A2", "*not emphasis* and <b>tags</b>"))));

			var text = Markdown(builder);

			Assert.Contains("\\# Not a heading", text);
			Assert.Contains("\\*not emphasis\\* and &lt;b>tags&lt;/b>", text);
		}

		[Fact]
		public void EmptySheet_HasOnlyItsHeading()
		{
			var text = Markdown(new XlsxBuilder().Sheet("Empty", "<sheetData/>"));

			Assert.Equal("## Empty", text);
		}
	}
}
