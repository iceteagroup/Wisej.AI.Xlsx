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

using Xunit;

namespace Wisej.AI.XlsX.Tests
{
	/// <summary>
	/// Tests the extraction of charts, pictures, text boxes and pictures placed in cells.
	/// </summary>
	public class DrawingTests : TestBase
	{
		private const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
		private const string ChartRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

		private static readonly byte[] Png = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3 };

		/// <summary>
		/// Returns a workbook with a drawing that contains a picture, a chart and a text box.
		/// </summary>
		private static XlsxBuilder DrawingWorkbook()
		{
			var drawing =
				$"<xdr:wsDr {XlsxBuilder.DrawingNamespaces}>" +
				"<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
				"<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
				"<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 1\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>" +
				"<xdr:xfrm/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"rId2\"/></a:graphicData></a:graphic></xdr:graphicFrame>" +
				"<xdr:clientData/></xdr:twoCellAnchor>" +
				"<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"1270000\" cy=\"635000\"/>" +
				"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"3\" name=\"Picture 2\" descr=\"Company logo\"/><xdr:cNvPicPr/></xdr:nvPicPr>" +
				"<xdr:blipFill><a:blip r:embed=\"rId1\"/></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1270000\" cy=\"635000\"/></a:xfrm></xdr:spPr></xdr:pic>" +
				"<xdr:clientData/></xdr:oneCellAnchor>" +
				"<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"1270000\" cy=\"635000\"/>" +
				"<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"4\" name=\"TextBox 3\"/><xdr:cNvSpPr txBox=\"1\"/></xdr:nvSpPr><xdr:spPr/>" +
				"<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Figures are preliminary</a:t></a:r></a:p></xdr:txBody></xdr:sp>" +
				"<xdr:clientData/></xdr:oneCellAnchor>" +
				"</xdr:wsDr>";

			var chart =
				"<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
				"<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea>" +
				"<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" +
				"<c:ser><c:idx val=\"0\"/><c:tx><c:strRef><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>2025</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
				"<c:cat><c:strRef><c:strCache><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>Q1</c:v></c:pt><c:pt idx=\"1\"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
				"<c:val><c:numRef><c:numCache><c:formatCode>#,##0</c:formatCode><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>1500</c:v></c:pt><c:pt idx=\"1\"><c:v>2250.4</c:v></c:pt></c:numCache></c:numRef></c:val>" +
				"</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";

			return new XlsxBuilder()
				.Sheet("Report", Data(Row(1, T("A1", "Summary"))) + "<drawing r:id=\"rIdD\"/>")
				.SheetPart(1, "rIdD", "drawing", "xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml", drawing)
				.PartRelationships("xl/drawings/drawing1.xml",
					$"<Relationship Id=\"rId1\" Type=\"{ImageRelationship}\" Target=\"../media/image1.png\"/>" +
					$"<Relationship Id=\"rId2\" Type=\"{ChartRelationship}\" Target=\"../charts/chart1.xml\"/>")
				.Part("xl/media/image1.png", null, Png)
				.Part("xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", System.Text.Encoding.UTF8.GetBytes(chart));
		}

		[Fact]
		public void Chart_BecomesADataTable()
		{
			var text = Markdown(DrawingWorkbook());

			Assert.Contains("**Chart: Revenue** (column chart; at D2)\n\n| Category | 2025 |\n| --- | --- |\n| Q1 | 1,500 |\n| Q2 | 2,250 |", text);
		}

		[Fact]
		public void Picture_IsReferencedAndListed()
		{
			var result = DrawingWorkbook().Extract(new XlsxExtractorOptions { PageHeaderFormat = "" });

			Assert.Contains("![Company logo](image1.png)", result.Text);

			var image = Assert.Single(result.Images);
			Assert.Equal("image1.png", image.FileName);
			Assert.Equal("A6", image.Cell);
			Assert.Equal("Report", image.SheetName);
			Assert.Equal(1, image.PageNumber);
			Assert.Equal(100, image.Width, 3);
			Assert.Equal(Png, image.Data);
			Assert.False(image.IsInCell);
			Assert.Same(image, Assert.Single(result.Pages[0].Images));
		}

		[Fact]
		public void TextBox_IsIncluded()
		{
			Assert.Contains("**Text box (at A10):** Figures are preliminary", Markdown(DrawingWorkbook()));
		}

		[Fact]
		public void Drawings_FollowAnchorOrder()
		{
			var text = Markdown(DrawingWorkbook());

			Assert.True(text.IndexOf("Chart: Revenue") < text.IndexOf("Company logo"));
			Assert.True(text.IndexOf("Company logo") < text.IndexOf("Figures are preliminary"));
		}

		[Fact]
		public void ImageDescriber_IsCalledOncePerPicture()
		{
			var calls = 0;
			var result = DrawingWorkbook().Extract(new XlsxExtractorOptions
			{
				PageHeaderFormat = "",
				IncludeImageData = false,
				ImageDescriber = image =>
				{
					calls++;
					Assert.NotNull(image.Data);
					return "A blue logo";
				},
			});

			Assert.Equal(1, calls);
			Assert.Contains("> **image1.png:** A blue logo", result.Text);
			Assert.Null(result.Images[0].Data);
		}

		[Fact]
		public void Drawings_CanBeExcluded()
		{
			var text = Markdown(DrawingWorkbook(), o =>
			{
				o.IncludeCharts = false;
				o.IncludeImages = false;
				o.IncludeTextBoxes = false;
			});

			Assert.Equal("## Report\n\nSummary", text);
		}

		[Fact]
		public void PictureInCell_IsResolvedThroughRichValues()
		{
			const string richData = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";

			var builder = new XlsxBuilder()
				.Sheet("Catalog", Data(
					Row(1, T("A1", "Product"), T("B1", "Photo")),
					Row(2, T("A2", "Apple"), "<c r=\"B2\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c>")))
				.WorkbookPart("rIdM", "sheetMetadata", "xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
					$"<metadata {XlsxBuilder.Namespaces} xmlns:xlrd=\"{richData}\">" +
					"<metadataTypes count=\"1\"><metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\"/></metadataTypes>" +
					"<futureMetadata name=\"XLRICHVALUE\" count=\"1\"><bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\"><xlrd:rvb i=\"0\"/></ext></extLst></bk></futureMetadata>" +
					"<valueMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></valueMetadata></metadata>")
				.WorkbookPart("rIdRV", "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue", "xl/richData/rdrichvalue.xml", "application/vnd.ms-excel.rdrichvalue+xml",
					$"<rvData xmlns=\"{richData}\" count=\"1\"><rv s=\"0\"><v>0</v><v>5</v><v>A red apple</v></rv></rvData>")
				.WorkbookPart("rIdRS", "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure", "xl/richData/rdrichvaluestructure.xml", "application/vnd.ms-excel.rdrichvaluestructure+xml",
					$"<rvStructures xmlns=\"{richData}\" count=\"1\"><s t=\"_localImage\"><k n=\"_rvRel:LocalImageIdentifier\" t=\"i\"/><k n=\"CalcOrigin\" t=\"i\"/><k n=\"Text\" t=\"s\"/></s></rvStructures>")
				.WorkbookPart("rIdRR", "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel", "xl/richData/richValueRel.xml", "application/vnd.ms-excel.richvaluerel+xml",
					"<richValueRels xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><rel r:id=\"rId1\"/></richValueRels>")
				.PartRelationships("xl/richData/richValueRel.xml", $"<Relationship Id=\"rId1\" Type=\"{ImageRelationship}\" Target=\"../media/image1.png\"/>")
				.Part("xl/media/image1.png", null, Png);

			var result = builder.Extract(new XlsxExtractorOptions { PageHeaderFormat = "" });

			Assert.Contains("| Apple | ![A red apple](image1.png) |", result.Text);

			var image = Assert.Single(result.Images);
			Assert.True(image.IsInCell);
			Assert.Equal("B2", image.Cell);
		}
	}
}
