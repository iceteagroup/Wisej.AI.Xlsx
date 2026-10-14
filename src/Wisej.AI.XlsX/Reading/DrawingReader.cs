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

using System;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wisej.AI.XlsX.Model;
using Wisej.AI.XlsX.Packaging;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// Reads the objects of a spreadsheet drawing part (<c>xdr:wsDr</c>): charts, pictures and shapes with text.
	/// </summary>
	internal static class DrawingReader
	{
		/// <summary>
		/// The maximum nesting depth of shape groups. Deeper content is ignored instead of exhausting the stack.
		/// </summary>
		private const int MaxNestingDepth = 32;

		/// <summary>
		/// Reads the objects of a drawing part into <see cref="Sheet.Drawings"/>.
		/// </summary>
		/// <param name="context">The read context.</param>
		/// <param name="part">The drawing part.</param>
		/// <param name="sheet">The sheet that displays the drawing.</param>
		public static void Read(ReadContext context, OpenXmlPart part, Sheet sheet)
		{
			var root = context.Package.GetXml(part);
			if (root == null)
				return;

			foreach (var anchor in root.ContentElements())
			{
				var from = anchor.Element(Ns.Xdr + "from");
				var row = from == null ? 0 : (XmlExtensions.ParseInt(from.Element(Ns.Xdr + "row")?.Value) ?? 0) + 1;
				var column = from == null ? 0 : (XmlExtensions.ParseInt(from.Element(Ns.Xdr + "col")?.Value) ?? 0) + 1;

				foreach (var item in anchor.ContentElements())
					ReadObject(context, part, sheet, item, row, column, 0);
			}
		}

		/// <summary>
		/// Reads a drawing object.
		/// </summary>
		/// <param name="context">The read context.</param>
		/// <param name="part">The drawing part.</param>
		/// <param name="sheet">The sheet that displays the drawing.</param>
		/// <param name="element">The object element.</param>
		/// <param name="row">The one based row of the anchor, or zero.</param>
		/// <param name="column">The one based column of the anchor, or zero.</param>
		/// <param name="depth">The nesting depth in shape groups.</param>
		private static void ReadObject(ReadContext context, OpenXmlPart part, Sheet sheet, XElement element, int row, int column, int depth)
		{
			var options = context.Options;
			var cell = row > 0 && column > 0 ? CellRange.CellName(row, column) : null;

			switch (element.Name.LocalName)
			{
				case "pic" when element.Name.Namespace == Ns.Xdr:
				{
					if (!options.IncludeImages)
						return;

					var properties = element.Element(Ns.Xdr + "nvPicPr")?.Element(Ns.Xdr + "cNvPr");
					if (IsHidden(properties) || (options.SkipDecorativeImages && IsDecorative(properties)))
						return;

					var blip = element.Element(Ns.Xdr + "blipFill")?.Element(Ns.A + "blip");
					var image = context.Images.Create(part, blip.Attr(Ns.R + "embed"), blip.Attr(Ns.R + "link"), sheet.Name);
					if (image == null)
						return;

					var extent = element.Element(Ns.Xdr + "spPr")?.Element(Ns.A + "xfrm")?.Element(Ns.A + "ext");
					image.AltText = Clean(properties.Attr("descr"));
					image.Title = Clean(properties.Attr("title"));
					image.Name = Clean(properties.Attr("name"));
					image.Width = XmlExtensions.EmuToPoints(extent.Attr("cx")) ?? 0;
					image.Height = XmlExtensions.EmuToPoints(extent.Attr("cy")) ?? 0;
					image.Cell = cell;

					sheet.Drawings.Add(new PictureItem { Image = image, Cell = cell, Row = row, Column = column });
					break;
				}

				case "graphicFrame" when element.Name.Namespace == Ns.Xdr:
				{
					if (!options.IncludeCharts)
						return;

					var properties = element.Element(Ns.Xdr + "nvGraphicFramePr")?.Element(Ns.Xdr + "cNvPr");
					if (IsHidden(properties))
						return;

					var data = element.Element(Ns.A + "graphic")?.Element(Ns.A + "graphicData");
					ChartData chart = null;

					if (data.Attr("uri") == Ns.GraphicChart)
					{
						var chartPart = XlsxPackage.GetPartById(part, data.Element(Ns.C + "chart").Attr(Ns.R + "id"));
						chart = ChartReader.Read(context.Package.GetXml(chartPart), context.Formatter);
					}
					else if (data.Attr("uri") == Ns.GraphicChartEx)
					{
						var chartPart = XlsxPackage.GetPartById(part, data.Element(Ns.CX + "chart").Attr(Ns.R + "id"));
						chart = ChartReader.ReadExtended(context.Package.GetXml(chartPart), context.Formatter);
					}

					if (chart != null)
						sheet.Drawings.Add(new ChartItem { Chart = chart, Cell = cell, Row = row, Column = column });
					break;
				}

				case "sp" when element.Name.Namespace == Ns.Xdr:
				{
					if (!options.IncludeTextBoxes || IsHidden(element.Element(Ns.Xdr + "nvSpPr")?.Element(Ns.Xdr + "cNvPr")))
						return;

					var paragraphs = element.Element(Ns.Xdr + "txBody")?.Elements(Ns.A + "p")
						.Select(p => String.Concat(p.Descendants().Where(d => d.Name == Ns.A + "t").Select(t => t.Value)).Trim())
						.Where(t => t.Length > 0)
						.ToList();

					if (paragraphs != null && paragraphs.Count > 0)
						sheet.Drawings.Add(new ShapeItem { Text = String.Join("\n", paragraphs), Cell = cell, Row = row, Column = column });
					break;
				}

				case "grpSp" when element.Name.Namespace == Ns.Xdr:
				{
					if (depth >= MaxNestingDepth)
						return;

					foreach (var child in element.ContentElements())
						ReadObject(context, part, sheet, child, row, column, depth + 1);
					break;
				}
			}
		}

		/// <summary>
		/// Determines whether a drawing object is hidden.
		/// </summary>
		/// <param name="properties">The <c>cNvPr</c> element. Can be <c>null</c>.</param>
		/// <returns><c>true</c> when the object is hidden.</returns>
		private static bool IsHidden(XElement properties)
		{
			return XmlExtensions.ParseBool(properties.Attr("hidden"), false);
		}

		/// <summary>
		/// Determines whether a picture is marked as decorative by the author.
		/// </summary>
		/// <param name="properties">The <c>cNvPr</c> element. Can be <c>null</c>.</param>
		/// <returns><c>true</c> for a decorative picture.</returns>
		private static bool IsDecorative(XElement properties)
		{
			return properties != null && properties.Descendants(Ns.ADec + "decorative").Any(d => XmlExtensions.ParseBool(d.Attr("val"), true));
		}

		/// <summary>
		/// Trims a property value, returning <c>null</c> for empty values.
		/// </summary>
		/// <param name="value">The value. Can be <c>null</c>.</param>
		/// <returns>The trimmed value or <c>null</c>.</returns>
		private static string Clean(string value)
		{
			return String.IsNullOrWhiteSpace(value) ? null : value.Trim();
		}
	}
}
