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
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wisej.AI.XlsX.Model;
using Wisej.AI.XlsX.Packaging;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// Reads a worksheet part: cell values, hidden rows and columns, merged ranges, hyperlinks,
	/// Excel tables, drawings and comments.
	/// </summary>
	/// <remarks>
	/// The cells are streamed one row at a time, so that a worksheet with millions of cells is never
	/// loaded as a whole; the other elements are small and are loaded with LINQ to XML.
	/// </remarks>
	internal sealed class WorksheetReader
	{
		/// <summary>The largest range a hyperlink is applied to cell by cell.</summary>
		private const int MaxEnumeratedRange = 10000;

		/// <summary>Matches the URL of a <c>HYPERLINK</c> formula with a literal target.</summary>
		private static readonly Regex HyperlinkFormula = new Regex("^\\s*HYPERLINK\\(\\s*\"([^\"]+)\"", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

		private readonly ReadContext _context;
		private readonly OpenXmlPart _part;
		private readonly Sheet _sheet;
		private readonly List<(CellRange Range, string Target)> _hyperlinks = new List<(CellRange, string)>();
		private XNamespace _ns = Ns.X;

		/// <summary>
		/// Initializes a new instance of the <see cref="WorksheetReader"/> class.
		/// </summary>
		/// <param name="context">The read context.</param>
		/// <param name="part">The worksheet part.</param>
		/// <param name="sheet">The sheet that receives the content.</param>
		public WorksheetReader(ReadContext context, OpenXmlPart part, Sheet sheet)
		{
			_context = context;
			_part = part;
			_sheet = sheet;
		}

		/// <summary>
		/// Reads the worksheet.
		/// </summary>
		public void Read()
		{
			try
			{
				ReadPart();
			}
			catch (XmlException)
			{
				// A damaged worksheet keeps the cells read before the error.
			}

			ApplyHyperlinks();
			CommentReader.Read(_context, _part, _sheet);
		}

		#region Part

		/// <summary>
		/// Streams the elements of the worksheet part.
		/// </summary>
		private void ReadPart()
		{
			using (var reader = XlsxPackage.OpenReader(_part))
			{
				if (reader.MoveToContent() != XmlNodeType.Element || reader.IsEmptyElement)
					return;

				_ns = reader.NamespaceURI;
				reader.Read();

				while (!reader.EOF)
				{
					if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1)
					{
						reader.Read();
						continue;
					}

					if (reader.NamespaceURI != _ns.NamespaceName)
					{
						reader.Skip();
						continue;
					}

					switch (reader.LocalName)
					{
						case "sheetData":
							ReadSheetData(reader);
							break;

						case "cols":
						case "mergeCells":
						case "hyperlinks":
						case "drawing":
						case "tableParts":
							var element = (XElement)XNode.ReadFrom(reader);
							XlsxPackage.NormalizeStrictNamespaces(element);
							ReadElement(element);
							break;

						default:
							reader.Skip();
							break;
					}
				}
			}
		}

		/// <summary>
		/// Reads an element of the worksheet other than the cells.
		/// </summary>
		/// <param name="element">The element, with Strict namespaces normalized.</param>
		private void ReadElement(XElement element)
		{
			var options = _context.Options;

			switch (element.Name.LocalName)
			{
				case "cols":
					if (options.IncludeHiddenColumns)
						break;

					foreach (var col in element.Elements(Ns.X + "col"))
					{
						if (!XmlExtensions.ParseBool(col.Attr("hidden"), false))
							continue;

						var min = Math.Max(1, XmlExtensions.ParseInt(col.Attr("min")) ?? 1);
						var max = Math.Min(CellRange.MaxColumn, XmlExtensions.ParseInt(col.Attr("max")) ?? min);
						for (var c = min; c <= max; c++)
							_sheet.HiddenColumns.Add(c);
					}
					break;

				case "mergeCells":
					foreach (var merge in element.Elements(Ns.X + "mergeCell"))
					{
						if (CellRange.TryParse(merge.Attr("ref"), out var range) && range.CellCount > 1)
							_sheet.MergedRanges.Add(range);
					}
					break;

				case "hyperlinks":
					foreach (var hyperlink in element.Elements(Ns.X + "hyperlink"))
					{
						var target = XlsxPackage.GetExternalTarget(_part, hyperlink.Attr(Ns.R + "id"));
						if (target != null && CellRange.TryParse(hyperlink.Attr("ref"), out var range))
							_hyperlinks.Add((range, target));
					}
					break;

				case "drawing":
					var drawingPart = XlsxPackage.GetPartById(_part, element.Attr(Ns.R + "id"));
					if (drawingPart != null)
						DrawingReader.Read(_context, drawingPart, _sheet);
					break;

				case "tableParts":
					foreach (var tablePart in element.Elements(Ns.X + "tablePart"))
						ReadTable(XlsxPackage.GetPartById(_part, tablePart.Attr(Ns.R + "id")));
					break;
			}
		}

		/// <summary>
		/// Reads an Excel table (list object).
		/// </summary>
		/// <param name="part">The table part. Can be <c>null</c>.</param>
		private void ReadTable(OpenXmlPart part)
		{
			var root = _context.Package.GetXml(part);
			if (root == null || !CellRange.TryParse(root.Attr("ref"), out var range))
				return;

			var table = new SheetTable
			{
				Name = root.Attr("displayName") ?? root.Attr("name"),
				Range = range,
				HeaderRowCount = Math.Max(0, Math.Min(1, XmlExtensions.ParseInt(root.Attr("headerRowCount")) ?? 1)),
				TotalsRowCount = Math.Max(0, Math.Min(1, XmlExtensions.ParseInt(root.Attr("totalsRowCount")) ?? 0)),
			};

			foreach (var column in root.Element(Ns.X + "tableColumns")?.Elements(Ns.X + "tableColumn") ?? Enumerable.Empty<XElement>())
				table.ColumnNames.Add(XmlExtensions.DecodeEscapes(column.Attr("name") ?? ""));

			_sheet.Tables.Add(table);
		}

		/// <summary>
		/// Assigns the hyperlinks to the cells they cover.
		/// </summary>
		private void ApplyHyperlinks()
		{
			foreach (var (range, target) in _hyperlinks)
			{
				if (range.CellCount <= MaxEnumeratedRange)
				{
					for (var row = range.FirstRow; row <= range.LastRow; row++)
					{
						for (var column = range.FirstColumn; column <= range.LastColumn; column++)
						{
							var cell = _sheet.GetCell(row, column);
							if (cell != null && cell.Hyperlink == null)
								cell.Hyperlink = target;
						}
					}
				}
				else
				{
					foreach (var cell in _sheet.Cells.Values.Where(c => c.Hyperlink == null && range.Contains(c.Row, c.Column)))
						cell.Hyperlink = target;
				}
			}
		}

		#endregion

		#region Cells

		/// <summary>
		/// Reads the rows of <c>sheetData</c>, leaving the reader after its end tag.
		/// </summary>
		/// <param name="reader">The reader, positioned on the <c>sheetData</c> element.</param>
		private void ReadSheetData(XmlReader reader)
		{
			if (reader.IsEmptyElement)
			{
				reader.Read();
				return;
			}

			var depth = reader.Depth;
			var lastRow = 0;
			reader.Read();

			while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
			{
				if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row" && reader.NamespaceURI == _ns.NamespaceName)
				{
					var row = (XElement)XNode.ReadFrom(reader);
					lastRow = ReadRow(row, lastRow);
				}
				else
				{
					reader.Read();
				}
			}

			reader.Read();
		}

		/// <summary>
		/// Reads the cells of a row.
		/// </summary>
		/// <param name="row">The <c>row</c> element.</param>
		/// <param name="previousRow">The number of the previous row, used when the row has no number.</param>
		/// <returns>The number of the row.</returns>
		private int ReadRow(XElement row, int previousRow)
		{
			var number = XmlExtensions.ParseInt(row.Attr("r")) ?? previousRow + 1;
			if (number < 1 || number > CellRange.MaxRow)
				return previousRow;

			if (!_context.Options.IncludeHiddenRows && XmlExtensions.ParseBool(row.Attr("hidden"), false))
			{
				_sheet.HiddenRows.Add(number);
				return number;
			}

			var column = 0;
			foreach (var c in row.Elements(_ns + "c"))
			{
				column = CellRange.TryParseCell(c.Attr("r"), out _, out var parsed) && parsed > 0 ? parsed : column + 1;
				if (column > CellRange.MaxColumn || _sheet.HiddenColumns.Contains(column))
					continue;

				var cell = ReadCell(c, number, column);
				if (cell != null)
					_sheet.Cells[Sheet.Key(number, column)] = cell;
			}

			return number;
		}

		/// <summary>
		/// Reads a cell and formats its value as Excel displays it.
		/// </summary>
		/// <param name="c">The <c>c</c> element.</param>
		/// <param name="row">The one based row.</param>
		/// <param name="column">The one based column.</param>
		/// <returns>The cell, or <c>null</c> when it has no value.</returns>
		private Cell ReadCell(XElement c, int row, int column)
		{
			var type = c.Attr("t") ?? "n";
			var value = c.Element(_ns + "v")?.Value;
			var formula = c.Element(_ns + "f")?.Value;
			var style = XmlExtensions.ParseInt(c.Attr("s")) ?? 0;
			var formatter = _context.Formatter;

			var cell = new Cell
			{
				Row = row,
				Column = column,
				Kind = CellKind.Text,
				Formula = String.IsNullOrWhiteSpace(formula) ? null : formula.Trim(),
			};

			switch (type)
			{
				case "s":
					if (value != null)
						cell.Text = formatter.FormatText(_context.Strings[XmlExtensions.ParseInt(value) ?? -1], _context.Styles.GetFormatCode(style));
					break;

				case "inlineStr":
					cell.Text = XmlExtensions.RichText(c.Element(_ns + "is"), _ns);
					break;

				case "str":
					cell.Text = XmlExtensions.DecodeEscapes(value);
					break;

				case "b":
					if (value != null)
					{
						cell.Kind = CellKind.Boolean;
						cell.Text = value.Trim() == "1" || value.Trim().Equals("true", StringComparison.OrdinalIgnoreCase) ? "TRUE" : "FALSE";
					}
					break;

				case "e":
					// Pictures placed in cells are stored as #VALUE! errors with value metadata.
					var vm = XmlExtensions.ParseInt(c.Attr("vm"));
					var image = vm != null ? _context.CellImages?.GetImage(vm.Value, _sheet.Name, CellRange.CellName(row, column)) : null;
					if (image != null)
					{
						cell.Image = image;
					}
					else if (value != null)
					{
						cell.Kind = CellKind.Error;
						cell.Text = value;
					}
					break;

				case "d":
					if (value != null && DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date))
					{
						var code = _context.Styles.GetFormatCode(style);
						if (!formatter.IsDateFormat(code))
							code = date.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss";

						cell.Kind = CellKind.Date;
						cell.Text = formatter.Format(formatter.ToSerial(date), code);
					}
					else
					{
						cell.Text = value;
					}
					break;

				default:
					if (XmlExtensions.ParseDouble(value) is double number)
					{
						var code = _context.Styles.GetFormatCode(style);
						cell.Kind = formatter.IsDateFormat(code) ? CellKind.Date : CellKind.Number;
						cell.Text = formatter.Format(number, code);
					}
					break;
			}

			cell.Text = cell.Text?.Trim();

			// A formula without a cached value (a workbook that was never calculated) is the only content of the cell.
			if (String.IsNullOrEmpty(cell.Text) && cell.Image == null && cell.Formula != null)
			{
				cell.Text = "=" + cell.Formula;
				cell.Formula = null;
			}

			if (cell.Formula != null && _context.Options.IncludeHyperlinks)
			{
				var match = HyperlinkFormula.Match(cell.Formula);
				if (match.Success)
					cell.Hyperlink = match.Groups[1].Value;
			}

			return cell.HasContent ? cell : null;
		}

		#endregion
	}
}
