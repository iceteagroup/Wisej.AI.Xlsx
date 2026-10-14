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
using System.Text;
using System.Text.RegularExpressions;
using Wisej.AI.XlsX.Layout;
using Wisej.AI.XlsX.Model;

namespace Wisej.AI.XlsX.Rendering
{
	/// <summary>
	/// Writes the workbook as plain text: tables are rows of pipe-separated cells.
	/// </summary>
	internal sealed class PlainTextRenderer : WorkbookRenderer
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="PlainTextRenderer"/> class.
		/// </summary>
		/// <param name="model">The workbook model.</param>
		/// <param name="options">The extraction options.</param>
		public PlainTextRenderer(WorkbookModel model, XlsxExtractorOptions options)
			: base(model, options)
		{
		}

		/// <inheritdoc/>
		protected override string DefaultPageHeaderFormat => "--- Page {0} ---";

		/// <inheritdoc/>
		protected override string FormatSheetHeading(Sheet sheet, bool continued)
		{
			return "Sheet: " + sheet.Name + (continued ? " (continued)" : "");
		}

		/// <inheritdoc/>
		protected override string FormatParagraph(Cell cell)
		{
			return FormatValue(cell, cell.Text);
		}

		/// <inheritdoc/>
		protected override string FormatTableCaption(TableBlock table, bool continued)
		{
			var range = Options.IncludeCellRanges ? table.Range.ToString() : null;
			var suffix = continued ? " (continued)" : "";

			if (table.Name != null)
				return "Table: " + table.Name + (range != null ? $" ({range})" : "") + suffix;

			return range != null ? "Range " + range + suffix : null;
		}

		/// <inheritdoc/>
		protected override IEnumerable<string> FormatTableHeader(List<string> labels)
		{
			yield return FormatTableRow(labels.Select(Collapse).ToList());
		}

		/// <inheritdoc/>
		protected override string FormatTableRow(List<string> cells)
		{
			return String.Join(" | ", cells).Trim();
		}

		/// <inheritdoc/>
		protected override string FormatCell(Cell cell)
		{
			return FormatValue(cell, Collapse(cell.Text ?? ""));
		}

		/// <inheritdoc/>
		protected override string FormatOmittedRows(int count)
		{
			return $"({count.ToString("N0", CultureInfo.InvariantCulture)} more rows not included)";
		}

		/// <inheritdoc/>
		protected override string FormatChart(ChartData chart, string cell)
		{
			var sb = new StringBuilder("Chart");
			if (chart.Title != null)
				sb.Append(": ").Append(chart.Title);

			var details = ChartDetails(chart, cell);
			if (details.Length > 0)
				sb.Append(" (").Append(details).Append(')');

			sb.Append('\n').Append(chart.CategoryLabel).Append(" | ").Append(String.Join(" | ", chart.Series.Select(s => s.Name)));

			for (var i = 0; i < chart.Categories.Count; i++)
			{
				sb.Append('\n').Append(chart.Categories[i]);
				foreach (var series in chart.Series)
					sb.Append(" | ").Append(i < series.Values.Count ? series.Values[i] : "");
			}

			return sb.ToString();
		}

		/// <inheritdoc/>
		protected override string FormatImage(XlsxImage image)
		{
			var alt = image.AltText ?? image.Title;
			return alt == null ? "[Image]" : "[Image: " + alt.Replace('\n', ' ') + "]";
		}

		/// <inheritdoc/>
		protected override string FormatImageDescription(XlsxImage image)
		{
			return $"[{image.FileName}] " + image.Description.Trim();
		}

		/// <inheritdoc/>
		protected override string FormatShape(string text, string cell)
		{
			return "Text box" + (cell != null ? " (at " + cell + ")" : "") + ": " + text;
		}

		/// <inheritdoc/>
		protected override string FormatComments(List<SheetComment> comments)
		{
			var sb = new StringBuilder("Comments:");

			foreach (var comment in comments)
			{
				sb.Append("\n- ").Append(comment.Cell);
				if (comment.Author != null)
					sb.Append(" (").Append(comment.Author).Append(')');
				sb.Append(": ").Append(Collapse(comment.Text));

				foreach (var reply in comment.Replies)
				{
					sb.Append("\n  - ");
					if (reply.Author != null)
						sb.Append(reply.Author).Append(": ");
					sb.Append(Collapse(reply.Text));
				}
			}

			return sb.ToString();
		}

		/// <summary>
		/// Formats the value of a cell with its picture, hyperlink and formula.
		/// </summary>
		/// <param name="cell">The cell.</param>
		/// <param name="text">The text of the value.</param>
		/// <returns>The formatted value.</returns>
		private string FormatValue(Cell cell, string text)
		{
			if (cell.Image != null)
				return FormatImage(cell.Image) + (cell.Image.Description != null ? " (" + Collapse(cell.Image.Description) + ")" : "");

			if (Options.IncludeHyperlinks && !String.IsNullOrWhiteSpace(cell.Hyperlink) && cell.Hyperlink.Trim() != text)
				text += " (" + cell.Hyperlink.Trim() + ")";

			if (Options.IncludeFormulas && cell.Formula != null)
				text += " (=" + Collapse(cell.Formula) + ")";

			return text;
		}

		/// <summary>
		/// Collapses line breaks into spaces.
		/// </summary>
		/// <param name="text">The text.</param>
		/// <returns>The text on a single line.</returns>
		private static string Collapse(string text)
		{
			return Regex.Replace(text ?? "", @"\s*\n\s*", " ").Trim();
		}
	}
}
