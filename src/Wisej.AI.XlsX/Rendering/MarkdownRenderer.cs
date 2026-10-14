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
	/// Writes the workbook as GitHub flavored Markdown.
	/// </summary>
	internal sealed class MarkdownRenderer : WorkbookRenderer
	{
		/// <summary>Matches line starts that Markdown would read as a heading, a quote, a list item or a rule.</summary>
		private static readonly Regex BlockMarker = new Regex(@"^(\s*)(#{1,6}(?=\s|$)|>|[+\-](?=\s|$)|[-=]+\s*$)", RegexOptions.CultureInvariant);

		/// <summary>Matches line starts that Markdown would read as an ordered list item.</summary>
		private static readonly Regex OrderedMarker = new Regex(@"^(\s*\d{1,9})([.)])(?=\s|$)", RegexOptions.CultureInvariant);

		/// <summary>
		/// Initializes a new instance of the <see cref="MarkdownRenderer"/> class.
		/// </summary>
		/// <param name="model">The workbook model.</param>
		/// <param name="options">The extraction options.</param>
		public MarkdownRenderer(WorkbookModel model, XlsxExtractorOptions options)
			: base(model, options)
		{
		}

		/// <inheritdoc/>
		protected override string DefaultPageHeaderFormat => "<!-- Page {0} -->";

		#region Blocks

		/// <inheritdoc/>
		protected override string FormatSheetHeading(Sheet sheet, bool continued)
		{
			return "## " + Escape(sheet.Name, InlineMode.Flat) + (continued ? " (continued)" : "");
		}

		/// <inheritdoc/>
		protected override string FormatParagraph(Cell cell)
		{
			if (cell.Image != null)
				return FormatImage(cell.Image) + (cell.Image.Description != null ? " " + Inline(cell.Image.Description) : "");

			var text = Link(EscapeLineStarts(Escape(cell.Text, InlineMode.Block)), cell.Hyperlink);
			return AppendFormula(text, cell, InlineMode.Block);
		}

		/// <inheritdoc/>
		protected override string FormatTableCaption(TableBlock table, bool continued)
		{
			var range = Options.IncludeCellRanges ? table.Range.ToString() : null;
			var suffix = continued ? " (continued)" : "";

			if (table.Name != null)
				return $"**Table: {Escape(table.Name, InlineMode.Flat)}**" + (range != null ? $" ({range})" : "") + suffix;

			return range != null ? $"**Range {range}**{suffix}" : null;
		}

		/// <inheritdoc/>
		protected override IEnumerable<string> FormatTableHeader(List<string> labels)
		{
			yield return FormatTableRow(labels.Select(l => Escape(Collapse(l), InlineMode.Cell)).ToList());
			yield return "|" + String.Concat(Enumerable.Repeat(" --- |", labels.Count));
		}

		/// <inheritdoc/>
		protected override string FormatTableRow(List<string> cells)
		{
			return "| " + String.Join(" | ", cells) + " |";
		}

		/// <inheritdoc/>
		protected override string FormatCell(Cell cell)
		{
			if (cell.Image != null)
				return FormatImage(cell.Image) + (cell.Image.Description != null ? " " + Escape(Inline(cell.Image.Description), InlineMode.Cell) : "");

			var text = Link(Escape(cell.Text, InlineMode.Cell), cell.Hyperlink);
			return AppendFormula(text, cell, InlineMode.Cell);
		}

		/// <inheritdoc/>
		protected override string FormatOmittedRows(int count)
		{
			return $"*({count.ToString("N0", CultureInfo.InvariantCulture)} more rows not included)*";
		}

		/// <inheritdoc/>
		protected override string FormatChart(ChartData chart, string cell)
		{
			var sb = new StringBuilder();
			sb.Append("**Chart").Append(chart.Title != null ? ": " + Escape(chart.Title, InlineMode.Flat) : "").Append("**");

			var details = ChartDetails(chart, cell);
			if (details.Length > 0)
				sb.Append(" (").Append(Escape(details, InlineMode.Flat)).Append(')');

			sb.Append("\n\n| ").Append(TableCell(chart.CategoryLabel));
			foreach (var series in chart.Series)
				sb.Append(" | ").Append(TableCell(series.Name));
			sb.Append(" |\n|").Append(String.Concat(Enumerable.Repeat(" --- |", chart.Series.Count + 1)));

			for (var i = 0; i < chart.Categories.Count; i++)
			{
				sb.Append("\n| ").Append(TableCell(chart.Categories[i]));
				foreach (var series in chart.Series)
					sb.Append(" | ").Append(TableCell(i < series.Values.Count ? series.Values[i] : ""));
				sb.Append(" |");
			}

			return sb.ToString();
		}

		/// <inheritdoc/>
		protected override string FormatImage(XlsxImage image)
		{
			var alt = (image.AltText ?? image.Title ?? "").Replace('\n', ' ').Replace("[", "(").Replace("]", ")");
			return $"![{alt}]({image.FileName})";
		}

		/// <inheritdoc/>
		protected override string FormatImageDescription(XlsxImage image)
		{
			var text = $"**{image.FileName}:** " + image.Description.Trim();
			return String.Join("\n", text.Split('\n').Select(line => "> " + line.TrimEnd('\r')));
		}

		/// <inheritdoc/>
		protected override string FormatShape(string text, string cell)
		{
			return "**Text box" + (cell != null ? " (at " + cell + ")" : "") + ":** " + EscapeLineStarts(Escape(text, InlineMode.Block)).Replace("\n", "  \n");
		}

		/// <inheritdoc/>
		protected override string FormatComments(List<SheetComment> comments)
		{
			var sb = new StringBuilder("**Comments:**");

			foreach (var comment in comments)
			{
				sb.Append("\n- **").Append(comment.Cell).Append("**");
				if (comment.Author != null)
					sb.Append(" (").Append(Escape(comment.Author, InlineMode.Flat)).Append(')');
				sb.Append(": ").Append(Escape(Collapse(comment.Text), InlineMode.Flat));

				foreach (var reply in comment.Replies)
				{
					sb.Append("\n  - ");
					if (reply.Author != null)
						sb.Append(Escape(reply.Author, InlineMode.Flat)).Append(": ");
					sb.Append(Escape(Collapse(reply.Text), InlineMode.Flat));
				}
			}

			return sb.ToString();
		}

		#endregion

		#region Escaping

		/// <summary>
		/// Where escaped text is written.
		/// </summary>
		private enum InlineMode
		{
			/// <summary>A paragraph: line breaks are kept.</summary>
			Block,

			/// <summary>A single line: line breaks become spaces.</summary>
			Flat,

			/// <summary>A table cell: line breaks become <c>&lt;br&gt;</c> and pipes are escaped.</summary>
			Cell,
		}

		/// <summary>
		/// Appends the formula of a cell when <see cref="XlsxExtractorOptions.IncludeFormulas"/> is set.
		/// </summary>
		/// <param name="text">The formatted value.</param>
		/// <param name="cell">The cell.</param>
		/// <param name="mode">Where the text is written.</param>
		/// <returns>The text with the formula.</returns>
		private string AppendFormula(string text, Cell cell, InlineMode mode)
		{
			if (!Options.IncludeFormulas || cell.Formula == null)
				return text;

			return text + " (=" + Escape(Collapse(cell.Formula), mode == InlineMode.Cell ? InlineMode.Cell : InlineMode.Flat) + ")";
		}

		/// <summary>
		/// Wraps text in a link when the cell has a hyperlink and links are included.
		/// </summary>
		/// <param name="text">The escaped text.</param>
		/// <param name="url">The target. Can be <c>null</c>.</param>
		/// <returns>The text or the link.</returns>
		private string Link(string text, string url)
		{
			if (!Options.IncludeHyperlinks || String.IsNullOrWhiteSpace(url) || text.Length == 0)
				return text;

			var target = url.Trim().Replace(" ", "%20").Replace("(", "%28").Replace(")", "%29").Replace("|", "%7C");
			return $"[{text.Replace("[", "\\[").Replace("]", "\\]")}]({target})";
		}

		/// <summary>
		/// Escapes text for a pipe table cell, collapsing line breaks.
		/// </summary>
		/// <param name="text">The text. Can be <c>null</c>.</param>
		/// <returns>The escaped text.</returns>
		private static string TableCell(string text)
		{
			return Escape(Collapse(text ?? ""), InlineMode.Cell);
		}

		/// <summary>
		/// Formats a description written inline, in parentheses.
		/// </summary>
		/// <param name="description">The description.</param>
		/// <returns>The inline description.</returns>
		private static string Inline(string description)
		{
			return "(" + Collapse(description) + ")";
		}

		/// <summary>
		/// Collapses line breaks into spaces.
		/// </summary>
		/// <param name="text">The text.</param>
		/// <returns>The text on a single line.</returns>
		private static string Collapse(string text)
		{
			return Regex.Replace(text, @"\s*\n\s*", " ").Trim();
		}

		/// <summary>
		/// Escapes the characters that Markdown would interpret.
		/// </summary>
		/// <param name="text">The text.</param>
		/// <param name="mode">Where the text is written.</param>
		/// <returns>The escaped text.</returns>
		private static string Escape(string text, InlineMode mode)
		{
			if (String.IsNullOrEmpty(text))
				return text ?? "";

			var sb = new StringBuilder(text.Length + 8);

			for (var i = 0; i < text.Length; i++)
			{
				var c = text[i];
				var previous = i > 0 ? text[i - 1] : ' ';
				var next = i + 1 < text.Length ? text[i + 1] : ' ';

				switch (c)
				{
					case '\\':
					case '*':
					case '`':
						sb.Append('\\').Append(c);
						break;

					case '_':
						// An underscore inside a word (snake_case) cannot start emphasis.
						if (Char.IsLetterOrDigit(previous) && Char.IsLetterOrDigit(next))
							sb.Append(c);
						else
							sb.Append("\\_");
						break;

					case '~':
						sb.Append(next == '~' || previous == '~' ? "\\~" : "~");
						break;

					case '<':
						sb.Append(Char.IsLetter(next) || next == '/' || next == '!' || next == '?' ? "&lt;" : "<");
						break;

					case '|':
						sb.Append(mode == InlineMode.Cell ? "\\|" : "|");
						break;

					case '\r':
						break;

					case '\n':
						sb.Append(mode == InlineMode.Block ? "\n" : mode == InlineMode.Cell ? "<br>" : " ");
						break;

					default:
						sb.Append(c);
						break;
				}
			}

			return sb.ToString();
		}

		/// <summary>
		/// Escapes the start of every line that Markdown would read as a heading, a quote or a list item.
		/// </summary>
		/// <param name="text">The text.</param>
		/// <returns>The escaped text.</returns>
		private static string EscapeLineStarts(string text)
		{
			var lines = text.Split('\n');
			for (var i = 0; i < lines.Length; i++)
			{
				lines[i] = BlockMarker.Replace(lines[i], @"$1\$2", 1);
				lines[i] = OrderedMarker.Replace(lines[i], @"$1\$2", 1);
			}

			return String.Join("\n", lines);
		}

		#endregion
	}
}
