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
using Wisej.AI.XlsX.Layout;
using Wisej.AI.XlsX.Model;

namespace Wisej.AI.XlsX.Rendering
{
	/// <summary>
	/// Writes the workbook model as text, one page per sheet (or per part of a sheet split by
	/// <see cref="XlsxExtractorOptions.RowsPerPage"/>). Subclasses define the markup.
	/// </summary>
	internal abstract class WorkbookRenderer
	{
		private readonly WorkbookModel _model;
		private readonly List<PageBuffer> _pages = new List<PageBuffer>();
		private PageBuffer _page;

		/// <summary>
		/// Initializes a new instance of the <see cref="WorkbookRenderer"/> class.
		/// </summary>
		/// <param name="model">The workbook model.</param>
		/// <param name="options">The extraction options.</param>
		protected WorkbookRenderer(WorkbookModel model, XlsxExtractorOptions options)
		{
			_model = model;
			Options = options;
		}

		/// <summary>Gets the extraction options.</summary>
		protected XlsxExtractorOptions Options { get; }

		/// <summary>Gets the default format of the line that introduces every page.</summary>
		protected abstract string DefaultPageHeaderFormat { get; }

		#region Render

		/// <summary>
		/// Renders the workbook.
		/// </summary>
		/// <returns>The extraction result.</returns>
		public XlsxExtractionResult Render()
		{
			foreach (var sheet in _model.Sheets)
				RenderSheet(sheet);

			var pages = _pages
				.Select(p => new XlsxPage(p.Number, p.Sheet.Name, p.Sheet.Index, String.Join("\n\n", p.Blocks).Trim(), p.Images))
				.ToList();

			var format = Options.PageHeaderFormat ?? DefaultPageHeaderFormat;
			var text = new StringBuilder();

			foreach (var page in pages)
			{
				if (text.Length > 0)
					text.Append("\n\n");

				if (format.Length > 0)
				{
					text.Append(String.Format(CultureInfo.InvariantCulture, format, page.Number, pages.Count, page.SheetName));
					if (page.Text.Length > 0)
						text.Append("\n\n");
				}

				text.Append(page.Text);
			}

			return new XlsxExtractionResult(text.ToString(), pages, _model.Images, _model.Metadata);
		}

		/// <summary>
		/// Renders a sheet.
		/// </summary>
		/// <param name="sheet">The sheet.</param>
		private void RenderSheet(Sheet sheet)
		{
			StartPage(sheet, false);

			var state = new SheetState();

			foreach (var block in SheetLayout.Build(sheet, Options))
			{
				switch (block)
				{
					case TextBlock text:
						foreach (var cell in text.Lines)
							AddBlock(FormatParagraph(cell));
						break;

					case TableBlock table:
						RenderTable(sheet, table, state);
						break;
				}
			}

			if (state.Omitted > 0)
				AddBlock(FormatOmittedRows(state.Omitted));

			foreach (var item in sheet.Drawings.OrderBy(d => d.Row).ThenBy(d => d.Column))
			{
				switch (item)
				{
					case ChartItem chart:
						AddBlock(FormatChart(chart.Chart, chart.Cell));
						break;

					case PictureItem picture:
						RegisterImage(picture.Image);
						AddBlock(FormatImage(picture.Image));
						if (picture.Image.Description != null)
							AddBlock(FormatImageDescription(picture.Image));
						break;

					case ShapeItem shape:
						AddBlock(FormatShape(shape.Text, shape.Cell));
						break;
				}
			}

			if (sheet.Comments.Count > 0)
				AddBlock(FormatComments(sheet.Comments));
		}

		/// <summary>
		/// Renders a table, splitting it across pages when <see cref="XlsxExtractorOptions.RowsPerPage"/> is set.
		/// </summary>
		/// <param name="sheet">The sheet.</param>
		/// <param name="table">The table.</param>
		/// <param name="state">The rows written and omitted in the sheet.</param>
		private void RenderTable(Sheet sheet, TableBlock table, SheetState state)
		{
			var maxRows = Options.MaxRowsPerSheet;
			var rowsPerPage = Options.RowsPerPage;

			if (maxRows > 0 && state.Written >= maxRows && table.Rows.Count > 0)
			{
				state.Omitted += table.Rows.Count;
				return;
			}

			var header = new List<string>();
			if (Options.IncludeRowNumbers)
				header.Add("Row");
			header.AddRange(table.Header);

			List<string> lines = null;

			void Begin(bool continued)
			{
				if (rowsPerPage > 0 && _page.Rows >= rowsPerPage)
					StartPage(sheet, true);

				var caption = FormatTableCaption(table, continued);
				if (caption != null)
					AddBlock(caption);

				lines = new List<string>(FormatTableHeader(header));
			}

			void Flush()
			{
				AddBlock(String.Join("\n", lines));
			}

			Begin(false);

			foreach (var row in table.Rows)
			{
				if (maxRows > 0 && state.Written >= maxRows)
				{
					state.Omitted++;
					continue;
				}

				if (rowsPerPage > 0 && _page.Rows >= rowsPerPage)
				{
					Flush();
					StartPage(sheet, true);
					Begin(true);
				}

				var cells = new List<string>();
				if (Options.IncludeRowNumbers)
					cells.Add(row.Number.ToString(CultureInfo.InvariantCulture));

				foreach (var cell in row.Cells)
				{
					if (cell?.Image != null)
						RegisterImage(cell.Image);

					cells.Add(cell == null ? "" : FormatCell(cell));
				}

				lines.Add(FormatTableRow(cells));
				_page.Rows++;
				state.Written++;
			}

			Flush();
		}

		#endregion

		#region Pages

		/// <summary>
		/// Starts a new page for a sheet, with the heading of the sheet.
		/// </summary>
		/// <param name="sheet">The sheet.</param>
		/// <param name="continued">Whether the page continues the sheet of the previous page.</param>
		private void StartPage(Sheet sheet, bool continued)
		{
			_page = new PageBuffer { Number = _pages.Count + 1, Sheet = sheet };
			_pages.Add(_page);

			AddBlock(FormatSheetHeading(sheet, continued));
		}

		/// <summary>
		/// Adds a block of text to the current page.
		/// </summary>
		/// <param name="text">The text. Empty text is ignored.</param>
		private void AddBlock(string text)
		{
			if (!String.IsNullOrWhiteSpace(text))
				_page.Blocks.Add(text.TrimEnd());
		}

		/// <summary>
		/// Records that an image is displayed on the current page.
		/// </summary>
		/// <param name="image">The image.</param>
		protected void RegisterImage(XlsxImage image)
		{
			if (image.PageNumber == 0)
				image.PageNumber = _page.Number;

			if (!_page.Images.Contains(image))
				_page.Images.Add(image);
		}

		#endregion

		#region Formatting

		/// <summary>Formats the heading of a sheet.</summary>
		/// <param name="sheet">The sheet.</param>
		/// <param name="continued">Whether the page continues the sheet of the previous page.</param>
		/// <returns>The heading.</returns>
		protected abstract string FormatSheetHeading(Sheet sheet, bool continued);

		/// <summary>Formats a cell written as a paragraph.</summary>
		/// <param name="cell">The cell.</param>
		/// <returns>The paragraph.</returns>
		protected abstract string FormatParagraph(Cell cell);

		/// <summary>Formats the caption of a table, or returns <c>null</c> for none.</summary>
		/// <param name="table">The table.</param>
		/// <param name="continued">Whether the table continues from the previous page.</param>
		/// <returns>The caption or <c>null</c>.</returns>
		protected abstract string FormatTableCaption(TableBlock table, bool continued);

		/// <summary>Formats the header of a table.</summary>
		/// <param name="labels">The labels of the columns.</param>
		/// <returns>The lines of the header.</returns>
		protected abstract IEnumerable<string> FormatTableHeader(List<string> labels);

		/// <summary>Formats a row of a table.</summary>
		/// <param name="cells">The formatted cells.</param>
		/// <returns>The line of the row.</returns>
		protected abstract string FormatTableRow(List<string> cells);

		/// <summary>Formats a cell of a table.</summary>
		/// <param name="cell">The cell.</param>
		/// <returns>The formatted cell.</returns>
		protected abstract string FormatCell(Cell cell);

		/// <summary>Formats the line that replaces the rows beyond <see cref="XlsxExtractorOptions.MaxRowsPerSheet"/>.</summary>
		/// <param name="count">The number of omitted rows.</param>
		/// <returns>The line.</returns>
		protected abstract string FormatOmittedRows(int count);

		/// <summary>Formats a chart as a table of its data.</summary>
		/// <param name="chart">The chart.</param>
		/// <param name="cell">The cell under the top left corner of the chart, or <c>null</c>.</param>
		/// <returns>The chart text.</returns>
		protected abstract string FormatChart(ChartData chart, string cell);

		/// <summary>Formats a picture floating over the sheet.</summary>
		/// <param name="image">The image.</param>
		/// <returns>The image reference.</returns>
		protected abstract string FormatImage(XlsxImage image);

		/// <summary>Formats the description of a picture produced by the image describer.</summary>
		/// <param name="image">The image.</param>
		/// <returns>The description block.</returns>
		protected abstract string FormatImageDescription(XlsxImage image);

		/// <summary>Formats the text of a shape or text box.</summary>
		/// <param name="text">The text.</param>
		/// <param name="cell">The cell under the top left corner of the shape, or <c>null</c>.</param>
		/// <returns>The formatted text.</returns>
		protected abstract string FormatShape(string text, string cell);

		/// <summary>Formats the comments of a sheet.</summary>
		/// <param name="comments">The comments, in cell order.</param>
		/// <returns>The comments block.</returns>
		protected abstract string FormatComments(List<SheetComment> comments);

		/// <summary>
		/// Returns the text of a chart, with its title, type and anchor, shared by the renderers.
		/// </summary>
		/// <param name="chart">The chart.</param>
		/// <param name="cell">The anchor cell, or <c>null</c>.</param>
		/// <returns>The details, such as <c>column chart; at F2</c>, or an empty string.</returns>
		protected static string ChartDetails(ChartData chart, string cell)
		{
			var details = new List<string>();
			if (chart.ChartTypes.Count > 0)
				details.Add(String.Join(" + ", chart.ChartTypes) + " chart");
			if (chart.CategoryAxisTitle != null)
				details.Add("x axis: " + chart.CategoryAxisTitle);
			if (chart.ValueAxisTitle != null)
				details.Add("y axis: " + chart.ValueAxisTitle);
			if (cell != null)
				details.Add("at " + cell);

			return String.Join("; ", details);
		}

		#endregion

		#region Types

		/// <summary>
		/// The text and images of a page being written.
		/// </summary>
		private sealed class PageBuffer
		{
			public int Number;
			public Sheet Sheet;
			public int Rows;
			public readonly List<string> Blocks = new List<string>();
			public readonly List<XlsxImage> Images = new List<XlsxImage>();
		}

		/// <summary>
		/// The rows written and omitted in the current sheet.
		/// </summary>
		private sealed class SheetState
		{
			public int Written;
			public int Omitted;
		}

		#endregion
	}
}
