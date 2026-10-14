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

using System.Collections.Generic;

namespace Wisej.AI.XlsX.Model
{
	/// <summary>
	/// The format independent model of a workbook, produced by the readers and consumed by the renderers.
	/// </summary>
	internal sealed class WorkbookModel
	{
		/// <summary>Gets the sheets to extract, in workbook order.</summary>
		public List<Sheet> Sheets { get; } = new List<Sheet>();

		/// <summary>Gets every image placement, in reading order.</summary>
		public List<XlsxImage> Images { get; } = new List<XlsxImage>();

		/// <summary>Gets or sets the workbook properties.</summary>
		public XlsxMetadata Metadata { get; set; } = new XlsxMetadata();
	}

	/// <summary>
	/// A worksheet or a chart sheet.
	/// </summary>
	internal sealed class Sheet
	{
		/// <summary>Gets or sets the name of the sheet.</summary>
		public string Name { get; set; }

		/// <summary>Gets or sets the zero based position of the sheet in the workbook.</summary>
		public int Index { get; set; }

		/// <summary>Gets or sets whether the sheet is a chart sheet, which has no cells.</summary>
		public bool IsChartSheet { get; set; }

		/// <summary>Gets the cells that have a value, indexed by <see cref="Key"/>.</summary>
		public Dictionary<long, Cell> Cells { get; } = new Dictionary<long, Cell>();

		/// <summary>Gets the merged ranges.</summary>
		public List<CellRange> MergedRanges { get; } = new List<CellRange>();

		/// <summary>Gets the Excel tables (list objects) of the sheet.</summary>
		public List<SheetTable> Tables { get; } = new List<SheetTable>();

		/// <summary>Gets the charts, pictures and shapes floating over the sheet.</summary>
		public List<DrawingItem> Drawings { get; } = new List<DrawingItem>();

		/// <summary>Gets the comments and notes of the sheet.</summary>
		public List<SheetComment> Comments { get; } = new List<SheetComment>();

		/// <summary>Gets the rows hidden in the sheet, when they are excluded.</summary>
		public HashSet<int> HiddenRows { get; } = new HashSet<int>();

		/// <summary>Gets the columns hidden in the sheet, when they are excluded.</summary>
		public HashSet<int> HiddenColumns { get; } = new HashSet<int>();

		/// <summary>
		/// Returns the key of a cell in <see cref="Cells"/>.
		/// </summary>
		/// <param name="row">The one based row.</param>
		/// <param name="column">The one based column.</param>
		/// <returns>The key.</returns>
		public static long Key(int row, int column) => ((long)row << 16) | (uint)column;

		/// <summary>
		/// Returns the cell at a position, or <c>null</c> when it is empty.
		/// </summary>
		/// <param name="row">The one based row.</param>
		/// <param name="column">The one based column.</param>
		/// <returns>The cell or <c>null</c>.</returns>
		public Cell GetCell(int row, int column)
		{
			return Cells.TryGetValue(Key(row, column), out var cell) ? cell : null;
		}

		/// <summary>
		/// Determines whether a row or column is hidden and excluded from the output.
		/// </summary>
		/// <param name="row">The one based row.</param>
		/// <param name="column">The one based column.</param>
		/// <returns><c>true</c> when the cell is hidden.</returns>
		public bool IsHidden(int row, int column)
		{
			return HiddenRows.Contains(row) || HiddenColumns.Contains(column);
		}
	}

	/// <summary>
	/// The kind of value of a cell.
	/// </summary>
	internal enum CellKind
	{
		/// <summary>A string.</summary>
		Text,

		/// <summary>A number.</summary>
		Number,

		/// <summary>A number displayed as a date or a time.</summary>
		Date,

		/// <summary>A logical value.</summary>
		Boolean,

		/// <summary>An error value such as <c>#DIV/0!</c>.</summary>
		Error,
	}

	/// <summary>
	/// A cell with a value.
	/// </summary>
	internal sealed class Cell
	{
		/// <summary>Gets or sets the one based row.</summary>
		public int Row { get; set; }

		/// <summary>Gets or sets the one based column.</summary>
		public int Column { get; set; }

		/// <summary>Gets or sets the kind of value.</summary>
		public CellKind Kind { get; set; }

		/// <summary>Gets or sets the value as Excel displays it.</summary>
		public string Text { get; set; }

		/// <summary>Gets or sets the formula, without the leading equal sign, or <c>null</c>.</summary>
		public string Formula { get; set; }

		/// <summary>Gets or sets the target of the hyperlink of the cell, or <c>null</c>.</summary>
		public string Hyperlink { get; set; }

		/// <summary>Gets or sets the picture placed in the cell, or <c>null</c>.</summary>
		public XlsxImage Image { get; set; }

		/// <summary>Gets whether the cell has content to write.</summary>
		public bool HasContent => !string.IsNullOrEmpty(Text) || Image != null;
	}

	/// <summary>
	/// An Excel table (list object).
	/// </summary>
	internal sealed class SheetTable
	{
		/// <summary>Gets or sets the display name of the table.</summary>
		public string Name { get; set; }

		/// <summary>Gets or sets the range of the table, including the header and totals rows.</summary>
		public CellRange Range { get; set; }

		/// <summary>Gets or sets the number of header rows (zero or one).</summary>
		public int HeaderRowCount { get; set; } = 1;

		/// <summary>Gets or sets the number of totals rows (zero or one).</summary>
		public int TotalsRowCount { get; set; }

		/// <summary>Gets the names of the columns, used when the header row is not shown.</summary>
		public List<string> ColumnNames { get; } = new List<string>();
	}

	/// <summary>
	/// A comment (threaded, with replies) or a note.
	/// </summary>
	internal sealed class SheetComment
	{
		/// <summary>Gets or sets the address of the cell, such as <c>B4</c>.</summary>
		public string Cell { get; set; }

		/// <summary>Gets or sets the one based row of the cell, used to order the comments.</summary>
		public int Row { get; set; }

		/// <summary>Gets or sets the one based column of the cell, used to order the comments.</summary>
		public int Column { get; set; }

		/// <summary>Gets or sets the author, or <c>null</c>.</summary>
		public string Author { get; set; }

		/// <summary>Gets or sets the text.</summary>
		public string Text { get; set; }

		/// <summary>Gets the replies of a threaded comment.</summary>
		public List<SheetComment> Replies { get; } = new List<SheetComment>();
	}

	/// <summary>
	/// An object floating over a sheet: a chart, a picture or a shape with text.
	/// </summary>
	internal abstract class DrawingItem
	{
		/// <summary>Gets or sets the address of the cell under the top left corner, or <c>null</c>.</summary>
		public string Cell { get; set; }

		/// <summary>Gets or sets the one based row of the anchor, used to order the drawings.</summary>
		public int Row { get; set; }

		/// <summary>Gets or sets the one based column of the anchor, used to order the drawings.</summary>
		public int Column { get; set; }
	}

	/// <summary>
	/// A chart.
	/// </summary>
	internal sealed class ChartItem : DrawingItem
	{
		/// <summary>Gets or sets the data of the chart.</summary>
		public ChartData Chart { get; set; }
	}

	/// <summary>
	/// A picture.
	/// </summary>
	internal sealed class PictureItem : DrawingItem
	{
		/// <summary>Gets or sets the image placement.</summary>
		public XlsxImage Image { get; set; }
	}

	/// <summary>
	/// A shape or text box with text.
	/// </summary>
	internal sealed class ShapeItem : DrawingItem
	{
		/// <summary>Gets or sets the text of the shape, one line per paragraph.</summary>
		public string Text { get; set; }
	}
}
