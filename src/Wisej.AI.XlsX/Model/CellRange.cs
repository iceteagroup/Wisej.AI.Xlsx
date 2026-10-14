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
using System.Globalization;
using System.Text;

namespace Wisej.AI.XlsX.Model
{
	/// <summary>
	/// A rectangular range of cells, with one based row and column numbers.
	/// </summary>
	internal readonly struct CellRange
	{
		/// <summary>The largest row number of a worksheet.</summary>
		public const int MaxRow = 1048576;

		/// <summary>The largest column number of a worksheet (<c>XFD</c>).</summary>
		public const int MaxColumn = 16384;

		/// <summary>
		/// Initializes a new instance of the <see cref="CellRange"/> struct.
		/// </summary>
		/// <param name="firstRow">The first row.</param>
		/// <param name="firstColumn">The first column.</param>
		/// <param name="lastRow">The last row.</param>
		/// <param name="lastColumn">The last column.</param>
		public CellRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
		{
			FirstRow = Math.Min(firstRow, lastRow);
			FirstColumn = Math.Min(firstColumn, lastColumn);
			LastRow = Math.Max(firstRow, lastRow);
			LastColumn = Math.Max(firstColumn, lastColumn);
		}

		/// <summary>Gets the first row.</summary>
		public int FirstRow { get; }

		/// <summary>Gets the first column.</summary>
		public int FirstColumn { get; }

		/// <summary>Gets the last row.</summary>
		public int LastRow { get; }

		/// <summary>Gets the last column.</summary>
		public int LastColumn { get; }

		/// <summary>Gets the number of rows.</summary>
		public int RowCount => LastRow - FirstRow + 1;

		/// <summary>Gets the number of columns.</summary>
		public int ColumnCount => LastColumn - FirstColumn + 1;

		/// <summary>Gets the number of cells, without overflowing.</summary>
		public long CellCount => (long)RowCount * ColumnCount;

		/// <summary>
		/// Determines whether the range contains a cell.
		/// </summary>
		/// <param name="row">The row.</param>
		/// <param name="column">The column.</param>
		/// <returns><c>true</c> when the cell is inside the range.</returns>
		public bool Contains(int row, int column)
		{
			return row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
		}

		/// <summary>
		/// Returns the address of the range, such as <c>A1:D20</c>, or of its only cell.
		/// </summary>
		/// <returns>The address.</returns>
		public override string ToString()
		{
			var first = CellName(FirstRow, FirstColumn);
			return FirstRow == LastRow && FirstColumn == LastColumn ? first : first + ":" + CellName(LastRow, LastColumn);
		}

		#region Parsing

		/// <summary>
		/// Parses a range address such as <c>A1:D20</c>, <c>$B$2</c> or <c>A:C</c>; the first range of a list is used.
		/// </summary>
		/// <param name="text">The address. Can be <c>null</c>.</param>
		/// <param name="range">The parsed range.</param>
		/// <returns><c>true</c> when the address is valid.</returns>
		public static bool TryParse(string text, out CellRange range)
		{
			range = default;
			if (String.IsNullOrWhiteSpace(text))
				return false;

			// A list of ranges ("A1:B2 D4:E5") is used for hyperlinks and merged areas: take the first.
			text = text.Trim().Split(' ')[0];

			var colon = text.IndexOf(':');
			var firstText = colon < 0 ? text : text.Substring(0, colon);
			var lastText = colon < 0 ? text : text.Substring(colon + 1);

			if (!TryParseCell(firstText, out var r1, out var c1) || !TryParseCell(lastText, out var r2, out var c2))
				return false;

			// Whole columns ("A:C") and whole rows ("2:5").
			if (r1 == 0) r1 = 1;
			if (r2 == 0) r2 = MaxRow;
			if (c1 == 0) c1 = 1;
			if (c2 == 0) c2 = MaxColumn;

			range = new CellRange(r1, c1, r2, c2);
			return true;
		}

		/// <summary>
		/// Parses a cell address such as <c>B12</c> or <c>$B$12</c>. A missing row or column part is returned as zero.
		/// </summary>
		/// <param name="text">The address.</param>
		/// <param name="row">The one based row, or zero.</param>
		/// <param name="column">The one based column, or zero.</param>
		/// <returns><c>true</c> when the address is valid.</returns>
		public static bool TryParseCell(string text, out int row, out int column)
		{
			row = 0;
			column = 0;

			if (String.IsNullOrEmpty(text))
				return false;

			var i = 0;
			if (i < text.Length && text[i] == '$')
				i++;

			while (i < text.Length && Char.IsLetter(text[i]))
			{
				column = column * 26 + (Char.ToUpperInvariant(text[i]) - 'A' + 1);
				if (column > MaxColumn)
					return false;
				i++;
			}

			if (i < text.Length && text[i] == '$')
				i++;

			var digits = text.Substring(i);
			if (digits.Length > 0)
			{
				if (!Int32.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out row) || row < 1 || row > MaxRow)
					return false;
			}

			return row > 0 || column > 0;
		}

		/// <summary>
		/// Returns the letters of a column, such as <c>A</c>, <c>Z</c> or <c>AB</c>.
		/// </summary>
		/// <param name="column">The one based column.</param>
		/// <returns>The column name.</returns>
		public static string ColumnName(int column)
		{
			var sb = new StringBuilder();
			while (column > 0)
			{
				var rest = (column - 1) % 26;
				sb.Insert(0, (char)('A' + rest));
				column = (column - 1) / 26;
			}

			return sb.ToString();
		}

		/// <summary>
		/// Returns the address of a cell, such as <c>D5</c>.
		/// </summary>
		/// <param name="row">The one based row.</param>
		/// <param name="column">The one based column.</param>
		/// <returns>The cell address.</returns>
		public static string CellName(int row, int column)
		{
			return ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
		}

		#endregion
	}
}
