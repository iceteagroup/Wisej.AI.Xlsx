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
using System.Linq;
using Wisej.AI.XlsX.Model;

namespace Wisej.AI.XlsX.Layout
{
	/// <summary>
	/// Splits the cells of a sheet into the blocks that are written: Excel tables, ranges of cells
	/// written as tables, and runs of single cells (titles, notes) written as paragraphs.
	/// </summary>
	/// <remarks>
	/// <para>
	/// A sheet is a grid, but most sheets contain several separate blocks: a title, one or more
	/// tables, notes. Writing the whole used range as one table would mix them and fill the output
	/// with empty cells. The cells that are not part of an Excel table are therefore split
	/// recursively at the empty rows and columns that separate them (an X-Y cut), and empty
	/// columns are removed from every block.
	/// </para>
	/// <para>
	/// Two column groups that span the same rows are kept together: they are one table with an
	/// empty spacer column, not two tables.
	/// </para>
	/// </remarks>
	internal static class SheetLayout
	{
		/// <summary>The largest merged range whose value is repeated in every covered cell.</summary>
		private const int MaxRepeatedMerge = 10000;

		/// <summary>The maximum recursion depth of the X-Y cut.</summary>
		private const int MaxCutDepth = 24;

		/// <summary>
		/// Splits a sheet into blocks.
		/// </summary>
		/// <param name="sheet">The sheet.</param>
		/// <param name="options">The extraction options.</param>
		/// <returns>The blocks, in reading order.</returns>
		public static List<SheetBlock> Build(Sheet sheet, XlsxExtractorOptions options)
		{
			var grid = BuildGrid(sheet, options);
			var blocks = new List<SheetBlock>();

			foreach (var table in sheet.Tables.OrderBy(t => t.Range.FirstRow).ThenBy(t => t.Range.FirstColumn))
			{
				var block = BuildTable(sheet, table, grid);
				if (block != null)
					blocks.Add(block);
			}

			var positions = grid.Select(pair => new Position(pair.Key, pair.Value)).ToList();
			foreach (var region in Cut(positions, 0))
				blocks.AddRange(BuildRegion(region, options));

			blocks = blocks
				.OrderBy(b => b.Range.FirstRow)
				.ThenBy(b => b.Range.FirstColumn)
				.ToList();

			ApplyColumnLabels(blocks);
			return blocks;
		}

		/// <summary>
		/// Uses a row of column labels as the header of the blocks below it.
		/// </summary>
		/// <param name="blocks">The blocks, in reading order. Label rows that are used are removed.</param>
		/// <remarks>
		/// Forms and price lists often have one row of labels above several blocks separated by empty
		/// rows. Each block alone has no header, so the labels are applied to every block below that
		/// shares at least two of their columns, until the next row of labels.
		/// </remarks>
		private static void ApplyColumnLabels(List<SheetBlock> blocks)
		{
			var used = new HashSet<TableBlock>();
			var tables = blocks.OfType<TableBlock>().Where(t => t.Name == null).ToList();
			Dictionary<int, string> labels = null;
			TableBlock source = null;

			Dictionary<int, string> Labels(TableBlock table)
			{
				var result = new Dictionary<int, string>();
				for (var i = 0; i < table.Columns.Count; i++)
				{
					var text = table.Rows[0].Cells[i]?.Text;
					if (!String.IsNullOrEmpty(text))
						result[table.Columns[i]] = text;
				}

				return result;
			}

			bool Accepts(TableBlock table, Dictionary<int, string> candidate)
			{
				return (table.HasLetterHeader || table.GuessedHeaderRow != null) && table.Columns.Count(candidate.ContainsKey) >= 2;
			}

			for (var index = 0; index < tables.Count; index++)
			{
				var table = tables[index];

				// A row of labels is used only when a block below, before the next row of labels, shares its columns;
				// otherwise it is a row of data and the current labels stay in effect.
				if (IsLabelRow(table))
				{
					var candidate = Labels(table);
					var applies = tables.Skip(index + 1).TakeWhile(t => !IsLabelRow(t)).Any(t => Accepts(t, candidate));
					if (applies)
					{
						source = table;
						labels = candidate;
						continue;
					}
				}

				if (labels == null || !Accepts(table, labels))
					continue;

				// A first row that was only guessed to be the header is data under the labels.
				if (table.GuessedHeaderRow != null)
				{
					table.Rows.Insert(0, table.GuessedHeaderRow);
					table.GuessedHeaderRow = null;
				}

				for (var i = 0; i < table.Columns.Count; i++)
					table.Header[i] = labels.TryGetValue(table.Columns[i], out var label) ? label : CellRange.ColumnName(table.Columns[i]);

				table.HasLetterHeader = false;
				used.Add(source);
			}

			blocks.RemoveAll(b => b is TableBlock t && used.Contains(t));
		}

		/// <summary>
		/// Determines whether a block is a single row of column labels without a header of its own.
		/// </summary>
		/// <param name="table">The block.</param>
		/// <returns><c>true</c> for a row of labels.</returns>
		private static bool IsLabelRow(TableBlock table)
		{
			if (!table.HasLetterHeader || table.Rows.Count != 1)
				return false;

			var cells = table.Rows[0].Cells.Where(c => c != null).ToList();
			return cells.Count >= 2 && cells.All(c => c.Kind == CellKind.Text && c.Image == null);
		}

		#region Grid

		/// <summary>
		/// Returns the cells to write, by position, with merged values repeated when required.
		/// </summary>
		/// <param name="sheet">The sheet.</param>
		/// <param name="options">The extraction options.</param>
		/// <returns>The cells by <see cref="Sheet.Key"/>.</returns>
		private static Dictionary<long, Cell> BuildGrid(Sheet sheet, XlsxExtractorOptions options)
		{
			var grid = new Dictionary<long, Cell>(sheet.Cells);

			if (options.MergedCells != MergedCellMode.Repeat)
				return grid;

			foreach (var range in sheet.MergedRanges)
			{
				if (range.CellCount > MaxRepeatedMerge)
					continue;

				var origin = sheet.GetCell(range.FirstRow, range.FirstColumn);
				if (origin == null)
					continue;

				for (var row = range.FirstRow; row <= range.LastRow; row++)
				{
					for (var column = range.FirstColumn; column <= range.LastColumn; column++)
					{
						if (!sheet.IsHidden(row, column))
							grid[Sheet.Key(row, column)] = origin;
					}
				}
			}

			return grid;
		}

		#endregion

		#region Excel Tables

		/// <summary>
		/// Builds the block of an Excel table and removes its cells from the grid.
		/// </summary>
		/// <param name="sheet">The sheet.</param>
		/// <param name="table">The table.</param>
		/// <param name="grid">The cells not yet assigned to a block.</param>
		/// <returns>The block, or <c>null</c> when the table is empty.</returns>
		private static TableBlock BuildTable(Sheet sheet, SheetTable table, Dictionary<long, Cell> grid)
		{
			var range = table.Range;
			var cells = grid.Where(pair => range.Contains(Row(pair.Key), Column(pair.Key))).ToList();
			if (cells.Count == 0 && table.HeaderRowCount == 0)
				return null;

			foreach (var pair in cells)
				grid.Remove(pair.Key);

			var lookup = cells.ToDictionary(p => p.Key, p => p.Value);
			var columns = Enumerable.Range(range.FirstColumn, range.ColumnCount).Where(c => !sheet.HiddenColumns.Contains(c)).ToList();

			var block = new TableBlock { Name = table.Name, Range = range, Columns = columns };

			for (var i = 0; i < columns.Count; i++)
			{
				var column = columns[i];
				string label = null;

				if (table.HeaderRowCount > 0 && lookup.TryGetValue(Sheet.Key(range.FirstRow, column), out var headerCell))
					label = headerCell.Text;

				var index = column - range.FirstColumn;
				if (String.IsNullOrEmpty(label) && index < table.ColumnNames.Count)
					label = table.ColumnNames[index];

				block.Header.Add(String.IsNullOrEmpty(label) ? CellRange.ColumnName(column) : label);
			}

			var firstDataRow = range.FirstRow + table.HeaderRowCount;
			var rows = cells.Select(p => Row(p.Key)).Where(r => r >= firstDataRow).Distinct().OrderBy(r => r);

			foreach (var row in rows)
				block.Rows.Add(BuildRow(row, columns, lookup));

			return block;
		}

		#endregion

		#region Regions

		/// <summary>
		/// Splits a set of cells at the empty rows and columns that separate them.
		/// </summary>
		/// <param name="positions">The cells.</param>
		/// <param name="depth">The recursion depth.</param>
		/// <returns>The regions.</returns>
		private static IEnumerable<List<Position>> Cut(List<Position> positions, int depth)
		{
			if (positions.Count == 0)
				yield break;

			if (depth >= MaxCutDepth)
			{
				yield return positions;
				yield break;
			}

			var rowGroups = Split(positions, p => p.Row);
			if (rowGroups.Count > 1)
			{
				foreach (var group in rowGroups)
				{
					foreach (var region in Cut(group, depth + 1))
						yield return region;
				}

				yield break;
			}

			var columnGroups = MergeAligned(Split(positions, p => p.Column));
			if (columnGroups.Count > 1)
			{
				foreach (var group in columnGroups)
				{
					foreach (var region in Cut(group, depth + 1))
						yield return region;
				}

				yield break;
			}

			yield return positions;
		}

		/// <summary>
		/// Splits cells into groups separated by at least one empty row or column.
		/// </summary>
		/// <param name="positions">The cells.</param>
		/// <param name="coordinate">Selects the row or the column of a cell.</param>
		/// <returns>The groups, in order.</returns>
		private static List<List<Position>> Split(List<Position> positions, Func<Position, int> coordinate)
		{
			var sorted = positions.OrderBy(coordinate).ToList();
			var groups = new List<List<Position>> { new List<Position>() };
			var previous = coordinate(sorted[0]);

			foreach (var position in sorted)
			{
				var value = coordinate(position);
				if (value > previous + 1)
					groups.Add(new List<Position>());

				groups[groups.Count - 1].Add(position);
				previous = value;
			}

			return groups;
		}

		/// <summary>
		/// Joins adjacent column groups that span exactly the same rows: they are one table with an empty spacer column.
		/// </summary>
		/// <param name="groups">The column groups, in order.</param>
		/// <returns>The merged groups.</returns>
		private static List<List<Position>> MergeAligned(List<List<Position>> groups)
		{
			var result = new List<List<Position>>();

			foreach (var group in groups)
			{
				var last = result.Count > 0 ? result[result.Count - 1] : null;
				if (last != null
					&& last.Min(p => p.Row) == group.Min(p => p.Row)
					&& last.Max(p => p.Row) == group.Max(p => p.Row))
				{
					last.AddRange(group);
				}
				else
				{
					result.Add(new List<Position>(group));
				}
			}

			return result;
		}

		/// <summary>
		/// Builds the blocks of a region: paragraphs when every row has a single cell, otherwise a table,
		/// preceded by the title rows (a single cell, possibly merged across the table) above it.
		/// </summary>
		/// <param name="region">The cells of the region.</param>
		/// <param name="options">The extraction options.</param>
		/// <returns>The blocks.</returns>
		private static IEnumerable<SheetBlock> BuildRegion(List<Position> region, XlsxExtractorOptions options)
		{
			var byRow = region.GroupBy(p => p.Row).OrderBy(g => g.Key).ToList();

			// A merged title repeats the same cell in every column it covers: it is still a single cell.
			bool IsSingleCell(IGrouping<int, Position> row) => row.Select(p => p.Cell).Distinct().Count() == 1;

			var titleRows = byRow.TakeWhile(IsSingleCell).Count();
			var allSingle = titleRows == byRow.Count;
			var columnCount = region.Select(p => p.Column).Distinct().Count();

			// A column of values is a table, not a list of paragraphs.
			if (allSingle && columnCount == 1 && byRow.Count > 1 && region.Any(p => p.Cell.Kind != CellKind.Text || p.Cell.Image != null))
				titleRows = 0;

			if (titleRows > 0)
			{
				var titles = byRow.Take(titleRows).ToList();
				var text = new TextBlock
				{
					Range = new CellRange(titles[0].Key, titles.SelectMany(g => g).Min(p => p.Column), titles[titles.Count - 1].Key, titles.SelectMany(g => g).Max(p => p.Column)),
				};
				text.Lines.AddRange(titles.Select(g => g.First().Cell));
				yield return text;

				if (allSingle)
					yield break;

				region = byRow.Skip(titleRows).SelectMany(g => g).ToList();
			}

			yield return BuildTableRegion(region, options);
		}

		/// <summary>
		/// Builds the table of a region.
		/// </summary>
		/// <param name="region">The cells of the region.</param>
		/// <param name="options">The extraction options.</param>
		/// <returns>The block.</returns>
		private static TableBlock BuildTableRegion(List<Position> region, XlsxExtractorOptions options)
		{
			var rows = region.Select(p => p.Row).Distinct().OrderBy(r => r).ToList();
			var columns = region.Select(p => p.Column).Distinct().OrderBy(c => c).ToList();
			var range = new CellRange(rows[0], columns[0], rows[rows.Count - 1], columns[columns.Count - 1]);
			var lookup = region.ToDictionary(p => Sheet.Key(p.Row, p.Column), p => p.Cell);

			var block = new TableBlock { Range = range, Columns = columns };
			var firstRow = BuildRow(rows[0], columns, lookup);

			if (UseFirstRowAsHeader(firstRow, rows.Count, options.TableHeaders))
			{
				for (var i = 0; i < columns.Count; i++)
					block.Header.Add(firstRow.Cells[i]?.Text is string label && label.Length > 0 ? label : CellRange.ColumnName(columns[i]));

				if (options.TableHeaders == TableHeaderMode.Auto)
					block.GuessedHeaderRow = firstRow;

				rows.RemoveAt(0);
			}
			else
			{
				block.Header.AddRange(columns.Select(CellRange.ColumnName));
				block.HasLetterHeader = true;
			}

			foreach (var row in rows)
				block.Rows.Add(BuildRow(row, columns, lookup));

			return block;
		}

		/// <summary>
		/// Determines whether the first row of a range is its header.
		/// </summary>
		/// <param name="firstRow">The first row.</param>
		/// <param name="rowCount">The number of rows of the range.</param>
		/// <param name="mode">The header mode.</param>
		/// <returns><c>true</c> when the first row is the header.</returns>
		private static bool UseFirstRowAsHeader(TableRow firstRow, int rowCount, TableHeaderMode mode)
		{
			switch (mode)
			{
				case TableHeaderMode.FirstRow:
					return rowCount > 1;

				case TableHeaderMode.ColumnLetters:
					return false;

				default:
					var filled = firstRow.Cells.Where(c => c != null).ToList();
					return rowCount > 1
						&& filled.Count * 2 >= firstRow.Cells.Count
						&& filled.All(c => c.Kind == CellKind.Text && c.Image == null && c.Text.Length > 0);
			}
		}

		#endregion

		#region Helpers

		/// <summary>
		/// Builds a row of a table.
		/// </summary>
		/// <param name="row">The one based row.</param>
		/// <param name="columns">The columns of the table.</param>
		/// <param name="lookup">The cells by key.</param>
		/// <returns>The row.</returns>
		private static TableRow BuildRow(int row, List<int> columns, Dictionary<long, Cell> lookup)
		{
			var result = new TableRow { Number = row };
			foreach (var column in columns)
				result.Cells.Add(lookup.TryGetValue(Sheet.Key(row, column), out var cell) ? cell : null);

			return result;
		}

		/// <summary>Returns the row of a key built by <see cref="Sheet.Key"/>.</summary>
		private static int Row(long key) => (int)(key >> 16);

		/// <summary>Returns the column of a key built by <see cref="Sheet.Key"/>.</summary>
		private static int Column(long key) => (int)(key & 0xFFFF);

		/// <summary>
		/// A cell and its position in the grid. Merged values share the cell of the top left position.
		/// </summary>
		private readonly struct Position
		{
			public Position(long key, Cell cell)
			{
				Row = SheetLayout.Row(key);
				Column = SheetLayout.Column(key);
				Cell = cell;
			}

			public int Row { get; }

			public int Column { get; }

			public Cell Cell { get; }
		}

		#endregion
	}

	/// <summary>
	/// A block of a sheet.
	/// </summary>
	internal abstract class SheetBlock
	{
		/// <summary>Gets or sets the range covered by the block.</summary>
		public CellRange Range { get; set; }
	}

	/// <summary>
	/// Cells written as paragraphs, one per row.
	/// </summary>
	internal sealed class TextBlock : SheetBlock
	{
		/// <summary>Gets the cells, in row order.</summary>
		public List<Cell> Lines { get; } = new List<Cell>();
	}

	/// <summary>
	/// Cells written as a table.
	/// </summary>
	internal sealed class TableBlock : SheetBlock
	{
		/// <summary>Gets or sets the name of the Excel table, or <c>null</c> for a range of cells.</summary>
		public string Name { get; set; }

		/// <summary>Gets or sets the columns of the sheet included in the table.</summary>
		public List<int> Columns { get; set; } = new List<int>();

		/// <summary>Gets the labels of the columns.</summary>
		public List<string> Header { get; } = new List<string>();

		/// <summary>Gets or sets whether the header contains the column letters instead of labels.</summary>
		public bool HasLetterHeader { get; set; }

		/// <summary>Gets or sets the first row when it was guessed to be the header (<see cref="TableHeaderMode.Auto"/>), or <c>null</c>.</summary>
		public TableRow GuessedHeaderRow { get; set; }

		/// <summary>Gets the data rows.</summary>
		public List<TableRow> Rows { get; } = new List<TableRow>();
	}

	/// <summary>
	/// A row of a table.
	/// </summary>
	internal sealed class TableRow
	{
		/// <summary>Gets or sets the one based row number in the sheet.</summary>
		public int Number { get; set; }

		/// <summary>Gets the cells, aligned with <see cref="TableBlock.Columns"/>; <c>null</c> for empty cells.</summary>
		public List<Cell> Cells { get; } = new List<Cell>();
	}
}
