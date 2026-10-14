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

namespace Wisej.AI.XlsX
{
	/// <summary>
	/// How the header of a range of cells that is not an Excel table is chosen.
	/// </summary>
	/// <remarks>
	/// Excel tables (Insert, Table) always use their own header row.
	/// </remarks>
	public enum TableHeaderMode
	{
		/// <summary>
		/// The first row is the header when it looks like one: it contains only text and labels
		/// at least half of the columns. Otherwise the column letters are used.
		/// </summary>
		Auto,

		/// <summary>
		/// The first row of the range is always the header.
		/// </summary>
		FirstRow,

		/// <summary>
		/// The header contains the column letters (<c>A</c>, <c>B</c>, <c>C</c>, ...) and every row of the range is data.
		/// </summary>
		ColumnLetters,
	}
}
