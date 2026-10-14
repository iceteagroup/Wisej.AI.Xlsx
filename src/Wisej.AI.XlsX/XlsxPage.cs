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

namespace Wisej.AI.XlsX
{
	/// <summary>
	/// A page of an extracted workbook: a sheet, or a part of a sheet split by <see cref="XlsxExtractorOptions.RowsPerPage"/>.
	/// </summary>
	public sealed class XlsxPage
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="XlsxPage"/> class.
		/// </summary>
		/// <param name="number">The one based page number.</param>
		/// <param name="sheetName">The name of the sheet.</param>
		/// <param name="sheetIndex">The zero based position of the sheet in the workbook.</param>
		/// <param name="text">The extracted text of the page.</param>
		/// <param name="images">The images displayed on the page.</param>
		internal XlsxPage(int number, string sheetName, int sheetIndex, string text, IReadOnlyList<XlsxImage> images)
		{
			Number = number;
			SheetName = sheetName;
			SheetIndex = sheetIndex;
			Text = text;
			Images = images;
		}

		/// <summary>
		/// Gets the one based page number.
		/// </summary>
		public int Number { get; }

		/// <summary>
		/// Gets the name of the sheet the page belongs to.
		/// </summary>
		public string SheetName { get; }

		/// <summary>
		/// Gets the zero based position of the sheet in the workbook, counting the hidden sheets.
		/// </summary>
		public int SheetIndex { get; }

		/// <summary>
		/// Gets the extracted text of the page, without the page header line.
		/// </summary>
		public string Text { get; }

		/// <summary>
		/// Gets the images displayed on the page.
		/// </summary>
		public IReadOnlyList<XlsxImage> Images { get; }

		/// <summary>
		/// Returns the page number, the sheet name and the length of the text.
		/// </summary>
		/// <returns>A string describing the page.</returns>
		public override string ToString()
		{
			return $"Page {Number}: {SheetName} ({Text.Length} characters)";
		}
	}
}
