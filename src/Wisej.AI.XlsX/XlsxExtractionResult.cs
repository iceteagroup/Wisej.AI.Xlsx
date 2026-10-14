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
	/// The result of the extraction of an Excel workbook.
	/// </summary>
	public sealed class XlsxExtractionResult
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="XlsxExtractionResult"/> class.
		/// </summary>
		/// <param name="text">The text of the whole workbook.</param>
		/// <param name="pages">The pages of the workbook.</param>
		/// <param name="images">The images of the workbook.</param>
		/// <param name="metadata">The workbook properties.</param>
		internal XlsxExtractionResult(
			string text,
			IReadOnlyList<XlsxPage> pages,
			IReadOnlyList<XlsxImage> images,
			XlsxMetadata metadata)
		{
			Text = text;
			Pages = pages;
			Images = images;
			Metadata = metadata;
		}

		/// <summary>
		/// Gets the text of the whole workbook, with a header line introducing every page
		/// (see <see cref="XlsxExtractorOptions.PageHeaderFormat"/>).
		/// </summary>
		public string Text { get; }

		/// <summary>
		/// Gets the pages of the workbook: one for every sheet, or more when a sheet is split by
		/// <see cref="XlsxExtractorOptions.RowsPerPage"/>.
		/// </summary>
		public IReadOnlyList<XlsxPage> Pages { get; }

		/// <summary>
		/// Gets the number of pages.
		/// </summary>
		public int PageCount => Pages.Count;

		/// <summary>
		/// Gets every image placed in the workbook, in reading order.
		/// </summary>
		public IReadOnlyList<XlsxImage> Images { get; }

		/// <summary>
		/// Gets the workbook properties.
		/// </summary>
		public XlsxMetadata Metadata { get; }

		/// <summary>
		/// Returns the text of the whole workbook.
		/// </summary>
		/// <returns>The value of <see cref="Text"/>.</returns>
		public override string ToString() => Text;
	}
}
