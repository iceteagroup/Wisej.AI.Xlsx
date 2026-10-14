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

namespace Wisej.AI.XlsX
{
	/// <summary>
	/// Options that control what the <see cref="XlsxExtractor"/> extracts and how it is formatted.
	/// </summary>
	/// <remarks>
	/// The defaults produce Markdown tuned for language model context and retrieval: every visible
	/// value is written as Excel displays it (with its number format), each block of cells becomes a
	/// table, and charts, images, comments and hyperlinks are kept, while hidden sheets, rows and
	/// columns and decorative images are skipped.
	/// </remarks>
	public sealed class XlsxExtractorOptions
	{
		#region Output

		/// <summary>
		/// Gets or sets the format of the extracted text. The default is <see cref="OutputFormat.Markdown"/>.
		/// </summary>
		public OutputFormat Format { get; set; } = OutputFormat.Markdown;

		/// <summary>
		/// Gets or sets the composite format of the line that introduces every page in
		/// <see cref="XlsxExtractionResult.Text"/>, where <c>{0}</c> is the page number, <c>{1}</c>
		/// the total number of pages and <c>{2}</c> the name of the sheet.
		/// </summary>
		/// <remarks>
		/// When <c>null</c> (the default), Markdown output uses <c>&lt;!-- Page {0} --&gt;</c> and plain text
		/// output uses <c>--- Page {0} ---</c>. Set it to an empty string to concatenate the pages
		/// without separators.
		/// </remarks>
		public string PageHeaderFormat { get; set; }

		/// <summary>
		/// Gets or sets the culture used to format numbers and dates: decimal and group separators,
		/// month and day names. The default is <see cref="CultureInfo.InvariantCulture"/>.
		/// </summary>
		/// <remarks>
		/// The built-in short date and date time formats, which Excel displays in the format of the
		/// regional settings of the machine, are written as ISO 8601 dates (<c>2026-10-14</c>) so that
		/// they are never ambiguous.
		/// </remarks>
		public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

		/// <summary>
		/// Gets or sets whether hyperlinks are written as Markdown links. The default is <c>true</c>.
		/// </summary>
		public bool IncludeHyperlinks { get; set; } = true;

		/// <summary>
		/// Gets or sets whether the formula of a cell is written after its value, as in
		/// <c>1,250 (=SUM(B2:B9))</c>. The default is <c>false</c>.
		/// </summary>
		/// <remarks>
		/// A formula without a cached value (a workbook saved by a program that does not calculate)
		/// is always written, since it is the only content of the cell.
		/// </remarks>
		public bool IncludeFormulas { get; set; }

		/// <summary>
		/// Gets or sets whether the address of every range is written with its table, as in
		/// <c>Range A1:D20</c>, so that the values can be traced back to their cells. The default is <c>true</c>.
		/// </summary>
		public bool IncludeCellRanges { get; set; } = true;

		/// <summary>
		/// Gets or sets whether every table starts with a column containing the row numbers of the sheet.
		/// The default is <c>false</c>.
		/// </summary>
		public bool IncludeRowNumbers { get; set; }

		#endregion

		#region Pagination

		/// <summary>
		/// Gets or sets the maximum number of table rows on a page. A sheet with more rows continues on
		/// the next page, where the table header is repeated. The default is <c>0</c>, which writes every
		/// sheet on a single page.
		/// </summary>
		/// <remarks>
		/// Splitting large sheets produces chunks of a predictable size for retrieval pipelines.
		/// </remarks>
		public int RowsPerPage { get; set; }

		/// <summary>
		/// Gets or sets the maximum number of table rows written for each sheet. The rows that exceed
		/// the limit are replaced by a line that tells how many were omitted. The default is <c>0</c>, which writes every row.
		/// </summary>
		public int MaxRowsPerSheet { get; set; }

		#endregion

		#region Content

		/// <summary>
		/// Gets or sets whether hidden and very hidden sheets are included. The default is <c>false</c>.
		/// </summary>
		public bool IncludeHiddenSheets { get; set; }

		/// <summary>
		/// Gets or sets whether hidden rows, including the rows hidden by a filter, are included. The default is <c>false</c>.
		/// </summary>
		public bool IncludeHiddenRows { get; set; }

		/// <summary>
		/// Gets or sets whether hidden columns are included. The default is <c>false</c>.
		/// </summary>
		public bool IncludeHiddenColumns { get; set; }

		/// <summary>
		/// Gets or sets whether comments and notes are included at the end of their sheet. The default is <c>true</c>.
		/// </summary>
		public bool IncludeComments { get; set; } = true;

		/// <summary>
		/// Gets or sets whether the data of charts is included as a table. The default is <c>true</c>.
		/// </summary>
		public bool IncludeCharts { get; set; } = true;

		/// <summary>
		/// Gets or sets whether the text of shapes and text boxes is included. The default is <c>true</c>.
		/// </summary>
		public bool IncludeTextBoxes { get; set; } = true;

		#endregion

		#region Tables

		/// <summary>
		/// Gets or sets how the header of a range that is not an Excel table is chosen.
		/// The default is <see cref="TableHeaderMode.Auto"/>.
		/// </summary>
		public TableHeaderMode TableHeaders { get; set; } = TableHeaderMode.Auto;

		/// <summary>
		/// Gets or sets how the cells covered by a merged range are written. The default is <see cref="MergedCellMode.Repeat"/>.
		/// </summary>
		public MergedCellMode MergedCells { get; set; } = MergedCellMode.Repeat;

		#endregion

		#region Images

		/// <summary>
		/// Gets or sets whether images are referenced in the text and listed in
		/// <see cref="XlsxExtractionResult.Images"/>. The default is <c>true</c>.
		/// </summary>
		public bool IncludeImages { get; set; } = true;

		/// <summary>
		/// Gets or sets whether the bytes of the images are loaded into <see cref="XlsxImage.Data"/>.
		/// The default is <c>true</c>.
		/// </summary>
		public bool IncludeImageData { get; set; } = true;

		/// <summary>
		/// Gets or sets whether images marked as decorative by the author are skipped. The default is <c>true</c>.
		/// </summary>
		public bool SkipDecorativeImages { get; set; } = true;

		/// <summary>
		/// Gets or sets an optional callback that describes an image, for example with a vision model or OCR.
		/// The returned text is written after the image reference. Return <c>null</c> to add nothing.
		/// </summary>
		/// <remarks>
		/// The callback is invoked once per distinct picture: placements of the same picture reuse the description.
		/// <see cref="XlsxImage.Data"/> is always loaded for the callback, regardless of <see cref="IncludeImageData"/>.
		/// </remarks>
		public Func<XlsxImage, string> ImageDescriber { get; set; }

		#endregion

		/// <summary>
		/// Creates a copy of these options.
		/// </summary>
		/// <returns>A new <see cref="XlsxExtractorOptions"/> with the same values.</returns>
		public XlsxExtractorOptions Clone()
		{
			return (XlsxExtractorOptions)MemberwiseClone();
		}
	}
}
