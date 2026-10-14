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
	/// Text format produced by the <see cref="XlsxExtractor"/>.
	/// </summary>
	public enum OutputFormat
	{
		/// <summary>
		/// GitHub flavored Markdown: a heading for every sheet, pipe tables for the cell ranges,
		/// links, images and chart data tables. The best format for language models, which read
		/// Markdown natively.
		/// </summary>
		Markdown,

		/// <summary>
		/// Plain text without markup. Tables are written as rows of pipe-separated cells.
		/// </summary>
		PlainText,
	}
}
