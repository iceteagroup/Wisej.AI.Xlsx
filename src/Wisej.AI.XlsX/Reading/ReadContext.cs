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
using Wisej.AI.XlsX.Packaging;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// The state shared by the readers of the parts of a workbook.
	/// </summary>
	internal sealed class ReadContext
	{
		/// <summary>Gets or sets the package.</summary>
		public XlsxPackage Package { get; set; }

		/// <summary>Gets or sets the extraction options.</summary>
		public XlsxExtractorOptions Options { get; set; }

		/// <summary>Gets or sets the number formatter, configured for the date system of the workbook.</summary>
		public NumberFormatter Formatter { get; set; }

		/// <summary>Gets or sets the number formats of the cell styles.</summary>
		public CellStyles Styles { get; set; }

		/// <summary>Gets or sets the shared string table.</summary>
		public SharedStrings Strings { get; set; }

		/// <summary>Gets or sets the registry that creates the image placements.</summary>
		public ImageRegistry Images { get; set; }

		/// <summary>Gets or sets the pictures placed in cells.</summary>
		public RichValueImages CellImages { get; set; }

		/// <summary>Gets or sets the authors of the threaded comments, by identifier.</summary>
		public Dictionary<string, string> Persons { get; set; } = new Dictionary<string, string>();
	}
}
