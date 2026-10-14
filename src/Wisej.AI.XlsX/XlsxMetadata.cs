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

namespace Wisej.AI.XlsX
{
	/// <summary>
	/// The properties of an Excel workbook: core properties, application properties, custom properties and sheet names.
	/// </summary>
	public sealed class XlsxMetadata
	{
		/// <summary>Gets the workbook title.</summary>
		public string Title { get; internal set; }

		/// <summary>Gets the workbook subject.</summary>
		public string Subject { get; internal set; }

		/// <summary>Gets the author of the workbook.</summary>
		public string Author { get; internal set; }

		/// <summary>Gets the keywords of the workbook.</summary>
		public string Keywords { get; internal set; }

		/// <summary>Gets the description (comments) of the workbook.</summary>
		public string Description { get; internal set; }

		/// <summary>Gets the category of the workbook.</summary>
		public string Category { get; internal set; }

		/// <summary>Gets the user who last modified the workbook.</summary>
		public string LastModifiedBy { get; internal set; }

		/// <summary>Gets the revision number of the workbook.</summary>
		public string Revision { get; internal set; }

		/// <summary>Gets the language of the workbook.</summary>
		public string Language { get; internal set; }

		/// <summary>Gets the creation date of the workbook.</summary>
		public DateTime? Created { get; internal set; }

		/// <summary>Gets the date the workbook was last modified.</summary>
		public DateTime? Modified { get; internal set; }

		/// <summary>Gets the company recorded in the application properties.</summary>
		public string Company { get; internal set; }

		/// <summary>Gets the manager recorded in the application properties.</summary>
		public string Manager { get; internal set; }

		/// <summary>Gets the name of the application that last saved the workbook.</summary>
		public string Application { get; internal set; }

		/// <summary>Gets the names of all the sheets of the workbook, in order, including the hidden ones.</summary>
		public IList<string> SheetNames { get; } = new List<string>();

		/// <summary>Gets the custom properties of the workbook.</summary>
		public IDictionary<string, string> CustomProperties { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
	}
}
