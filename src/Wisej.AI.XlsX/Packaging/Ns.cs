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
using System.Xml.Linq;

namespace Wisej.AI.XlsX.Packaging
{
	/// <summary>
	/// XML namespaces used by SpreadsheetML workbooks and the DrawingML parts they reference.
	/// </summary>
	/// <remarks>
	/// Workbooks saved in the ISO/IEC 29500 <i>Strict</i> conformance class use different namespace
	/// URIs for the same vocabularies. <see cref="StrictToTransitional"/> maps them so that the rest of
	/// the library only needs to know the transitional names.
	/// </remarks>
	internal static class Ns
	{
		/// <summary>SpreadsheetML main namespace (default namespace of the workbook parts).</summary>
		public static readonly XNamespace X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

		/// <summary>SpreadsheetML main namespace of the Strict conformance class.</summary>
		public static readonly XNamespace XStrict = "http://purl.oclc.org/ooxml/spreadsheetml/main";

		/// <summary>Office relationships namespace (<c>r:</c>).</summary>
		public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

		/// <summary>Office relationships namespace of the Strict conformance class.</summary>
		public static readonly XNamespace RStrict = "http://purl.oclc.org/ooxml/officeDocument/relationships";

		/// <summary>Spreadsheet drawing namespace (<c>xdr:</c>).</summary>
		public static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

		/// <summary>DrawingML main namespace (<c>a:</c>).</summary>
		public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

		/// <summary>DrawingML chart namespace (<c>c:</c>).</summary>
		public static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

		/// <summary>DrawingML extended chart namespace used by Office 2016 chart types (<c>cx:</c>).</summary>
		public static readonly XNamespace CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";

		/// <summary>Markup compatibility namespace (<c>mc:</c>).</summary>
		public static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

		/// <summary>Decorative image extension namespace (<c>adec:</c>).</summary>
		public static readonly XNamespace ADec = "http://schemas.microsoft.com/office/drawing/2017/decorative";

		/// <summary>Threaded comments namespace (Excel 365 comments with replies).</summary>
		public static readonly XNamespace ThreadedComments = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

		/// <summary>Package relationships namespace used by <c>.rels</c> parts.</summary>
		public static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

		/// <summary>Extended (application) properties namespace used by <c>docProps/app.xml</c>.</summary>
		public static readonly XNamespace ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

		/// <summary>Custom properties namespace used by <c>docProps/custom.xml</c>.</summary>
		public static readonly XNamespace CustomProperties = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

		/// <summary><c>graphicData/@uri</c> of a chart.</summary>
		public const string GraphicChart = "http://schemas.openxmlformats.org/drawingml/2006/chart";

		/// <summary><c>graphicData/@uri</c> of an extended (Office 2016) chart.</summary>
		public const string GraphicChartEx = "http://schemas.microsoft.com/office/drawing/2014/chartex";

		/// <summary>
		/// Maps the namespaces of the Strict conformance class to their transitional equivalents.
		/// </summary>
		public static readonly IReadOnlyDictionary<string, XNamespace> StrictToTransitional = new Dictionary<string, XNamespace>
		{
			[XStrict.NamespaceName] = X,
			[RStrict.NamespaceName] = R,
			["http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"] = Xdr,
			["http://purl.oclc.org/ooxml/drawingml/main"] = A,
			["http://purl.oclc.org/ooxml/drawingml/chart"] = C,
			["http://purl.oclc.org/ooxml/officeDocument/extendedProperties"] = ExtendedProperties,
			["http://purl.oclc.org/ooxml/officeDocument/customProperties"] = CustomProperties,
		};
	}
}
