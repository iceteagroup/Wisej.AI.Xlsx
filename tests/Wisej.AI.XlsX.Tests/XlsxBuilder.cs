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
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace Wisej.AI.XlsX.Tests
{
	/// <summary>
	/// Builds minimal SpreadsheetML packages from raw XML, so that every test controls exactly the markup it exercises.
	/// </summary>
	internal sealed class XlsxBuilder
	{
		/// <summary>The namespace declarations added to the root element of the workbook parts.</summary>
		public const string Namespaces =
			"xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
			"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
			"xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"";

		/// <summary>The namespace declarations of a drawing part.</summary>
		public const string DrawingNamespaces =
			"xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" " +
			"xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
			"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
			"xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"";

		private const string RelationshipBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

		private readonly Dictionary<string, byte[]> _parts = new Dictionary<string, byte[]>();
		private readonly List<(string PartName, string ContentType)> _overrides = new List<(string, string)>();
		private readonly List<(string Id, string Type, string Target, bool External)> _workbookRelationships = new List<(string, string, string, bool)>();
		private readonly List<string> _sheets = new List<string>();
		private readonly Dictionary<int, List<(string Id, string Type, string Target, bool External)>> _sheetRelationships = new Dictionary<int, List<(string, string, string, bool)>>();
		private string _workbookPr = "";

		/// <summary>
		/// Adds a worksheet.
		/// </summary>
		/// <param name="name">The name of the sheet.</param>
		/// <param name="xml">The inner XML of <c>worksheet</c>, such as a <c>sheetData</c> element.</param>
		/// <param name="state">The visibility state (<c>hidden</c>, <c>veryHidden</c>), or <c>null</c>.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder Sheet(string name, string xml, string state = null)
		{
			var number = _sheets.Count + 1;
			var id = "rIdSheet" + number;

			_sheets.Add($"<sheet name=\"{name}\" sheetId=\"{number}\" r:id=\"{id}\"{(state != null ? $" state=\"{state}\"" : "")}/>");
			_parts[$"xl/worksheets/sheet{number}.xml"] = Encoding.UTF8.GetBytes($"<worksheet {Namespaces}>{xml}</worksheet>");
			_overrides.Add(($"/xl/worksheets/sheet{number}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
			_workbookRelationships.Add((id, RelationshipBase + "worksheet", $"worksheets/sheet{number}.xml", false));
			_sheetRelationships[number] = new List<(string, string, string, bool)>();
			return this;
		}

		/// <summary>
		/// Adds a part related to a worksheet.
		/// </summary>
		/// <param name="sheet">The one based number of the sheet.</param>
		/// <param name="id">The relationship identifier used in the worksheet.</param>
		/// <param name="type">The relationship type, relative to the Office relationships namespace, or absolute.</param>
		/// <param name="partName">The name of the part, such as <c>xl/tables/table1.xml</c>.</param>
		/// <param name="contentType">The content type of the part.</param>
		/// <param name="content">The content of the part.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder SheetPart(int sheet, string id, string type, string partName, string contentType, byte[] content)
		{
			_parts[partName] = content;
			_overrides.Add(("/" + partName, contentType));
			_sheetRelationships[sheet].Add((id, type.StartsWith("http") ? type : RelationshipBase + type, "../" + partName.Substring(3), false));
			return this;
		}

		/// <summary>
		/// Adds an XML part related to a worksheet.
		/// </summary>
		/// <param name="sheet">The one based number of the sheet.</param>
		/// <param name="id">The relationship identifier used in the worksheet.</param>
		/// <param name="type">The relationship type.</param>
		/// <param name="partName">The name of the part.</param>
		/// <param name="contentType">The content type of the part.</param>
		/// <param name="xml">The XML of the part.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder SheetPart(int sheet, string id, string type, string partName, string contentType, string xml)
		{
			return SheetPart(sheet, id, type, partName, contentType, Encoding.UTF8.GetBytes(xml));
		}

		/// <summary>
		/// Adds an external hyperlink relationship to a worksheet.
		/// </summary>
		/// <param name="sheet">The one based number of the sheet.</param>
		/// <param name="id">The relationship identifier used in the <c>hyperlink</c> element.</param>
		/// <param name="url">The target.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder SheetHyperlink(int sheet, string id, string url)
		{
			_sheetRelationships[sheet].Add((id, RelationshipBase + "hyperlink", url, true));
			return this;
		}

		/// <summary>
		/// Adds a part related to another part (an image of a drawing, for example).
		/// </summary>
		/// <param name="ownerPart">The name of the owner part, such as <c>xl/drawings/drawing1.xml</c>.</param>
		/// <param name="relationships">The relationships XML elements of the owner part.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder PartRelationships(string ownerPart, string relationships)
		{
			var folder = ownerPart.Substring(0, ownerPart.LastIndexOf('/') + 1);
			var name = ownerPart.Substring(folder.Length);
			_parts[$"{folder}_rels/{name}.rels"] = Encoding.UTF8.GetBytes($"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{relationships}</Relationships>");
			return this;
		}

		/// <summary>
		/// Adds a part that is not related directly to a sheet or to the workbook.
		/// </summary>
		/// <param name="partName">The name of the part.</param>
		/// <param name="contentType">The content type, or <c>null</c> for a default extension type.</param>
		/// <param name="content">The content.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder Part(string partName, string contentType, byte[] content)
		{
			_parts[partName] = content;
			if (contentType != null)
				_overrides.Add(("/" + partName, contentType));
			return this;
		}

		/// <summary>
		/// Adds a part related to the workbook.
		/// </summary>
		/// <param name="id">The relationship identifier.</param>
		/// <param name="type">The relationship type, relative to the Office relationships namespace, or absolute.</param>
		/// <param name="partName">The name of the part.</param>
		/// <param name="contentType">The content type.</param>
		/// <param name="xml">The XML of the part.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder WorkbookPart(string id, string type, string partName, string contentType, string xml)
		{
			_parts[partName] = Encoding.UTF8.GetBytes(xml);
			_overrides.Add(("/" + partName, contentType));
			_workbookRelationships.Add((id, type.StartsWith("http") ? type : RelationshipBase + type, partName.Substring(3), false));
			return this;
		}

		/// <summary>
		/// Adds the styles part.
		/// </summary>
		/// <param name="cellXfs">The number format identifier of every cell style, in order.</param>
		/// <param name="numFmts">The custom number formats, as <c>numFmt</c> elements.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder Styles(int[] cellXfs, string numFmts = "")
		{
			var xfs = String.Concat(cellXfs.Select(id => $"<xf numFmtId=\"{id}\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"));
			var formats = numFmts.Length > 0 ? $"<numFmts>{numFmts}</numFmts>" : "";
			return WorkbookPart("rIdStyles", "styles", "xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
				$"<styleSheet {Namespaces}>{formats}<cellXfs count=\"{cellXfs.Length}\">{xfs}</cellXfs></styleSheet>");
		}

		/// <summary>
		/// Adds the shared strings part.
		/// </summary>
		/// <param name="strings">The strings.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder SharedStrings(params string[] strings)
		{
			var items = String.Concat(strings.Select(s => $"<si><t xml:space=\"preserve\">{System.Security.SecurityElement.Escape(s)}</t></si>"));
			return WorkbookPart("rIdStrings", "sharedStrings", "xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
				$"<sst {Namespaces} count=\"{strings.Length}\" uniqueCount=\"{strings.Length}\">{items}</sst>");
		}

		/// <summary>
		/// Sets the inner XML of <c>workbookPr</c> attributes, such as <c>date1904="1"</c>.
		/// </summary>
		/// <param name="attributes">The attributes.</param>
		/// <returns>This builder.</returns>
		public XlsxBuilder WorkbookProperties(string attributes)
		{
			_workbookPr = $"<workbookPr {attributes}/>";
			return this;
		}

		/// <summary>
		/// Builds the package.
		/// </summary>
		/// <returns>The package, positioned at the start.</returns>
		public MemoryStream Build()
		{
			var stream = new MemoryStream();

			using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
			{
				void Write(string name, string text)
				{
					using (var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)))
						writer.Write(text);
				}

				var types = new StringBuilder("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
				types.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
				types.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
				types.Append("<Default Extension=\"png\" ContentType=\"image/png\"/>");
				types.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
				foreach (var (partName, contentType) in _overrides)
					types.Append($"<Override PartName=\"{partName}\" ContentType=\"{contentType}\"/>");
				types.Append("</Types>");
				Write("[Content_Types].xml", types.ToString());

				Write("_rels/.rels",
					"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
					"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
					"</Relationships>");

				Write("xl/workbook.xml", $"<workbook {Namespaces}>{_workbookPr}<sheets>{String.Concat(_sheets)}</sheets></workbook>");
				Write("xl/_rels/workbook.xml.rels", Relationships(_workbookRelationships));

				foreach (var pair in _sheetRelationships.Where(p => p.Value.Count > 0))
					Write($"xl/worksheets/_rels/sheet{pair.Key}.xml.rels", Relationships(pair.Value));

				foreach (var part in _parts)
				{
					using (var entry = zip.CreateEntry(part.Key).Open())
						entry.Write(part.Value, 0, part.Value.Length);
				}
			}

			stream.Position = 0;
			return stream;
		}

		/// <summary>
		/// Builds the package and extracts it.
		/// </summary>
		/// <param name="options">The extraction options.</param>
		/// <returns>The extraction result.</returns>
		public XlsxExtractionResult Extract(XlsxExtractorOptions options)
		{
			using (var stream = Build())
				return new XlsxExtractor(options).Extract(stream);
		}

		/// <summary>
		/// Returns a relationships part.
		/// </summary>
		/// <param name="relationships">The relationships.</param>
		/// <returns>The XML of the part.</returns>
		private static string Relationships(IEnumerable<(string Id, string Type, string Target, bool External)> relationships)
		{
			var sb = new StringBuilder("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
			foreach (var (id, type, target, external) in relationships)
				sb.Append($"<Relationship Id=\"{id}\" Type=\"{type}\" Target=\"{System.Security.SecurityElement.Escape(target)}\"{(external ? " TargetMode=\"External\"" : "")}/>");
			sb.Append("</Relationships>");
			return sb.ToString();
		}
	}
}
