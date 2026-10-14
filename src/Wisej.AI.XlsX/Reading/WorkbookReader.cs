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
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wisej.AI.XlsX.Model;
using Wisej.AI.XlsX.Packaging;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// Reads a workbook package into the <see cref="WorkbookModel"/>.
	/// </summary>
	internal sealed class WorkbookReader
	{
		private readonly XlsxPackage _package;
		private readonly XlsxExtractorOptions _options;

		/// <summary>
		/// Initializes a new instance of the <see cref="WorkbookReader"/> class.
		/// </summary>
		/// <param name="package">The opened package.</param>
		/// <param name="options">The extraction options.</param>
		public WorkbookReader(XlsxPackage package, XlsxExtractorOptions options)
		{
			_package = package;
			_options = options;
		}

		/// <summary>
		/// Reads the workbook.
		/// </summary>
		/// <returns>The workbook model.</returns>
		/// <exception cref="InvalidDataException">Thrown when the workbook part cannot be read.</exception>
		public WorkbookModel Read()
		{
			var workbookPart = _package.WorkbookPart;
			var workbook = _package.GetXml(workbookPart);
			if (workbook == null)
				throw new InvalidDataException("The workbook part does not contain well-formed XML.");

			var model = new WorkbookModel { Metadata = MetadataReader.Read(_package) };
			var date1904 = XmlExtensions.ParseBool(workbook.Element(Ns.X + "workbookPr")?.Attr("date1904"), false);

			var context = new ReadContext
			{
				Package = _package,
				Options = _options,
				Formatter = new NumberFormatter(_options.Culture, date1904),
				Styles = new CellStyles(_package.GetXml(XlsxPackage.GetRelatedPart(workbookPart, "/styles"))),
				Strings = new SharedStrings(XlsxPackage.GetRelatedPart(workbookPart, "/sharedStrings")),
				Images = new ImageRegistry(model.Images, _options.IncludeImageData || _options.ImageDescriber != null),
				Persons = CommentReader.ReadPersons(_package.GetXml(XlsxPackage.GetRelatedPart(workbookPart, "/person"))),
			};

			if (_options.IncludeImages)
				context.CellImages = new RichValueImages(context, workbookPart);

			var index = 0;
			foreach (var element in workbook.Element(Ns.X + "sheets")?.Elements(Ns.X + "sheet") ?? Enumerable.Empty<System.Xml.Linq.XElement>())
			{
				var name = element.Attr("name") ?? $"Sheet{index + 1}";
				var state = element.Attr("state");
				var position = index++;

				model.Metadata.SheetNames.Add(name);

				var hidden = state == "hidden" || state == "veryHidden";
				if (hidden && !_options.IncludeHiddenSheets)
					continue;

				var part = XlsxPackage.GetPartById(workbookPart, element.Attr(Ns.R + "id"));
				if (part == null)
					continue;

				var sheet = new Sheet { Name = name, Index = position };

				if (part.RelationshipType.EndsWith("/worksheet", StringComparison.OrdinalIgnoreCase))
				{
					new WorksheetReader(context, part, sheet).Read();
				}
				else if (part.RelationshipType.EndsWith("/chartsheet", StringComparison.OrdinalIgnoreCase))
				{
					sheet.IsChartSheet = true;
					ReadChartSheet(context, part, sheet);
				}
				else
				{
					// Dialog sheets and Excel 4 macro sheets carry no data.
					continue;
				}

				model.Sheets.Add(sheet);
			}

			return model;
		}

		/// <summary>
		/// Reads the chart of a chart sheet.
		/// </summary>
		/// <param name="context">The read context.</param>
		/// <param name="part">The chart sheet part.</param>
		/// <param name="sheet">The sheet that receives the chart.</param>
		private static void ReadChartSheet(ReadContext context, OpenXmlPart part, Sheet sheet)
		{
			var root = context.Package.GetXml(part);
			var drawingId = root?.Element(Ns.X + "drawing")?.Attr(Ns.R + "id");
			var drawingPart = XlsxPackage.GetPartById(part, drawingId) ?? XlsxPackage.GetRelatedPart(part, "/drawing");

			if (drawingPart != null)
				DrawingReader.Read(context, drawingPart, sheet);
		}
	}
}
