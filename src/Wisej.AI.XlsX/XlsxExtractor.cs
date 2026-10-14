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
using Wisej.AI.XlsX.Model;
using Wisej.AI.XlsX.Packaging;
using Wisej.AI.XlsX.Reading;
using Wisej.AI.XlsX.Rendering;

namespace Wisej.AI.XlsX
{
	/// <summary>
	/// Extracts the content of Excel workbooks (XLSX, XLSM, XLTX, XLTM and Flat OPC) as Markdown or
	/// plain text, one page per sheet, ready to be used as language model context.
	/// </summary>
	/// <remarks>
	/// <para>
	/// Every value is written as Excel displays it: numbers, currencies, percentages, dates and times
	/// are formatted with their number format. Each sheet is split into its blocks of cells: Excel
	/// tables and ranges become Markdown tables (with merged cells resolved and empty rows and columns
	/// removed) and isolated cells such as titles become paragraphs. Hyperlinks, formulas, comments,
	/// charts (as data tables), pictures (floating or placed in cells), text boxes and the workbook
	/// properties are extracted as well.
	/// </para>
	/// <para>
	/// Instances are not thread safe but are cheap: create one per thread or per workbook.
	/// </para>
	/// <para>
	/// Example:
	/// <code><![CDATA[
	/// var extractor = new XlsxExtractor();
	/// var result = extractor.Extract("sales.xlsx");
	///
	/// foreach (var page in result.Pages)
	///     Console.WriteLine($"{page.SheetName}: {page.Text.Length} characters");
	///
	/// File.WriteAllText("sales.md", result.Text);
	/// ]]></code>
	/// </para>
	/// </remarks>
	public sealed class XlsxExtractor
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="XlsxExtractor"/> class with the default options.
		/// </summary>
		public XlsxExtractor()
			: this(new XlsxExtractorOptions())
		{
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="XlsxExtractor"/> class.
		/// </summary>
		/// <param name="options">The extraction options.</param>
		/// <exception cref="ArgumentNullException">Thrown when <paramref name="options"/> is <c>null</c>.</exception>
		public XlsxExtractor(XlsxExtractorOptions options)
		{
			Options = options ?? throw new ArgumentNullException(nameof(options));
		}

		/// <summary>
		/// Gets the extraction options.
		/// </summary>
		public XlsxExtractorOptions Options { get; }

		#region Extract

		/// <summary>
		/// Extracts the content of an Excel workbook stored in a file.
		/// </summary>
		/// <param name="path">The path of the workbook.</param>
		/// <returns>The extracted content.</returns>
		/// <exception cref="ArgumentNullException">Thrown when <paramref name="path"/> is <c>null</c>.</exception>
		/// <exception cref="InvalidDataException">Thrown when the file is not a valid Excel workbook.</exception>
		public XlsxExtractionResult Extract(string path)
		{
			if (path == null)
				throw new ArgumentNullException(nameof(path));

			using (var stream = File.OpenRead(path))
				return Extract(stream);
		}

		/// <summary>
		/// Extracts the content of an Excel workbook from a stream.
		/// </summary>
		/// <param name="stream">The stream containing the workbook. It is read from its current position and is not closed.</param>
		/// <returns>The extracted content.</returns>
		/// <exception cref="ArgumentNullException">Thrown when <paramref name="stream"/> is <c>null</c>.</exception>
		/// <exception cref="InvalidDataException">Thrown when the stream does not contain a valid Excel workbook.</exception>
		public XlsxExtractionResult Extract(Stream stream)
		{
			if (stream == null)
				throw new ArgumentNullException(nameof(stream));

			var model = ReadModel(stream);

			DescribeImages(model);

			var renderer = Options.Format == OutputFormat.PlainText
				? (WorkbookRenderer)new PlainTextRenderer(model, Options)
				: new MarkdownRenderer(model, Options);

			var result = renderer.Render();

			// The bytes were only loaded for the describer.
			if (!Options.IncludeImageData)
			{
				foreach (var image in model.Images)
					image.Data = null;
			}

			return result;
		}

		/// <summary>
		/// Extracts the text of an Excel workbook.
		/// </summary>
		/// <param name="stream">The stream containing the workbook. It is read from its current position and is not closed.</param>
		/// <param name="options">The extraction options, or <c>null</c> for the defaults.</param>
		/// <returns>The text of the workbook, with a header line introducing every page.</returns>
		/// <exception cref="ArgumentNullException">Thrown when <paramref name="stream"/> is <c>null</c>.</exception>
		/// <exception cref="InvalidDataException">Thrown when the stream does not contain a valid Excel workbook.</exception>
		public static string ExtractText(Stream stream, XlsxExtractorOptions options = null)
		{
			return new XlsxExtractor(options ?? new XlsxExtractorOptions()).Extract(stream).Text;
		}

		#endregion

		#region Implementation

		/// <summary>
		/// Reads a workbook into the model.
		/// </summary>
		/// <param name="stream">The stream containing the workbook.</param>
		/// <returns>The workbook model.</returns>
		internal WorkbookModel ReadModel(Stream stream)
		{
			var buffer = XlsxPackage.Buffer(stream);

			try
			{
				return ReadModel(buffer, repair: false);
			}
			catch (Exception ex) when (XlsxPackage.IsRepairable(ex) && !IsUnsupportedFormat(ex))
			{
				// Excel opens many packages that do not strictly follow the Open Packaging Conventions.
				return ReadModel(buffer, repair: true);
			}
		}

		/// <summary>
		/// Opens a buffered workbook and reads it into the model.
		/// </summary>
		/// <param name="buffer">The buffered workbook.</param>
		/// <param name="repair">Whether the package is repaired before it is opened.</param>
		/// <returns>The workbook model.</returns>
		private WorkbookModel ReadModel(MemoryStream buffer, bool repair)
		{
			using (var package = XlsxPackage.Open(buffer, repair))
				return new WorkbookReader(package, Options).Read();
		}

		/// <summary>
		/// Determines whether an exception reports a format that repairing cannot read (legacy or binary workbooks).
		/// </summary>
		/// <param name="ex">The exception.</param>
		/// <returns><c>true</c> when the format is not supported.</returns>
		private static bool IsUnsupportedFormat(Exception ex)
		{
			return ex is InvalidDataException && ex.InnerException == null && ex.Message.Contains("not supported");
		}

		/// <summary>
		/// Invokes the image describer once for every distinct picture and assigns the description to all its placements.
		/// </summary>
		/// <param name="model">The workbook model.</param>
		private void DescribeImages(WorkbookModel model)
		{
			var describer = Options.ImageDescriber;
			if (describer == null)
				return;

			var descriptions = new Dictionary<string, string>(StringComparer.Ordinal);

			foreach (var image in model.Images)
			{
				var key = image.PartUri ?? image.FileName;
				if (!descriptions.TryGetValue(key, out var description))
				{
					description = describer(image);
					descriptions[key] = description;
				}

				image.Description = String.IsNullOrWhiteSpace(description) ? null : description.Trim();
			}
		}

		#endregion
	}
}
