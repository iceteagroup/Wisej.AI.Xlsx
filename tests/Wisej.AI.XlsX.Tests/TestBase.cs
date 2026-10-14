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

namespace Wisej.AI.XlsX.Tests
{
	/// <summary>
	/// Helpers shared by the tests: markup shortcuts and extraction with test friendly defaults.
	/// </summary>
	public abstract class TestBase
	{
		/// <summary>
		/// Returns a <c>sheetData</c> element.
		/// </summary>
		/// <param name="rows">The rows markup.</param>
		/// <returns>The markup.</returns>
		protected static string Data(params string[] rows)
		{
			return "<sheetData>" + String.Concat(rows) + "</sheetData>";
		}

		/// <summary>
		/// Returns a row.
		/// </summary>
		/// <param name="number">The one based row number.</param>
		/// <param name="cells">The cells markup.</param>
		/// <returns>The markup.</returns>
		protected static string Row(int number, params string[] cells)
		{
			return $"<row r=\"{number}\">{String.Concat(cells)}</row>";
		}

		/// <summary>
		/// Returns a cell with an inline string.
		/// </summary>
		/// <param name="reference">The cell address.</param>
		/// <param name="text">The text.</param>
		/// <returns>The markup.</returns>
		protected static string T(string reference, string text)
		{
			return $"<c r=\"{reference}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{System.Security.SecurityElement.Escape(text)}</t></is></c>";
		}

		/// <summary>
		/// Returns a cell with a number.
		/// </summary>
		/// <param name="reference">The cell address.</param>
		/// <param name="value">The value.</param>
		/// <param name="style">The cell style index, or <c>-1</c> for none.</param>
		/// <returns>The markup.</returns>
		protected static string N(string reference, double value, int style = -1)
		{
			var s = style >= 0 ? $" s=\"{style}\"" : "";
			return $"<c r=\"{reference}\"{s}><v>{value.ToString("R", CultureInfo.InvariantCulture)}</v></c>";
		}

		/// <summary>
		/// Returns a cell that references a shared string.
		/// </summary>
		/// <param name="reference">The cell address.</param>
		/// <param name="index">The index of the shared string.</param>
		/// <returns>The markup.</returns>
		protected static string S(string reference, int index)
		{
			return $"<c r=\"{reference}\" t=\"s\"><v>{index}</v></c>";
		}

		/// <summary>
		/// Extracts a workbook as Markdown without page separators.
		/// </summary>
		/// <param name="builder">The workbook builder.</param>
		/// <param name="configure">Changes the options before the extraction. Can be <c>null</c>.</param>
		/// <returns>The extracted text.</returns>
		internal static string Markdown(XlsxBuilder builder, Action<XlsxExtractorOptions> configure = null)
		{
			var options = new XlsxExtractorOptions { PageHeaderFormat = "" };
			configure?.Invoke(options);

			return builder.Extract(options).Text.Replace("\r\n", "\n");
		}

		/// <summary>
		/// Extracts a workbook as plain text without page separators.
		/// </summary>
		/// <param name="builder">The workbook builder.</param>
		/// <param name="configure">Changes the options before the extraction. Can be <c>null</c>.</param>
		/// <returns>The extracted text.</returns>
		internal static string PlainText(XlsxBuilder builder, Action<XlsxExtractorOptions> configure = null)
		{
			return Markdown(builder, options =>
			{
				options.Format = OutputFormat.PlainText;
				configure?.Invoke(options);
			});
		}
	}
}
