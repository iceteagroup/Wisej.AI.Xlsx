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
using Wisej.AI.XlsX.Packaging;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// Resolves the number format of the cell styles (<c>cellXfs</c>) of a workbook.
	/// </summary>
	/// <remarks>
	/// Only the number format affects the text of a cell; fonts, fills and borders are ignored.
	/// </remarks>
	internal sealed class CellStyles
	{
		private readonly Dictionary<int, string> _customFormats = new Dictionary<int, string>();
		private readonly List<int> _formatIds = new List<int>();

		/// <summary>
		/// Initializes a new instance of the <see cref="CellStyles"/> class.
		/// </summary>
		/// <param name="styleSheet">The root element of the styles part. Can be <c>null</c>.</param>
		public CellStyles(XElement styleSheet)
		{
			if (styleSheet == null)
				return;

			foreach (var numFmt in styleSheet.Element(Ns.X + "numFmts")?.Elements(Ns.X + "numFmt") ?? new XElement[0])
			{
				var id = XmlExtensions.ParseInt(numFmt.Attr("numFmtId"));
				var code = numFmt.Attr("formatCode");
				if (id != null && code != null)
					_customFormats[id.Value] = code;
			}

			foreach (var xf in styleSheet.Element(Ns.X + "cellXfs")?.Elements(Ns.X + "xf") ?? new XElement[0])
				_formatIds.Add(XmlExtensions.ParseInt(xf.Attr("numFmtId")) ?? 0);
		}

		/// <summary>
		/// Returns the format code of a cell style.
		/// </summary>
		/// <param name="styleIndex">The index of the style in <c>cellXfs</c> (the <c>s</c> attribute of a cell).</param>
		/// <returns>The format code; <c>General</c> for unknown styles and formats.</returns>
		public string GetFormatCode(int styleIndex)
		{
			if (styleIndex < 0 || styleIndex >= _formatIds.Count)
				return "General";

			var id = _formatIds[styleIndex];
			if (_customFormats.TryGetValue(id, out var code))
				return code;

			return NumberFormatter.BuiltInFormat(id) ?? "General";
		}
	}
}
