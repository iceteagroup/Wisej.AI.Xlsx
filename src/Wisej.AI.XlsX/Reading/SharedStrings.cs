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
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wisej.AI.XlsX.Packaging;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// The shared string table of a workbook, which stores the text of the cells.
	/// </summary>
	internal sealed class SharedStrings
	{
		private readonly List<string> _strings = new List<string>();

		/// <summary>
		/// Reads the shared strings part, streaming it since it can hold millions of strings.
		/// </summary>
		/// <param name="part">The shared strings part. Can be <c>null</c>.</param>
		public SharedStrings(OpenXmlPart part)
		{
			if (part == null)
				return;

			try
			{
				using (var reader = XlsxPackage.OpenReader(part))
				{
					if (reader.MoveToContent() != XmlNodeType.Element)
						return;

					XNamespace ns = reader.NamespaceURI;

					if (!reader.ReadToDescendant("si", ns.NamespaceName))
						return;

					while (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
					{
						var si = (XElement)XNode.ReadFrom(reader);
						_strings.Add(XmlExtensions.RichText(si, ns));

						reader.MoveToContent();
					}
				}
			}
			catch (XmlException)
			{
				// A damaged table keeps the strings read before the error.
			}
		}

		/// <summary>
		/// Returns the string at an index, or an empty string when the index is invalid.
		/// </summary>
		/// <param name="index">The zero based index.</param>
		/// <returns>The string.</returns>
		public string this[int index] => index >= 0 && index < _strings.Count ? _strings[index] : "";
	}
}
