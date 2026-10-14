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
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wisej.AI.XlsX.Packaging;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// Resolves the pictures placed in cells (Excel 365 Place in Cell, and the <c>IMAGE</c> function results that are cached).
	/// </summary>
	/// <remarks>
	/// <para>
	/// A cell that contains a picture stores a <c>#VALUE!</c> error with a value metadata index
	/// (<c>vm</c>). The picture is found through a chain of parts:
	/// </para>
	/// <list type="number">
	/// <item><c>xl/metadata.xml</c> maps the value metadata to a rich value index (<c>XLRICHVALUE</c>).</item>
	/// <item><c>xl/richData/rdrichvalue.xml</c> holds the rich value, whose structure (<c>rdrichvaluestructure.xml</c>) names its keys.</item>
	/// <item>The <c>_rvRel:LocalImageIdentifier</c> key indexes <c>richValueRel.xml</c>, whose relationship targets the image part.</item>
	/// </list>
	/// </remarks>
	internal sealed class RichValueImages
	{
		private const string ImageKey = "_rvRel:LocalImageIdentifier";
		private const string AltTextKey = "Text";

		private readonly ReadContext _context;
		private readonly OpenXmlPart _relPart;
		private readonly List<int> _valueMetadata = new List<int>();
		private readonly List<(int Structure, List<string> Values)> _richValues = new List<(int, List<string>)>();
		private readonly List<List<string>> _structures = new List<List<string>>();
		private readonly List<string> _relationships = new List<string>();

		/// <summary>
		/// Initializes a new instance of the <see cref="RichValueImages"/> class.
		/// </summary>
		/// <param name="context">The read context.</param>
		/// <param name="workbookPart">The workbook part.</param>
		public RichValueImages(ReadContext context, OpenXmlPart workbookPart)
		{
			_context = context;

			var package = context.Package;
			var metadata = package.GetXml(XlsxPackage.GetRelatedPart(workbookPart, "/sheetMetadata"));
			var richValues = package.GetXml(XlsxPackage.GetRelatedPart(workbookPart, "/rdRichValue"));
			var structures = package.GetXml(XlsxPackage.GetRelatedPart(workbookPart, "/rdRichValueStructure"));
			_relPart = XlsxPackage.GetRelatedPart(workbookPart, "/richValueRel");
			var relationships = package.GetXml(_relPart);

			if (metadata == null || richValues == null || structures == null || relationships == null)
				return;

			// Value metadata: rc/@t is the one based index of the metadata type, rc/@v the index in its future metadata.
			var types = metadata.Element(Ns.X + "metadataTypes")?.Elements(Ns.X + "metadataType").Select(t => t.Attr("name")).ToList() ?? new List<string>();
			var future = metadata.Elements(Ns.X + "futureMetadata")
				.FirstOrDefault(f => f.Attr("name") == "XLRICHVALUE")
				?.Elements(Ns.X + "bk")
				.Select(bk => XmlExtensions.ParseInt(bk.Descendants().FirstOrDefault(d => d.Name.LocalName == "rvb").Attr("i")) ?? -1)
				.ToList() ?? new List<int>();

			foreach (var bk in metadata.Element(Ns.X + "valueMetadata")?.Elements(Ns.X + "bk") ?? Enumerable.Empty<XElement>())
			{
				var rc = bk.Element(Ns.X + "rc");
				var type = XmlExtensions.ParseInt(rc.Attr("t")) ?? 0;
				var index = XmlExtensions.ParseInt(rc.Attr("v")) ?? -1;

				var isRichValue = type >= 1 && type <= types.Count && types[type - 1] == "XLRICHVALUE";
				_valueMetadata.Add(isRichValue && index >= 0 && index < future.Count ? future[index] : -1);
			}

			foreach (var s in structures.Elements().Where(e => e.Name.LocalName == "s"))
				_structures.Add(s.Elements().Where(k => k.Name.LocalName == "k").Select(k => k.Attr("n")).ToList());

			foreach (var rv in richValues.Elements().Where(e => e.Name.LocalName == "rv"))
				_richValues.Add((XmlExtensions.ParseInt(rv.Attr("s")) ?? -1, rv.Elements().Where(v => v.Name.LocalName == "v").Select(v => v.Value).ToList()));

			foreach (var rel in relationships.Elements().Where(e => e.Name.LocalName == "rel"))
				_relationships.Add(rel.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value);
		}

		/// <summary>
		/// Returns the picture of a cell, creating its placement.
		/// </summary>
		/// <param name="valueMetadata">The one based value metadata index (the <c>vm</c> attribute of the cell).</param>
		/// <param name="sheetName">The name of the sheet.</param>
		/// <param name="cell">The address of the cell.</param>
		/// <returns>The placement, or <c>null</c> when the value is not a picture.</returns>
		public XlsxImage GetImage(int valueMetadata, string sheetName, string cell)
		{
			var index = valueMetadata - 1;
			if (index < 0 || index >= _valueMetadata.Count)
				return null;

			var richIndex = _valueMetadata[index];
			if (richIndex < 0 || richIndex >= _richValues.Count)
				return null;

			var (structure, values) = _richValues[richIndex];
			if (structure < 0 || structure >= _structures.Count)
				return null;

			var keys = _structures[structure];
			var position = keys.IndexOf(ImageKey);
			if (position < 0 || position >= values.Count)
				return null;

			var relIndex = XmlExtensions.ParseInt(values[position]) ?? -1;
			if (relIndex < 0 || relIndex >= _relationships.Count)
				return null;

			var image = _context.Images.Create(_relPart, _relationships[relIndex], null, sheetName);
			if (image == null)
				return null;

			var altPosition = keys.IndexOf(AltTextKey);
			if (altPosition >= 0 && altPosition < values.Count && !String.IsNullOrWhiteSpace(values[altPosition]))
				image.AltText = values[altPosition].Trim();

			image.Cell = cell;
			image.IsInCell = true;
			return image;
		}
	}
}
