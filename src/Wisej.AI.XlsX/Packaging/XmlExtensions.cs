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
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Wisej.AI.XlsX.Packaging
{
	/// <summary>
	/// Helpers that read the attribute and value conventions used by SpreadsheetML.
	/// </summary>
	internal static class XmlExtensions
	{
		/// <summary>
		/// Namespaces of <c>mc:Choice</c> branches that the library cannot interpret
		/// (ink and 3D models): the <c>mc:Fallback</c> branch is used instead.
		/// </summary>
		private static readonly HashSet<string> UnsupportedChoiceNamespaces = new HashSet<string>(StringComparer.Ordinal)
		{
			"http://schemas.microsoft.com/office/drawing/2016/ink",
			"http://schemas.microsoft.com/office/drawing/2017/model3d",
		};

		#region Attributes

		/// <summary>
		/// Returns the value of the attribute with the specified name, or <c>null</c>.
		/// </summary>
		/// <param name="element">The element to read. Can be <c>null</c>.</param>
		/// <param name="name">The qualified attribute name.</param>
		/// <returns>The attribute value or <c>null</c> when the element or the attribute are missing.</returns>
		public static string Attr(this XElement element, XName name)
		{
			return element?.Attribute(name)?.Value;
		}

		/// <summary>
		/// Reads an <c>xsd:boolean</c> attribute value.
		/// </summary>
		/// <param name="value">The attribute value. Can be <c>null</c>.</param>
		/// <param name="defaultValue">Value returned when <paramref name="value"/> is <c>null</c> or empty.</param>
		/// <returns>The parsed value.</returns>
		public static bool ParseBool(string value, bool defaultValue)
		{
			if (String.IsNullOrEmpty(value))
				return defaultValue;

			switch (value.Trim().ToLowerInvariant())
			{
				case "0":
				case "false":
				case "off":
					return false;

				default:
					return true;
			}
		}

		/// <summary>
		/// Parses an integer attribute value, tolerating decimals and surrounding white space.
		/// </summary>
		/// <param name="value">The attribute value. Can be <c>null</c>.</param>
		/// <returns>The parsed value or <c>null</c> when the value is missing or invalid.</returns>
		public static int? ParseInt(string value)
		{
			if (String.IsNullOrWhiteSpace(value))
				return null;

			if (Int32.TryParse(value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
				return result;

			if (Double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double real) && Math.Abs(real) < Int32.MaxValue)
				return (int)Math.Round(real);

			return null;
		}

		/// <summary>
		/// Parses a floating point attribute value using the invariant culture.
		/// </summary>
		/// <param name="value">The attribute value. Can be <c>null</c>.</param>
		/// <returns>The parsed value or <c>null</c> when the value is missing or invalid.</returns>
		public static double? ParseDouble(string value)
		{
			if (String.IsNullOrWhiteSpace(value))
				return null;

			if (Double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double result))
				return result;

			return null;
		}

		/// <summary>
		/// Converts an English Metric Unit measure (<c>ST_Coordinate</c>) to points.
		/// </summary>
		/// <param name="value">The attribute value.</param>
		/// <returns>The measure in points, or <c>null</c> when the value is missing or invalid.</returns>
		public static double? EmuToPoints(string value)
		{
			return ParseDouble(value) / 12700.0;
		}

		#endregion

		#region Strings

		/// <summary>
		/// Decodes the <c>_xHHHH_</c> escape sequences that SpreadsheetML uses to store characters
		/// that XML cannot carry, such as carriage returns (<c>_x000D_</c>).
		/// </summary>
		/// <param name="text">The text as stored in the part. Can be <c>null</c>.</param>
		/// <returns>The decoded text.</returns>
		public static string DecodeEscapes(string text)
		{
			if (String.IsNullOrEmpty(text) || text.IndexOf("_x", StringComparison.Ordinal) < 0)
				return text;

			var sb = new StringBuilder(text.Length);

			for (var i = 0; i < text.Length; i++)
			{
				if (text[i] == '_' && i + 6 < text.Length && text[i + 1] == 'x' && text[i + 6] == '_'
					&& Int32.TryParse(text.Substring(i + 2, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var code))
				{
					// _x000D_ is the carriage return of a CR LF pair: the line feed follows it.
					if (code != 0x0D)
						sb.Append((char)code);

					i += 6;
				}
				else
				{
					sb.Append(text[i]);
				}
			}

			return sb.ToString();
		}

		/// <summary>
		/// Returns the text of a SpreadsheetML rich string (<c>si</c>, <c>is</c> or a comment <c>text</c>):
		/// its <c>t</c> element or the concatenation of the <c>t</c> elements of its runs, without phonetic hints.
		/// </summary>
		/// <param name="element">The string element. Can be <c>null</c>.</param>
		/// <param name="ns">The namespace of the elements.</param>
		/// <returns>The text, or an empty string.</returns>
		public static string RichText(XElement element, XNamespace ns)
		{
			if (element == null)
				return "";

			var sb = new StringBuilder();
			foreach (var child in element.Elements())
			{
				if (child.Name == ns + "t")
					sb.Append(child.Value);
				else if (child.Name == ns + "r")
					sb.Append(child.Element(ns + "t")?.Value);
			}

			return DecodeEscapes(sb.ToString());
		}

		#endregion

		#region Markup Compatibility

		/// <summary>
		/// Enumerates the child elements of <paramref name="element"/>, replacing every
		/// <c>mc:AlternateContent</c> with the content of the branch that should be processed.
		/// </summary>
		/// <param name="element">The parent element.</param>
		/// <returns>The child elements with markup compatibility resolved.</returns>
		/// <remarks>
		/// The first <c>mc:Choice</c> is preferred because it carries the richer, modern markup
		/// (extended charts, for example). The <c>mc:Fallback</c> branch is used only when the
		/// choice requires a vocabulary the library cannot interpret.
		/// </remarks>
		public static IEnumerable<XElement> ContentElements(this XElement element)
		{
			foreach (var child in element.Elements())
			{
				if (child.Name == Ns.MC + "AlternateContent")
				{
					var branch = SelectAlternateContentBranch(child);
					if (branch == null)
						continue;

					foreach (var resolved in branch.ContentElements())
						yield return resolved;
				}
				else
				{
					yield return child;
				}
			}
		}

		/// <summary>
		/// Picks the branch of an <c>mc:AlternateContent</c> element that should be processed.
		/// </summary>
		/// <param name="alternateContent">The <c>mc:AlternateContent</c> element.</param>
		/// <returns>The selected <c>mc:Choice</c> or <c>mc:Fallback</c> element, or <c>null</c>.</returns>
		private static XElement SelectAlternateContentBranch(XElement alternateContent)
		{
			foreach (var choice in alternateContent.Elements(Ns.MC + "Choice"))
			{
				if (IsChoiceSupported(choice))
					return choice;
			}

			return alternateContent.Element(Ns.MC + "Fallback");
		}

		/// <summary>
		/// Determines whether all the namespaces required by an <c>mc:Choice</c> can be interpreted.
		/// </summary>
		/// <param name="choice">The <c>mc:Choice</c> element.</param>
		/// <returns><c>true</c> when the choice can be processed.</returns>
		private static bool IsChoiceSupported(XElement choice)
		{
			var requires = choice.Attribute("Requires")?.Value;
			if (String.IsNullOrWhiteSpace(requires))
				return true;

			foreach (var prefix in requires.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
			{
				var ns = choice.GetNamespaceOfPrefix(prefix);
				if (ns == null || UnsupportedChoiceNamespaces.Contains(ns.NamespaceName))
					return false;
			}

			return true;
		}

		#endregion
	}
}
