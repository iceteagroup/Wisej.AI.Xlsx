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
using Wisej.AI.XlsX.Model;
using Wisej.AI.XlsX.Packaging;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// Reads the comments of a worksheet: threaded comments with their replies (Excel 365) and notes
	/// (the comments of the earlier versions).
	/// </summary>
	/// <remarks>
	/// Excel also writes every threaded comment as a note, prefixed with an explanation for the older
	/// versions; those copies are skipped when the threaded comment is read.
	/// </remarks>
	internal static class CommentReader
	{
		/// <summary>
		/// Reads the authors of the threaded comments.
		/// </summary>
		/// <param name="personList">The root element of the persons part. Can be <c>null</c>.</param>
		/// <returns>The display names, by person identifier.</returns>
		public static Dictionary<string, string> ReadPersons(XElement personList)
		{
			var persons = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

			foreach (var person in personList?.Elements(Ns.ThreadedComments + "person") ?? Enumerable.Empty<XElement>())
			{
				var id = person.Attr("id");
				if (id != null)
					persons[id] = person.Attr("displayName");
			}

			return persons;
		}

		/// <summary>
		/// Reads the comments of a worksheet into <see cref="Sheet.Comments"/>.
		/// </summary>
		/// <param name="context">The read context.</param>
		/// <param name="part">The worksheet part.</param>
		/// <param name="sheet">The sheet.</param>
		public static void Read(ReadContext context, OpenXmlPart part, Sheet sheet)
		{
			if (!context.Options.IncludeComments)
				return;

			var threadedCells = ReadThreaded(context, part, sheet);
			ReadNotes(context, part, sheet, threadedCells);

			sheet.Comments.Sort((a, b) => a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Column.CompareTo(b.Column));
		}

		/// <summary>
		/// Reads the threaded comments.
		/// </summary>
		/// <param name="context">The read context.</param>
		/// <param name="part">The worksheet part.</param>
		/// <param name="sheet">The sheet.</param>
		/// <returns>The addresses of the cells that have a threaded comment.</returns>
		private static HashSet<string> ReadThreaded(ReadContext context, OpenXmlPart part, Sheet sheet)
		{
			var cells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
			var root = context.Package.GetXml(XlsxPackage.GetRelatedPart(part, "/threadedComment"));
			if (root == null)
				return cells;

			var threads = new Dictionary<string, SheetComment>(StringComparer.OrdinalIgnoreCase);

			foreach (var element in root.Elements(Ns.ThreadedComments + "threadedComment"))
			{
				var reference = element.Attr("ref");
				if (!CellRange.TryParseCell(reference, out var row, out var column) || row == 0 || column == 0)
					continue;

				cells.Add(reference);

				var text = element.Element(Ns.ThreadedComments + "text")?.Value?.Trim();
				if (String.IsNullOrEmpty(text) || sheet.IsHidden(row, column))
					continue;

				var personId = element.Attr("personId");
				var comment = new SheetComment
				{
					Cell = reference,
					Row = row,
					Column = column,
					Author = personId != null && context.Persons.TryGetValue(personId, out var name) ? name : null,
					Text = text,
				};

				var parentId = element.Attr("parentId");
				if (parentId != null && threads.TryGetValue(parentId, out var parent))
				{
					parent.Replies.Add(comment);
				}
				else
				{
					sheet.Comments.Add(comment);
					var id = element.Attr("id");
					if (id != null)
						threads[id] = comment;
				}
			}

			return cells;
		}

		/// <summary>
		/// Reads the notes, skipping the copies of the threaded comments.
		/// </summary>
		/// <param name="context">The read context.</param>
		/// <param name="part">The worksheet part.</param>
		/// <param name="sheet">The sheet.</param>
		/// <param name="threadedCells">The addresses of the cells that have a threaded comment.</param>
		private static void ReadNotes(ReadContext context, OpenXmlPart part, Sheet sheet, HashSet<string> threadedCells)
		{
			var root = context.Package.GetXml(XlsxPackage.GetRelatedPart(part, "/comments"));
			if (root == null)
				return;

			var authors = root.Element(Ns.X + "authors")?.Elements(Ns.X + "author").Select(a => a.Value).ToList() ?? new List<string>();

			foreach (var element in root.Element(Ns.X + "commentList")?.Elements(Ns.X + "comment") ?? Enumerable.Empty<XElement>())
			{
				var reference = element.Attr("ref");
				if (!CellRange.TryParseCell(reference, out var row, out var column) || row == 0 || column == 0)
					continue;

				if (threadedCells.Contains(reference) || sheet.IsHidden(row, column))
					continue;

				var text = XmlExtensions.RichText(element.Element(Ns.X + "text"), Ns.X).Trim();
				if (text.StartsWith("[Threaded comment]", StringComparison.Ordinal))
					continue;

				var authorId = XmlExtensions.ParseInt(element.Attr("authorId")) ?? -1;
				var author = authorId >= 0 && authorId < authors.Count ? authors[authorId] : null;

				// Threaded comment placeholders use "tc={guid}" as the author.
				if (author != null && author.StartsWith("tc=", StringComparison.Ordinal))
					author = null;

				// Excel starts the text of a note with the name of its author.
				if (!String.IsNullOrEmpty(author) && text.StartsWith(author + ":", StringComparison.Ordinal))
					text = text.Substring(author.Length + 1).Trim();

				if (text.Length == 0)
					continue;

				sheet.Comments.Add(new SheetComment
				{
					Cell = reference,
					Row = row,
					Column = column,
					Author = String.IsNullOrWhiteSpace(author) ? null : author.Trim(),
					Text = text,
				});
			}
		}
	}
}
