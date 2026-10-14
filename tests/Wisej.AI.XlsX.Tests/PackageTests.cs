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
using System.IO.Compression;
using System.Text;
using Xunit;

namespace Wisej.AI.XlsX.Tests
{
	/// <summary>
	/// Tests the handling of packages: invalid input, unsupported formats and damaged parts.
	/// </summary>
	public class PackageTests : TestBase
	{
		[Fact]
		public void NullArguments_Throw()
		{
			Assert.Throws<ArgumentNullException>(() => new XlsxExtractor(null));
			Assert.Throws<ArgumentNullException>(() => new XlsxExtractor().Extract((Stream)null));
			Assert.Throws<ArgumentNullException>(() => new XlsxExtractor().Extract((string)null));
		}

		[Fact]
		public void RandomBytes_ThrowInvalidData()
		{
			using (var stream = new MemoryStream(Encoding.ASCII.GetBytes("this is not a workbook")))
				Assert.Throws<InvalidDataException>(() => new XlsxExtractor().Extract(stream));
		}

		[Fact]
		public void LegacyXls_IsReported()
		{
			var bytes = new byte[512];
			new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(bytes, 0);

			using (var stream = new MemoryStream(bytes))
			{
				var ex = Assert.Throws<InvalidDataException>(() => new XlsxExtractor().Extract(stream));
				Assert.Contains(".xls", ex.Message);
			}
		}

		[Fact]
		public void WordDocument_IsRejected()
		{
			var stream = new MemoryStream();
			using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
			{
				void Write(string name, string text)
				{
					using (var writer = new StreamWriter(zip.CreateEntry(name).Open()))
						writer.Write(text);
				}

				Write("[Content_Types].xml",
					"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
					"<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
					"<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
				Write("_rels/.rels",
					"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
					"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
				Write("word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>");
			}

			stream.Position = 0;
			Assert.Throws<InvalidDataException>(() => new XlsxExtractor().Extract(stream));
		}

		[Fact]
		public void MissingSharedStrings_LeaveCellsEmpty()
		{
			var builder = new XlsxBuilder().Sheet("Data", Data(
				Row(1, S("A1", 5), T("B1", "Kept"))));

			Assert.Equal("## Data\n\nKept", Markdown(builder));
		}

		[Fact]
		public void Date1904_IsHonored()
		{
			var builder = new XlsxBuilder()
				.WorkbookProperties("date1904=\"1\"")
				.Styles(new[] { 0, 14 })
				.Sheet("Dates", Data(Row(1, N("A1", 0, 1))));

			Assert.Equal("## Dates\n\n1904-01-01", Markdown(builder, o => o.TableHeaders = TableHeaderMode.ColumnLetters));
		}

		[Fact]
		public void ExtractText_ReturnsTheText()
		{
			var builder = new XlsxBuilder().Sheet("Data", Data(Row(1, T("A1", "Hello"))));

			using (var stream = builder.Build())
				Assert.Equal("<!-- Page 1 -->\n\n## Data\n\nHello", XlsxExtractor.ExtractText(stream));
		}
	}
}
