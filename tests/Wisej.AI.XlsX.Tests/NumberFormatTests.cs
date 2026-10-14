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
using Wisej.AI.XlsX.Reading;
using Xunit;

namespace Wisej.AI.XlsX.Tests
{
	/// <summary>
	/// Tests the formatting of values with Excel number format codes.
	/// </summary>
	public class NumberFormatTests
	{
		private static readonly NumberFormatter Formatter = new NumberFormatter(CultureInfo.InvariantCulture, false);

		[Theory]
		[InlineData("General", 0.30000000000000004, "0.3")]
		[InlineData("General", 1234567890123, "1234567890123")]
		[InlineData("General", -42, "-42")]
		[InlineData("General", 0.125, "0.125")]
		[InlineData("0", 2.5, "3")]
		[InlineData("0.00", 2.675, "2.68")]
		[InlineData("0.00", -3.14159, "-3.14")]
		[InlineData("#,##0", 1234567, "1,234,567")]
		[InlineData("#,##0", 5, "5")]
		[InlineData("#,##0.00", -1234.5, "-1,234.50")]
		[InlineData("#,##0.##", 1234.5, "1,234.5")]
		[InlineData("#.00", 0.5, ".50")]
		[InlineData("000-0000", 5551234, "555-1234")]
		[InlineData("00000", 42, "00042")]
		[InlineData("0%", 0.256, "26%")]
		[InlineData("0.0%", 0.1234, "12.3%")]
		[InlineData("#,##0,\"K\"", 15300, "15K")]
		[InlineData("#,##0.0,,\"M\"", 2500000, "2.5M")]
		[InlineData("0.00E+00", 12345, "1.23E+04")]
		[InlineData("0.00E+00", 0.00012345, "1.23E-04")]
		[InlineData("##0.0E+0", 12345, "12.3E+3")]
		[InlineData("# ?/?", 1.75, "1 3/4")]
		[InlineData("# ??/??", 3.14159, "3 14/99")]
		[InlineData("# ?/8", 0.5, "4/8")]
		[InlineData("?/?", 2.5, "5/2")]
		[InlineData("# ?/?", 3, "3")]
		[InlineData("\"$\"#,##0.00_);(\"$\"#,##0.00)", 1234.5, "$1,234.50")]
		[InlineData("\"$\"#,##0.00_);(\"$\"#,##0.00)", -1234.5, "($1,234.50)")]
		[InlineData("$#,##0", -5, "-$5")]
		[InlineData("[$€-407] #,##0.00", 1234.5, "€ 1,234.50")]
		[InlineData("#,##0.00 [$€-1]", 99, "99.00 €")]
		[InlineData("0.00;[Red]-0.00", -2, "-2.00")]
		[InlineData("0;-0;\"zero\"", 0, "zero")]
		[InlineData("0;-0;;@", 0, "")]
		[InlineData("[<1000]0;#,##0,\"K\"", 999, "999")]
		[InlineData("[<1000]0;#,##0,\"K\"", 15000, "15K")]
		[InlineData("[>=100]\"High\";[<0]\"Negative\";\"Low\"", -3, "Negative")]
		[InlineData("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", -12.5, "(12.50)")]
		[InlineData("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", 0, "-")]
		[InlineData("0 \"units\"", 12, "12 units")]
		[InlineData("General\" kg\"", 2.5, "2.5 kg")]
		[InlineData("\\(0\\)", 7, "(7)")]
		public void FormatsNumbers(string format, double value, string expected)
		{
			Assert.Equal(expected, Formatter.Format(value, format).Trim());
		}

		[Theory]
		[InlineData("yyyy-mm-dd", 44927, "2023-01-01")]
		[InlineData("yyyy-mm-dd", 45214, "2023-10-15")]
		[InlineData("d-mmm-yy", 44927, "1-Jan-23")]
		[InlineData("mmmm d, yyyy", 44927, "January 1, 2023")]
		[InlineData("dddd", 44927, "Sunday")]
		[InlineData("ddd dd/mm/yyyy", 44928, "Mon 02/01/2023")]
		[InlineData("mmm-yy", 44927, "Jan-23")]
		[InlineData("mmmmm", 44927, "J")]
		[InlineData("h:mm AM/PM", 0.75, "6:00 PM")]
		[InlineData("h:mm am/pm", 0.25, "6:00 am")]
		[InlineData("h:mm A/P", 0.75, "6:00 P")]
		[InlineData("hh:mm:ss", 0.5, "12:00:00")]
		[InlineData("h:mm", 0.4375, "10:30")]
		[InlineData("mm:ss", 0.0006944444444444445, "01:00")]
		[InlineData("[h]:mm", 1.5, "36:00")]
		[InlineData("[mm]:ss", 0.125, "180:00")]
		[InlineData("mm:ss.0", 0.00001273148148, "00:01.1")]
		[InlineData("yyyy-mm-dd hh:mm", 44927.5, "2023-01-01 12:00")]
		[InlineData("hh:mm:ss", 0.43749999, "10:30:00")]
		[InlineData("yyyy-mm-dd", 1, "1900-01-01")]
		[InlineData("yyyy-mm-dd", 59, "1900-02-28")]
		[InlineData("yyyy-mm-dd", 60, "1900-02-29")]
		[InlineData("yyyy-mm-dd", 61, "1900-03-01")]
		[InlineData("[$-F800]dddd, mmmm dd, yyyy", 44927, "Sunday, January 1, 2023")]
		[InlineData("[$-409]d-mmm-yyyy;@", 44927, "1-Jan-2023")]
		[InlineData("dd.mm.yyyy", 44927, "01.01.2023")]
		public void FormatsDates(string format, double value, string expected)
		{
			Assert.Equal(expected, Formatter.Format(value, format));
			Assert.True(Formatter.IsDateFormat(format));
		}

		[Fact]
		public void BuiltInFormats_AreResolved()
		{
			Assert.Equal("2023-01-01", Formatter.Format(44927, NumberFormatter.BuiltInFormat(14)));
			Assert.Equal("1,234.50", Formatter.Format(1234.5, NumberFormatter.BuiltInFormat(4)));
			Assert.Equal("12.00%", Formatter.Format(0.12, NumberFormatter.BuiltInFormat(10)));
			Assert.Equal("36:00:00", Formatter.Format(1.5, NumberFormatter.BuiltInFormat(46)));
			Assert.Null(NumberFormatter.BuiltInFormat(164));
		}

		[Fact]
		public void Date1904_StartsIn1904()
		{
			var formatter = new NumberFormatter(CultureInfo.InvariantCulture, true);

			Assert.Equal("1904-01-01", formatter.Format(0, "yyyy-mm-dd"));
			Assert.Equal("1904-01-02", formatter.Format(1, "yyyy-mm-dd"));
			Assert.Equal(1, formatter.ToSerial(new DateTime(1904, 1, 2)));
		}

		[Fact]
		public void ToSerial_MatchesExcel()
		{
			Assert.Equal(44927, Formatter.ToSerial(new DateTime(2023, 1, 1)));
			Assert.Equal(1, Formatter.ToSerial(new DateTime(1900, 1, 1)));
			Assert.Equal(61, Formatter.ToSerial(new DateTime(1900, 3, 1)));
		}

		[Fact]
		public void NonDateFormats_AreNotDates()
		{
			Assert.False(Formatter.IsDateFormat("0.00E+00"));
			Assert.False(Formatter.IsDateFormat("#,##0.00"));
			Assert.False(Formatter.IsDateFormat("General"));
			Assert.False(Formatter.IsDateFormat("0 \"days\""));
		}

		[Fact]
		public void InvalidDates_UseGeneral()
		{
			Assert.Equal("-5", Formatter.Format(-5, "yyyy-mm-dd"));
		}

		[Fact]
		public void TextSection_FormatsText()
		{
			Assert.Equal("Code: ABC", Formatter.FormatText("ABC", "0;-0;0;\"Code: \"@"));
			Assert.Equal("ABC", Formatter.FormatText("ABC", "0.00"));
		}

		[Fact]
		public void Culture_ChangesSeparatorsAndNames()
		{
			var formatter = new NumberFormatter(CultureInfo.GetCultureInfo("de-DE"), false);

			Assert.Equal("1.234,50", formatter.Format(1234.5, "#,##0.00"));
			Assert.Equal("1. Januar 2023", formatter.Format(44927, "d. mmmm yyyy"));
		}
	}
}
