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
using System.Linq;
using System.Text;

namespace Wisej.AI.XlsX.Reading
{
	/// <summary>
	/// Formats cell values with Excel number format codes, so that the extracted text shows what Excel displays.
	/// </summary>
	/// <remarks>
	/// <para>
	/// The formatter implements the number format language: up to four sections (positive, negative,
	/// zero, text), conditions (<c>[&lt;1000]</c>), digit placeholders (<c>0 # ?</c>), thousands
	/// separators and scaling, percentages, scientific notation, fractions, literal text, currency
	/// tags (<c>[$€-407]</c>), dates and times (including elapsed time such as <c>[h]:mm</c> and
	/// fractions of seconds), and the 1900 and 1904 date systems with Excel's 1900 leap year bug.
	/// </para>
	/// <para>
	/// Colors, padding (<c>_x</c>) and fill (<c>*x</c>) affect only the appearance and are ignored.
	/// </para>
	/// </remarks>
	internal sealed class NumberFormatter
	{
		/// <summary>The largest serial number that Excel displays as a date (9999-12-31).</summary>
		private const double MaxDateSerial = 2958466.0;

		private readonly Dictionary<string, FormatCode> _cache = new Dictionary<string, FormatCode>(StringComparer.Ordinal);
		private readonly CultureInfo _culture;
		private readonly bool _date1904;

		/// <summary>
		/// Initializes a new instance of the <see cref="NumberFormatter"/> class.
		/// </summary>
		/// <param name="culture">The culture that provides the separators and the month and day names.</param>
		/// <param name="date1904">Whether the workbook uses the 1904 date system.</param>
		public NumberFormatter(CultureInfo culture, bool date1904)
		{
			_culture = culture ?? CultureInfo.InvariantCulture;
			_date1904 = date1904;
		}

		#region Public API

		/// <summary>
		/// Formats a number.
		/// </summary>
		/// <param name="value">The value.</param>
		/// <param name="formatCode">The format code; <c>null</c> or empty for <c>General</c>.</param>
		/// <returns>The formatted value.</returns>
		public string Format(double value, string formatCode)
		{
			if (Double.IsNaN(value) || Double.IsInfinity(value))
				return "#NUM!";

			return Get(formatCode).Format(value, this);
		}

		/// <summary>
		/// Formats a text value with the text section of a format code, if it has one.
		/// </summary>
		/// <param name="text">The text.</param>
		/// <param name="formatCode">The format code. Can be <c>null</c>.</param>
		/// <returns>The formatted text.</returns>
		public string FormatText(string text, string formatCode)
		{
			return String.IsNullOrEmpty(formatCode) ? text : Get(formatCode).FormatText(text);
		}

		/// <summary>
		/// Determines whether a format code displays numbers as dates or times.
		/// </summary>
		/// <param name="formatCode">The format code. Can be <c>null</c>.</param>
		/// <returns><c>true</c> for a date or time format.</returns>
		public bool IsDateFormat(string formatCode)
		{
			return !String.IsNullOrEmpty(formatCode) && Get(formatCode).IsDate;
		}

		/// <summary>
		/// Converts a date to an Excel serial number in the date system of the workbook.
		/// </summary>
		/// <param name="date">The date.</param>
		/// <returns>The serial number.</returns>
		public double ToSerial(DateTime date)
		{
			if (_date1904)
				return (date - new DateTime(1904, 1, 1)).TotalDays;

			var serial = (date - new DateTime(1899, 12, 30)).TotalDays;

			// Excel counts the nonexistent 1900-02-29 (serial 60).
			return serial < 61 ? serial - 1 : serial;
		}

		/// <summary>
		/// Returns the format code of a built-in number format, or <c>null</c> when the identifier is not built in.
		/// </summary>
		/// <param name="id">The number format identifier (<c>numFmtId</c>).</param>
		/// <returns>The format code or <c>null</c>.</returns>
		/// <remarks>
		/// The short date (14) and date time (22) formats, which Excel displays with the regional settings
		/// of the machine, and the East Asian date formats are mapped to ISO 8601 so that they are never ambiguous.
		/// </remarks>
		public static string BuiltInFormat(int id)
		{
			switch (id)
			{
				case 0: return "General";
				case 1: return "0";
				case 2: return "0.00";
				case 3: return "#,##0";
				case 4: return "#,##0.00";
				case 5: return "\"$\"#,##0_);(\"$\"#,##0)";
				case 6: return "\"$\"#,##0_);[Red](\"$\"#,##0)";
				case 7: return "\"$\"#,##0.00_);(\"$\"#,##0.00)";
				case 8: return "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
				case 9: return "0%";
				case 10: return "0.00%";
				case 11: return "0.00E+00";
				case 12: return "# ?/?";
				case 13: return "# ??/??";
				case 14: return "yyyy-mm-dd";
				case 15: return "d-mmm-yy";
				case 16: return "d-mmm";
				case 17: return "mmm-yy";
				case 18: return "h:mm AM/PM";
				case 19: return "h:mm:ss AM/PM";
				case 20: return "h:mm";
				case 21: return "h:mm:ss";
				case 22: return "yyyy-mm-dd h:mm";
				case 37: return "#,##0_);(#,##0)";
				case 38: return "#,##0_);[Red](#,##0)";
				case 39: return "#,##0.00_);(#,##0.00)";
				case 40: return "#,##0.00_);[Red](#,##0.00)";
				case 41: return "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)";
				case 42: return "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
				case 43: return "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";
				case 44: return "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";
				case 45: return "mm:ss";
				case 46: return "[h]:mm:ss";
				case 47: return "mm:ss.0";
				case 48: return "##0.0E+0";
				case 49: return "@";

				// East Asian built-in date and time formats.
				case 27: case 28: case 29: case 30: case 31: case 34: case 35: case 36:
				case 50: case 51: case 52: case 53: case 54: case 57: case 58:
					return "yyyy-mm-dd";
				case 32: case 33: case 55: case 56:
					return "h:mm:ss";

				default: return null;
			}
		}

		#endregion

		#region Implementation

		/// <summary>
		/// Returns the parsed format code, parsing and caching it on first use.
		/// </summary>
		/// <param name="formatCode">The format code. Can be <c>null</c>.</param>
		/// <returns>The parsed format.</returns>
		private FormatCode Get(string formatCode)
		{
			formatCode = String.IsNullOrEmpty(formatCode) ? "General" : formatCode;

			if (!_cache.TryGetValue(formatCode, out var format))
			{
				format = FormatCode.Parse(formatCode);
				_cache[formatCode] = format;
			}

			return format;
		}

		/// <summary>
		/// Formats a number with the <c>General</c> format: up to 15 significant digits, no grouping.
		/// </summary>
		/// <param name="value">The value.</param>
		/// <returns>The formatted value.</returns>
		private string General(double value)
		{
			if (value == 0)
				return "0";

			if (value == Math.Floor(value) && Math.Abs(value) < 1e15)
				return value.ToString("0", _culture);

			return value.ToString("G15", _culture);
		}

		/// <summary>
		/// Splits a non-negative number, rounded to <paramref name="decimals"/> places, into its integer and fraction digits.
		/// </summary>
		/// <param name="value">The value.</param>
		/// <param name="decimals">The number of decimal places.</param>
		/// <param name="integerDigits">The digits of the integer part, empty for zero.</param>
		/// <param name="fractionDigits">Exactly <paramref name="decimals"/> digits of the fraction.</param>
		/// <remarks>
		/// The value is converted to <see cref="Decimal"/>, which keeps the 15 significant digits that
		/// Excel displays, and rounded half away from zero as Excel does: 2.675 becomes 2.68.
		/// </remarks>
		private static void SplitNumber(double value, int decimals, out string integerDigits, out string fractionDigits)
		{
			value = Math.Abs(value);
			decimals = Math.Min(decimals, 28);

			if (value < 7.9e27)
			{
				var rounded = Math.Round((decimal)value, decimals, MidpointRounding.AwayFromZero);
				var integer = Decimal.Truncate(rounded);

				integerDigits = integer == 0 ? "" : integer.ToString("0", CultureInfo.InvariantCulture);
				fractionDigits = decimals == 0
					? ""
					: (rounded - integer).ToString("0." + new string('0', decimals), CultureInfo.InvariantCulture).Substring(2);
			}
			else
			{
				integerDigits = Math.Round(value).ToString("0", CultureInfo.InvariantCulture);
				fractionDigits = new string('0', decimals);
			}
		}

		/// <summary>
		/// Formats a non-negative number with a number section that has no exponent or fraction.
		/// </summary>
		/// <param name="tokens">The tokens of the section.</param>
		/// <param name="start">The index of the first token.</param>
		/// <param name="end">The index after the last token.</param>
		/// <param name="value">The value, already scaled.</param>
		/// <param name="grouping">Whether digits are grouped by thousands.</param>
		/// <returns>The formatted number.</returns>
		private string FormatFixed(List<Token> tokens, int start, int end, double value, bool grouping)
		{
			var decimalIndex = -1;
			for (var i = start; i < end; i++)
			{
				if (tokens[i].Kind == TokenKind.DecimalPoint)
				{
					decimalIndex = i;
					break;
				}
			}

			var integerEnd = decimalIndex < 0 ? end : decimalIndex;
			var decimals = 0;
			for (var i = integerEnd; i < end; i++)
			{
				if (tokens[i].Kind == TokenKind.Digit)
					decimals++;
			}

			SplitNumber(value, decimals, out var integerDigits, out var fractionDigits);

			var sb = new StringBuilder();
			sb.Append(FormatInteger(tokens, start, integerEnd, integerDigits, grouping));

			if (decimalIndex >= 0)
			{
				sb.Append(_culture.NumberFormat.NumberDecimalSeparator);

				// Optional placeholders (# and ?) drop the trailing zeros of the fraction.
				var placeholders = new List<Token>();
				for (var i = decimalIndex + 1; i < end; i++)
				{
					if (tokens[i].Kind == TokenKind.Digit)
						placeholders.Add(tokens[i]);
				}

				var visible = placeholders.Count;
				while (visible > 0 && placeholders[visible - 1].Placeholder != '0' && fractionDigits[visible - 1] == '0')
					visible--;

				var digit = 0;
				for (var i = decimalIndex + 1; i < end; i++)
				{
					var token = tokens[i];
					if (token.Kind == TokenKind.Digit)
					{
						if (digit < visible)
							sb.Append(fractionDigits[digit]);
						else if (token.Placeholder == '?')
							sb.Append(' ');

						digit++;
					}
					else
					{
						sb.Append(LiteralText(token));
					}
				}
			}

			return sb.ToString();
		}

		/// <summary>
		/// Formats the integer part of a number: digits fill the placeholders from the right, extra
		/// digits go to the leftmost placeholder, and literals stay where they are written.
		/// </summary>
		/// <param name="tokens">The tokens of the section.</param>
		/// <param name="start">The index of the first token of the integer part.</param>
		/// <param name="end">The index after the last token of the integer part.</param>
		/// <param name="digits">The integer digits, empty for zero.</param>
		/// <param name="grouping">Whether digits are grouped by thousands.</param>
		/// <returns>The formatted integer part.</returns>
		private string FormatInteger(List<Token> tokens, int start, int end, string digits, bool grouping)
		{
			var placeholders = 0;
			for (var i = start; i < end; i++)
			{
				if (tokens[i].Kind == TokenKind.Digit)
					placeholders++;
			}

			var sb = new StringBuilder();
			var separator = _culture.NumberFormat.NumberGroupSeparator;
			var next = digits.Length - 1;
			var emitted = 0;
			var seen = 0;

			void Emit(char digit)
			{
				if (grouping && emitted > 0 && emitted % 3 == 0)
					sb.Insert(0, separator);

				sb.Insert(0, digit);
				emitted++;
			}

			void Pad(char placeholder)
			{
				if (placeholder == '0')
					Emit('0');
				else if (placeholder == '?')
					sb.Insert(0, ' ');
			}

			// Without placeholders (".00"), the integer digits are written before the decimal point.
			if (placeholders == 0)
			{
				while (next >= 0)
					Emit(digits[next--]);
			}

			for (var i = end - 1; i >= start; i--)
			{
				var token = tokens[i];
				if (token.Kind == TokenKind.Digit)
				{
					seen++;
					if (seen == placeholders)
					{
						if (next < 0)
							Pad(token.Placeholder);

						while (next >= 0)
							Emit(digits[next--]);
					}
					else if (next >= 0)
					{
						Emit(digits[next--]);
					}
					else
					{
						Pad(token.Placeholder);
					}
				}
				else
				{
					sb.Insert(0, LiteralText(token));
				}
			}

			return sb.ToString();
		}

		/// <summary>
		/// Formats a non-negative number in scientific notation.
		/// </summary>
		/// <param name="section">The section.</param>
		/// <param name="value">The value, already scaled.</param>
		/// <returns>The formatted number.</returns>
		private string FormatScientific(Section section, double value)
		{
			var tokens = section.Tokens;
			var exponentIndex = section.ExponentIndex;

			var integerPlaceholders = 0;
			var optional = false;
			var decimals = 0;
			var inFraction = false;

			for (var i = 0; i < exponentIndex; i++)
			{
				var token = tokens[i];
				if (token.Kind == TokenKind.DecimalPoint)
					inFraction = true;
				else if (token.Kind == TokenKind.Digit && inFraction)
					decimals++;
				else if (token.Kind == TokenKind.Digit)
				{
					integerPlaceholders++;
					optional |= token.Placeholder == '#';
				}
			}

			integerPlaceholders = Math.Max(1, integerPlaceholders);

			// ##0.0E+0 is engineering notation: the exponent is a multiple of the number of integer placeholders.
			int Exponent(double v)
			{
				var e = v == 0 ? 0 : (int)Math.Floor(Math.Log10(v));
				return optional && integerPlaceholders > 1
					? (int)Math.Floor(e / (double)integerPlaceholders) * integerPlaceholders
					: e - (integerPlaceholders - 1);
			}

			var exponent = Exponent(value);
			var mantissa = value == 0 ? 0 : value / Math.Pow(10, exponent);

			// Rounding can carry the mantissa to the next power of ten (9.99 becomes 10.0).
			var rounded = Math.Round(mantissa, Math.Min(decimals, 15), MidpointRounding.AwayFromZero);
			if (value != 0 && rounded >= Math.Pow(10, integerPlaceholders))
			{
				exponent = Exponent(rounded * Math.Pow(10, exponent));
				mantissa = value / Math.Pow(10, exponent);
			}

			var sb = new StringBuilder(FormatFixed(tokens, 0, exponentIndex, mantissa, false));

			var marker = tokens[exponentIndex];
			sb.Append(marker.Text);
			if (exponent < 0)
				sb.Append('-');
			else if (marker.ExponentPlus)
				sb.Append('+');

			var minimum = 0;
			for (var i = exponentIndex + 1; i < tokens.Count; i++)
			{
				if (tokens[i].Kind == TokenKind.Digit && tokens[i].Placeholder == '0')
					minimum++;
			}

			sb.Append(Math.Abs(exponent).ToString(CultureInfo.InvariantCulture).PadLeft(minimum, '0'));

			for (var i = exponentIndex + 1; i < tokens.Count; i++)
			{
				if (tokens[i].Kind != TokenKind.Digit)
					sb.Append(LiteralText(tokens[i]));
			}

			return sb.ToString();
		}

		/// <summary>
		/// Formats a non-negative number as a fraction, such as <c>1 3/4</c> or <c>7/4</c>.
		/// </summary>
		/// <param name="section">The section.</param>
		/// <param name="value">The value, already scaled.</param>
		/// <returns>The formatted number.</returns>
		private string FormatFraction(Section section, double value)
		{
			var tokens = section.Tokens;
			var slash = section.SlashIndex;

			// The numerator is the run of placeholders before the slash.
			var numeratorStart = slash;
			while (numeratorStart > 0 && tokens[numeratorStart - 1].Kind == TokenKind.Digit)
				numeratorStart--;

			var integerTokens = tokens.Take(numeratorStart).ToList();
			var hasInteger = integerTokens.Any(t => t.Kind == TokenKind.Digit);

			// The denominator is a run of placeholders or a fixed number.
			var denominatorDigits = 0;
			var fixedDenominator = new StringBuilder();
			var index = slash + 1;
			while (index < tokens.Count && tokens[index].Kind == TokenKind.Digit)
			{
				denominatorDigits++;
				index++;
			}

			if (denominatorDigits == 0)
			{
				while (index < tokens.Count && tokens[index].Kind == TokenKind.Literal && tokens[index].Text.Length > 0 && tokens[index].Text.All(Char.IsDigit))
					fixedDenominator.Append(tokens[index++].Text);
			}

			var whole = hasInteger ? Math.Floor(value) : 0;
			var fraction = value - whole;

			long numerator, denominator;
			if (fixedDenominator.Length > 0 && Int64.TryParse(fixedDenominator.ToString(), out denominator) && denominator > 0)
			{
				numerator = (long)Math.Round(fraction * denominator, MidpointRounding.AwayFromZero);
			}
			else
			{
				var maxDenominator = (long)Math.Min(Math.Pow(10, Math.Max(1, denominatorDigits)) - 1, 9999);
				BestFraction(fraction, maxDenominator, out numerator, out denominator);
			}

			if (hasInteger && numerator == denominator)
			{
				whole++;
				numerator = 0;
			}

			var sb = new StringBuilder();

			// Literals before the first placeholder, such as a currency symbol.
			foreach (var token in tokens)
			{
				if (token.Kind == TokenKind.Digit)
					break;
				sb.Append(LiteralText(token));
			}

			var integerText = whole > 0 ? whole.ToString("0", _culture) : "";
			if (integerText.Length == 0 && hasInteger && integerTokens.Any(t => t.Kind == TokenKind.Digit && t.Placeholder == '0'))
				integerText = "0";

			if (numerator == 0 && hasInteger)
			{
				sb.Append(integerText.Length > 0 ? integerText : "0");
			}
			else
			{
				if (integerText.Length > 0)
					sb.Append(integerText).Append(' ');

				sb.Append(numerator.ToString(CultureInfo.InvariantCulture)).Append('/').Append(denominator.ToString(CultureInfo.InvariantCulture));
			}

			for (var i = index; i < tokens.Count; i++)
				sb.Append(LiteralText(tokens[i]));

			return sb.ToString();
		}

		/// <summary>
		/// Finds the fraction closest to a value with a denominator not larger than <paramref name="maxDenominator"/>.
		/// </summary>
		/// <param name="value">The value.</param>
		/// <param name="maxDenominator">The largest denominator.</param>
		/// <param name="numerator">The numerator.</param>
		/// <param name="denominator">The denominator.</param>
		private static void BestFraction(double value, long maxDenominator, out long numerator, out long denominator)
		{
			numerator = (long)Math.Round(value, MidpointRounding.AwayFromZero);
			denominator = 1;
			var best = Math.Abs(value - numerator);

			for (long d = 2; d <= maxDenominator && best > 1e-12; d++)
			{
				var n = (long)Math.Round(value * d, MidpointRounding.AwayFromZero);
				var error = Math.Abs(value - (double)n / d);
				if (error < best - 1e-15)
				{
					best = error;
					numerator = n;
					denominator = d;
				}
			}
		}

		/// <summary>
		/// Formats a serial number as a date or time.
		/// </summary>
		/// <param name="section">The date section.</param>
		/// <param name="serial">The serial number.</param>
		/// <returns>The formatted date, or <c>null</c> when the serial number is not a valid date.</returns>
		private string FormatDate(Section section, double serial)
		{
			if (serial < 0 || serial >= MaxDateSerial)
				return null;

			// Round to the smallest unit displayed: Excel rounds 10:29:59.6 to 10:30:00 when seconds are shown.
			var subSecondDigits = section.Tokens.Where(t => t.Kind == TokenKind.SubSecond).Select(t => t.Count).DefaultIfEmpty(0).Max();
			var showsSeconds = subSecondDigits > 0 || section.Tokens.Any(t => (t.Kind == TokenKind.DatePart || t.Kind == TokenKind.Elapsed) && t.Letter == 's');

			var totalMs = (long)Math.Round(serial * 86400000.0, MidpointRounding.AwayFromZero);
			if (showsSeconds)
			{
				var unit = (long)Math.Pow(10, 3 - Math.Min(3, subSecondDigits));
				totalMs = (long)Math.Round(totalMs / (double)unit, MidpointRounding.AwayFromZero) * unit;
			}

			var days = (int)(totalMs / 86400000);
			var msOfDay = totalMs % 86400000;
			var hour = (int)(msOfDay / 3600000);
			var minute = (int)(msOfDay / 60000 % 60);
			var second = (int)(msOfDay / 1000 % 60);
			var millisecond = (int)(msOfDay % 1000);

			GetDate(days, out var year, out var month, out var day, out var dayOfWeek);

			var names = _culture.DateTimeFormat;
			var sb = new StringBuilder();

			foreach (var token in section.Tokens)
			{
				switch (token.Kind)
				{
					case TokenKind.DatePart:
						switch (token.Letter)
						{
							case 'y':
								sb.Append(token.Count <= 2 ? (year % 100).ToString("00", CultureInfo.InvariantCulture) : year.ToString("0000", CultureInfo.InvariantCulture));
								break;

							case 'm' when token.IsMinute:
								sb.Append(minute.ToString(token.Count == 1 ? "0" : "00", CultureInfo.InvariantCulture));
								break;

							case 'm':
								if (token.Count == 1)
									sb.Append(month.ToString(CultureInfo.InvariantCulture));
								else if (token.Count == 2)
									sb.Append(month.ToString("00", CultureInfo.InvariantCulture));
								else if (token.Count == 3)
									sb.Append(names.AbbreviatedMonthNames[month - 1]);
								else if (token.Count == 4)
									sb.Append(names.MonthNames[month - 1]);
								else
									sb.Append(names.MonthNames[month - 1].Substring(0, 1));
								break;

							case 'd':
								if (token.Count == 1)
									sb.Append(day.ToString(CultureInfo.InvariantCulture));
								else if (token.Count == 2)
									sb.Append(day.ToString("00", CultureInfo.InvariantCulture));
								else if (token.Count == 3)
									sb.Append(names.AbbreviatedDayNames[(int)dayOfWeek]);
								else
									sb.Append(names.DayNames[(int)dayOfWeek]);
								break;

							case 'h':
								var h = section.Hour12 ? (hour % 12 == 0 ? 12 : hour % 12) : hour;
								sb.Append(h.ToString(token.Count == 1 ? "0" : "00", CultureInfo.InvariantCulture));
								break;

							case 's':
								sb.Append(second.ToString(token.Count == 1 ? "0" : "00", CultureInfo.InvariantCulture));
								break;
						}
						break;

					case TokenKind.SubSecond:
						var digits = millisecond.ToString("000", CultureInfo.InvariantCulture).Substring(0, Math.Min(3, token.Count));
						sb.Append(_culture.NumberFormat.NumberDecimalSeparator).Append(digits.PadRight(token.Count, '0'));
						break;

					case TokenKind.AmPm:
						var pm = hour >= 12;
						if (token.Text.Length == 3)
						{
							var letter = pm ? token.Text[2] : token.Text[0];
							sb.Append(letter);
						}
						else
						{
							var text = pm ? "PM" : "AM";
							sb.Append(Char.IsLower(token.Text[0]) ? text.ToLowerInvariant() : text);
						}
						break;

					case TokenKind.Elapsed:
						var elapsed = token.Letter == 'h' ? totalMs / 3600000 : token.Letter == 'm' ? totalMs / 60000 : totalMs / 1000;
						sb.Append(elapsed.ToString(CultureInfo.InvariantCulture).PadLeft(token.Count, '0'));
						break;

					default:
						sb.Append(LiteralText(token));
						break;
				}
			}

			return sb.ToString();
		}

		/// <summary>
		/// Converts a day serial number to its date, in the date system of the workbook.
		/// </summary>
		/// <param name="days">The integer part of the serial number.</param>
		/// <param name="year">The year.</param>
		/// <param name="month">The month.</param>
		/// <param name="day">The day; zero for the serial number zero of the 1900 system.</param>
		/// <param name="dayOfWeek">The day of the week.</param>
		private void GetDate(int days, out int year, out int month, out int day, out DayOfWeek dayOfWeek)
		{
			DateTime date;

			if (_date1904)
			{
				date = new DateTime(1904, 1, 1).AddDays(days);
			}
			else if (days == 0)
			{
				// Excel displays the serial number zero as January 0, 1900.
				year = 1900;
				month = 1;
				day = 0;
				dayOfWeek = DayOfWeek.Saturday;
				return;
			}
			else if (days == 60)
			{
				// The nonexistent February 29, 1900, kept by Excel for compatibility with Lotus 1-2-3.
				year = 1900;
				month = 2;
				day = 29;
				dayOfWeek = DayOfWeek.Wednesday;
				return;
			}
			else
			{
				date = days < 60 ? new DateTime(1899, 12, 31).AddDays(days) : new DateTime(1899, 12, 30).AddDays(days);
			}

			year = date.Year;
			month = date.Month;
			day = date.Day;
			dayOfWeek = date.DayOfWeek;
		}

		/// <summary>
		/// Returns the text that a non-placeholder token contributes to the output.
		/// </summary>
		/// <param name="token">The token.</param>
		/// <returns>The text.</returns>
		private string LiteralText(Token token)
		{
			switch (token.Kind)
			{
				case TokenKind.Literal: return token.Text;
				case TokenKind.Percent: return "%";
				case TokenKind.Comma: return ",";
				case TokenKind.Slash: return "/";
				case TokenKind.DecimalPoint: return ".";
				default: return "";
			}
		}

		#endregion

		#region Format Code

		/// <summary>
		/// The kinds of tokens of a format section.
		/// </summary>
		private enum TokenKind
		{
			Literal,
			Digit,
			DecimalPoint,
			Comma,
			Percent,
			Exponent,
			Slash,
			Text,
			General,
			DatePart,
			AmPm,
			Elapsed,
			SubSecond,
		}

		/// <summary>
		/// A token of a format section.
		/// </summary>
		private sealed class Token
		{
			public TokenKind Kind;
			public string Text = "";
			public char Placeholder;
			public char Letter;
			public int Count;
			public bool ExponentPlus;
			public bool IsMinute;

			public static Token Literal(string text) => new Token { Kind = TokenKind.Literal, Text = text };
		}

		/// <summary>
		/// A section of a format code.
		/// </summary>
		private sealed class Section
		{
			public readonly List<Token> Tokens = new List<Token>();
			public bool IsDate;
			public bool Hour12;
			public bool HasText;
			public bool HasDigits;
			public bool HasGeneral;
			public bool Grouping;
			public int PercentCount;
			public int ScaleCount;
			public int ExponentIndex = -1;
			public int SlashIndex = -1;
			public string ConditionOperator;
			public double ConditionValue;

			/// <summary>
			/// Determines whether the condition of the section accepts a value.
			/// </summary>
			public bool Matches(double value)
			{
				switch (ConditionOperator)
				{
					case "<": return value < ConditionValue;
					case "<=": return value <= ConditionValue;
					case ">": return value > ConditionValue;
					case ">=": return value >= ConditionValue;
					case "=": return value == ConditionValue;
					case "<>": return value != ConditionValue;
					default: return true;
				}
			}

			/// <summary>
			/// Gets whether the condition accepts only negative numbers, which are then written without a minus sign.
			/// </summary>
			public bool IsNegativeCondition => (ConditionOperator == "<" && ConditionValue <= 0) || (ConditionOperator == "<=" && ConditionValue < 0);
		}

		/// <summary>
		/// A parsed format code.
		/// </summary>
		private sealed class FormatCode
		{
			private readonly List<Section> _sections = new List<Section>();

			/// <summary>
			/// Gets whether the format displays numbers as dates or times.
			/// </summary>
			public bool IsDate => _sections.Count > 0 && _sections[0].IsDate;

			/// <summary>
			/// Parses a format code.
			/// </summary>
			public static FormatCode Parse(string code)
			{
				var format = new FormatCode();
				foreach (var text in SplitSections(code))
					format._sections.Add(ParseSection(text));

				return format;
			}

			/// <summary>
			/// Formats a number.
			/// </summary>
			public string Format(double value, NumberFormatter formatter)
			{
				if (!SelectSection(value, out var section, out var minus))
					return formatter.General(value);

				var magnitude = Math.Abs(value);
				string text;

				if (section.Tokens.Count == 0)
				{
					return "";
				}
				else if (section.IsDate)
				{
					text = minus ? null : formatter.FormatDate(section, magnitude);
					if (text == null)
						return formatter.General(value);

					return text;
				}
				else if (section.HasGeneral || (!section.HasDigits && !section.HasText))
				{
					var sb = new StringBuilder();
					foreach (var token in section.Tokens)
						sb.Append(token.Kind == TokenKind.General ? formatter.General(magnitude) : formatter.LiteralText(token));

					// A section of literals only ("Zero") is written as is.
					if (!section.HasGeneral && !section.HasDigits)
						return sb.ToString();

					text = sb.ToString();
				}
				else if (!section.HasDigits)
				{
					// A text-only section ("@") applied to a number.
					return formatter.General(value);
				}
				else
				{
					var scaled = magnitude * Math.Pow(100, section.PercentCount) / Math.Pow(1000, section.ScaleCount);

					if (section.SlashIndex >= 0)
						text = formatter.FormatFraction(section, scaled);
					else if (section.ExponentIndex >= 0)
						text = formatter.FormatScientific(section, scaled);
					else
						text = formatter.FormatFixed(section.Tokens, 0, section.Tokens.Count, scaled, section.Grouping);
				}

				return minus ? "-" + text : text;
			}

			/// <summary>
			/// Formats a text value with the text section, if the format has one.
			/// </summary>
			public string FormatText(string text)
			{
				Section section = null;
				if (_sections.Count >= 4)
					section = _sections[3];
				else if (_sections.Count == 1 && _sections[0].HasText)
					section = _sections[0];

				if (section == null)
					return text;

				var sb = new StringBuilder();
				foreach (var token in section.Tokens)
				{
					if (token.Kind == TokenKind.Text)
						sb.Append(text);
					else if (token.Kind == TokenKind.Literal)
						sb.Append(token.Text);
				}

				return sb.ToString();
			}

			/// <summary>
			/// Chooses the section that formats a number.
			/// </summary>
			/// <returns><c>false</c> when no section applies and the number is written with the General format.</returns>
			private bool SelectSection(double value, out Section section, out bool minus)
			{
				section = null;
				minus = false;

				// The fourth section formats text only.
				var numeric = _sections.Take(3).ToList();
				if (numeric.Count == 0)
					return false;

				if (numeric.Any(s => s.ConditionOperator != null))
				{
					foreach (var candidate in numeric)
					{
						if (candidate.Matches(value))
						{
							section = candidate;
							minus = value < 0 && !candidate.IsNegativeCondition;
							return true;
						}
					}

					return false;
				}

				if (numeric.Count == 1)
				{
					section = numeric[0];
					minus = value < 0;
				}
				else if (value > 0 || (value == 0 && numeric.Count == 2))
				{
					section = numeric[0];
				}
				else if (value < 0)
				{
					section = numeric[1];
				}
				else
				{
					section = numeric[2];
				}

				return true;
			}

			/// <summary>
			/// Splits a format code into its sections, honoring quotes, escapes and brackets.
			/// </summary>
			private static List<string> SplitSections(string code)
			{
				var sections = new List<string>();
				var start = 0;

				for (var i = 0; i < code.Length; i++)
				{
					var c = code[i];
					if (c == '"')
					{
						var close = code.IndexOf('"', i + 1);
						i = close < 0 ? code.Length : close;
					}
					else if (c == '\\' || c == '_' || c == '*')
					{
						i++;
					}
					else if (c == '[')
					{
						var close = code.IndexOf(']', i + 1);
						i = close < 0 ? code.Length : close;
					}
					else if (c == ';')
					{
						sections.Add(code.Substring(start, i - start));
						start = i + 1;
					}
				}

				sections.Add(code.Substring(Math.Min(start, code.Length)));
				return sections;
			}

			/// <summary>
			/// Parses a section into tokens and analyzes it.
			/// </summary>
			private static Section ParseSection(string text)
			{
				var section = new Section();
				var tokens = section.Tokens;
				string systemFormat = null;

				for (var i = 0; i < text.Length; i++)
				{
					var c = text[i];
					switch (c)
					{
						case '"':
							var close = text.IndexOf('"', i + 1);
							if (close < 0)
								close = text.Length;
							tokens.Add(Token.Literal(text.Substring(i + 1, close - i - 1)));
							i = close;
							break;

						case '\\':
							if (i + 1 < text.Length)
								tokens.Add(Token.Literal(text[++i].ToString()));
							break;

						case '_':
						case '*':
							// Padding and fill characters only align the text in the cell.
							i++;
							break;

						case '[':
							var end = text.IndexOf(']', i + 1);
							if (end < 0)
								end = text.Length;
							systemFormat = ParseBracket(section, text.Substring(i + 1, end - i - 1)) ?? systemFormat;
							i = end;
							break;

						case '0':
						case '#':
						case '?':
							tokens.Add(new Token { Kind = TokenKind.Digit, Placeholder = c });
							break;

						case '.':
							var last = tokens.LastOrDefault(t => t.Kind == TokenKind.DatePart || t.Kind == TokenKind.Elapsed);
							if (last != null && last.Letter == 's' && i + 1 < text.Length && text[i + 1] == '0')
							{
								var count = 0;
								while (i + 1 < text.Length && text[i + 1] == '0')
								{
									count++;
									i++;
								}
								tokens.Add(new Token { Kind = TokenKind.SubSecond, Count = count });
							}
							else
							{
								tokens.Add(new Token { Kind = TokenKind.DecimalPoint });
							}
							break;

						case ',':
							tokens.Add(new Token { Kind = TokenKind.Comma });
							break;

						case '%':
							tokens.Add(new Token { Kind = TokenKind.Percent });
							break;

						case '/':
							tokens.Add(new Token { Kind = TokenKind.Slash });
							break;

						case '@':
							tokens.Add(new Token { Kind = TokenKind.Text });
							break;

						case 'E':
						case 'e':
							if (i + 1 < text.Length && (text[i + 1] == '+' || text[i + 1] == '-'))
							{
								tokens.Add(new Token { Kind = TokenKind.Exponent, Text = c.ToString(), ExponentPlus = text[i + 1] == '+' });
								i++;
							}
							else if (c == 'e')
							{
								// "e" is the year in the era calendars.
								tokens.Add(new Token { Kind = TokenKind.DatePart, Letter = 'y', Count = 4 });
							}
							else
							{
								tokens.Add(Token.Literal("E"));
							}
							break;

						default:
							if (String.Compare(text, i, "General", 0, 7, StringComparison.OrdinalIgnoreCase) == 0)
							{
								tokens.Add(new Token { Kind = TokenKind.General });
								i += 6;
							}
							else if (String.Compare(text, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
							{
								tokens.Add(new Token { Kind = TokenKind.AmPm, Text = text.Substring(i, 5) });
								i += 4;
							}
							else if (String.Compare(text, i, "A/P", 0, 3, StringComparison.OrdinalIgnoreCase) == 0)
							{
								tokens.Add(new Token { Kind = TokenKind.AmPm, Text = text.Substring(i, 3) });
								i += 2;
							}
							else if ("yYmMdDhHsS".IndexOf(c) >= 0)
							{
								var letter = Char.ToLowerInvariant(c);
								var count = 1;
								while (i + 1 < text.Length && Char.ToLowerInvariant(text[i + 1]) == letter)
								{
									count++;
									i++;
								}
								tokens.Add(new Token { Kind = TokenKind.DatePart, Letter = letter, Count = count });
							}
							else if (Char.IsDigit(c) && tokens.Count > 0 && tokens[tokens.Count - 1].Kind == TokenKind.Slash)
							{
								// A fixed denominator ("# ?/16"): its digits, zeros included, are not placeholders.
								var start = i;
								while (i + 1 < text.Length && Char.IsDigit(text[i + 1]))
									i++;
								tokens.Add(Token.Literal(text.Substring(start, i - start + 1)));
							}
							else if ((c == 'b' || c == 'B' || c == 'g' || c == 'G') && i + 1 < text.Length && Char.IsDigit(text[i + 1]))
							{
								// Calendar selectors (B1, B2) do not produce text.
								i++;
							}
							else
							{
								tokens.Add(Token.Literal(c.ToString()));
							}
							break;
					}
				}

				// [$-F800] and [$-F400] stand for the long date and the time format of the system.
				if (systemFormat != null)
				{
					var replacement = ParseSection(systemFormat);
					tokens.Clear();
					tokens.AddRange(replacement.Tokens);
				}

				Analyze(section);
				return section;
			}

			/// <summary>
			/// Interprets the content of a bracket: a condition, an elapsed time, a currency or locale tag, or a color.
			/// </summary>
			/// <returns>The system format that replaces the section, or <c>null</c>.</returns>
			private static string ParseBracket(Section section, string content)
			{
				if (content.Length == 0)
					return null;

				if (content[0] == '$')
				{
					var dash = content.IndexOf('-');
					var symbol = dash < 0 ? content.Substring(1) : content.Substring(1, dash - 1);
					if (symbol.Length > 0)
						section.Tokens.Add(Token.Literal(symbol));

					var locale = dash < 0 ? "" : content.Substring(dash + 1).ToUpperInvariant();
					if (locale.EndsWith("F800", StringComparison.Ordinal))
						return "dddd, mmmm d, yyyy";
					if (locale.EndsWith("F400", StringComparison.Ordinal))
						return "h:mm:ss AM/PM";

					return null;
				}

				if (content[0] == '<' || content[0] == '>' || content[0] == '=')
				{
					var length = content.StartsWith("<=") || content.StartsWith(">=") || content.StartsWith("<>") ? 2 : 1;
					if (Double.TryParse(content.Substring(length), NumberStyles.Float, CultureInfo.InvariantCulture, out var operand))
					{
						section.ConditionOperator = content.Substring(0, length);
						section.ConditionValue = operand;
					}

					return null;
				}

				var letter = Char.ToLowerInvariant(content[0]);
				if ((letter == 'h' || letter == 'm' || letter == 's') && content.All(ch => Char.ToLowerInvariant(ch) == letter))
					section.Tokens.Add(new Token { Kind = TokenKind.Elapsed, Letter = letter, Count = content.Length });

				// Colors ([Red], [Color10]) and other tags do not produce text.
				return null;
			}

			/// <summary>
			/// Analyzes the tokens of a section: date or number, separators, scaling, exponent and fraction.
			/// </summary>
			private static void Analyze(Section section)
			{
				var tokens = section.Tokens;

				section.IsDate = tokens.Any(t => t.Kind == TokenKind.DatePart || t.Kind == TokenKind.AmPm || t.Kind == TokenKind.Elapsed);
				section.HasText = tokens.Any(t => t.Kind == TokenKind.Text);
				section.HasGeneral = tokens.Any(t => t.Kind == TokenKind.General);

				if (section.IsDate)
				{
					section.Hour12 = tokens.Any(t => t.Kind == TokenKind.AmPm);
					ResolveMinutes(tokens);

					for (var i = 0; i < tokens.Count; i++)
					{
						var kind = tokens[i].Kind;
						if (kind == TokenKind.Digit)
							tokens[i] = Token.Literal(tokens[i].Placeholder.ToString());
						else if (kind == TokenKind.Exponent || kind == TokenKind.General || kind == TokenKind.Text)
							tokens[i] = Token.Literal("");
					}

					return;
				}

				section.HasDigits = tokens.Any(t => t.Kind == TokenKind.Digit);
				section.PercentCount = tokens.Count(t => t.Kind == TokenKind.Percent);
				section.ExponentIndex = tokens.FindIndex(t => t.Kind == TokenKind.Exponent);

				var slash = tokens.FindIndex(t => t.Kind == TokenKind.Slash);
				if (slash > 0 && section.ExponentIndex < 0 && tokens[slash - 1].Kind == TokenKind.Digit
					&& slash + 1 < tokens.Count && (tokens[slash + 1].Kind == TokenKind.Digit || (tokens[slash + 1].Kind == TokenKind.Literal && tokens[slash + 1].Text.All(Char.IsDigit))))
				{
					section.SlashIndex = slash;
				}

				var integerEnd = tokens.FindIndex(t => t.Kind == TokenKind.DecimalPoint);
				if (integerEnd < 0)
					integerEnd = section.ExponentIndex >= 0 ? section.ExponentIndex : tokens.Count;

				// A comma between digit placeholders groups thousands; commas after the last placeholder divide by 1000.
				for (var i = 0; i < tokens.Count; i++)
				{
					if (tokens[i].Kind != TokenKind.Comma)
						continue;

					var digitBefore = tokens.Take(i).Any(t => t.Kind == TokenKind.Digit);
					var digitAfter = tokens.Skip(i + 1).Any(t => t.Kind == TokenKind.Digit);

					if (digitBefore && !digitAfter)
					{
						section.ScaleCount++;
						tokens[i] = Token.Literal("");
					}
					else if (digitBefore && i < integerEnd)
					{
						section.Grouping = true;
						tokens[i] = Token.Literal("");
					}
					else
					{
						tokens[i] = Token.Literal(",");
					}
				}
			}

			/// <summary>
			/// Marks the <c>m</c> tokens that mean minutes: after an hour or before a second.
			/// </summary>
			private static void ResolveMinutes(List<Token> tokens)
			{
				var dateTokens = tokens.Where(t => t.Kind == TokenKind.DatePart || t.Kind == TokenKind.Elapsed).ToList();

				for (var i = 0; i < dateTokens.Count; i++)
				{
					var token = dateTokens[i];
					if (token.Kind != TokenKind.DatePart || token.Letter != 'm' || token.Count > 2)
						continue;

					var previous = i > 0 ? dateTokens[i - 1] : null;
					var next = i + 1 < dateTokens.Count ? dateTokens[i + 1] : null;

					token.IsMinute = (previous != null && previous.Letter == 'h') || (next != null && next.Letter == 's');
				}
			}
		}

		#endregion
	}
}
