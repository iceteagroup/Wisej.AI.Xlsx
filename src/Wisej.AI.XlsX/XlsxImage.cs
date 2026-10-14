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

namespace Wisej.AI.XlsX
{
	/// <summary>
	/// An image placed in the workbook: a picture floating over a sheet or a picture placed in a cell.
	/// </summary>
	/// <remarks>
	/// <para>
	/// Every placement of a picture produces a <see cref="XlsxImage"/>. When the same picture is used
	/// more than once the placements share the same <see cref="FileName"/> and <see cref="Data"/>.
	/// </para>
	/// <para>
	/// The text output references the image by <see cref="FileName"/> (<c>![alt text](image1.png)</c>
	/// in Markdown), so a caller can save <see cref="Data"/> under that name, describe it with a vision
	/// model, or run OCR on it through <see cref="XlsxExtractorOptions.ImageDescriber"/>.
	/// </para>
	/// </remarks>
	public sealed class XlsxImage
	{
		/// <summary>
		/// Gets the file name assigned to the image, such as <c>image3.png</c>.
		/// </summary>
		public string FileName { get; internal set; }

		/// <summary>
		/// Gets the MIME type of the image data, such as <c>image/png</c> or <c>image/x-emf</c>.
		/// </summary>
		public string ContentType { get; internal set; }

		/// <summary>
		/// Gets the image bytes, or <c>null</c> when the image is linked to an external file or
		/// <see cref="XlsxExtractorOptions.IncludeImageData"/> is <c>false</c>.
		/// </summary>
		public byte[] Data { get; internal set; }

		/// <summary>
		/// Gets the target of an image linked to an external file instead of embedded, or <c>null</c>.
		/// </summary>
		public string ExternalUri { get; internal set; }

		/// <summary>
		/// Gets the alternative text (description) written by the author for accessibility, or <c>null</c>.
		/// </summary>
		public string AltText { get; internal set; }

		/// <summary>
		/// Gets the title of the image, or <c>null</c>.
		/// </summary>
		public string Title { get; internal set; }

		/// <summary>
		/// Gets the name of the drawing object, such as <c>Picture 3</c>, or <c>null</c> for a picture in a cell.
		/// </summary>
		public string Name { get; internal set; }

		/// <summary>
		/// Gets the displayed width of the image, in points, or <c>0</c> when it is not known.
		/// </summary>
		public double Width { get; internal set; }

		/// <summary>
		/// Gets the displayed height of the image, in points, or <c>0</c> when it is not known.
		/// </summary>
		public double Height { get; internal set; }

		/// <summary>
		/// Gets the name of the sheet that displays the image.
		/// </summary>
		public string SheetName { get; internal set; }

		/// <summary>
		/// Gets the address of the cell that contains the image, or under the top left corner of a
		/// floating picture, such as <c>D5</c>; <c>null</c> when the picture is not anchored to a cell.
		/// </summary>
		public string Cell { get; internal set; }

		/// <summary>
		/// Gets whether the picture is placed in a cell (Excel's Place in Cell) instead of floating over the sheet.
		/// </summary>
		public bool IsInCell { get; internal set; }

		/// <summary>
		/// Gets the number of the page that displays the image.
		/// </summary>
		public int PageNumber { get; internal set; }

		/// <summary>
		/// Gets the description of the image produced by <see cref="XlsxExtractorOptions.ImageDescriber"/>, or <c>null</c>.
		/// </summary>
		public string Description { get; internal set; }

		/// <summary>
		/// Gets the internal URI of the image part, which identifies the picture inside the package.
		/// </summary>
		internal string PartUri { get; set; }

		/// <summary>
		/// Returns the file name and the alternative text of the image.
		/// </summary>
		/// <returns>A string describing the image.</returns>
		public override string ToString()
		{
			return AltText == null ? FileName : $"{FileName} ({AltText})";
		}
	}
}
