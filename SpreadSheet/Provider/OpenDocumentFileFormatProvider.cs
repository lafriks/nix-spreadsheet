using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

using Nix.SpreadSheet.Provider.Zip;

namespace Nix.SpreadSheet.Provider
{
	/// <summary>
	/// OpenDocument Spreadsheet file format provider.
	/// </summary>
	public class OpenDocumentFileFormatProvider : IFileFormatProvider
	{
		public Stream CreateManifestStream()
		{
			XmlWriterSettings xws = new XmlWriterSettings();
			xws.Indent = false;

			MemoryStream ms = new MemoryStream();
			XmlWriter xml = XmlWriter.Create(ms, xws);
			xml.WriteStartElement("manifest", "manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0");

			xml.WriteStartElement("manifest", "file-entry");

			xml.WriteStartAttribute("media-type", "manifest");
			xml.WriteString("application/vnd.oasis.opendocument.spreadsheet");
			xml.WriteEndAttribute();

			xml.WriteStartAttribute("version", "manifest");
			xml.WriteString("1.2");
			xml.WriteEndAttribute();

			xml.WriteStartAttribute("full-path", "manifest");
			xml.WriteString("/");
			xml.WriteEndAttribute();

			xml.WriteEndElement();

			xml.WriteEndElement();

			xml.Close();

			ms.Seek(0, SeekOrigin.Begin);

			return ms;
		}
		
		/// <summary>
		/// Save document to stream.
		/// </summary>
		/// <param name="document">Spreadsheet document.</param>
		/// <param name="stream">Stream to write to.</param>
		public void Save(SpreadSheetDocument document, System.IO.Stream stream)
		{
			ZipStream zip = new ZipStream(stream);
			zip.AddFileWithStringContent("mimetype", "application/vnd.oasis.opendocument.spreadsheet");
			using (Stream sm = this.CreateManifestStream())
			{
				zip.AddStream("META-INF/manifest.xml", sm);
				sm.Close();
			}
			zip.Close();
		}
	}
}
