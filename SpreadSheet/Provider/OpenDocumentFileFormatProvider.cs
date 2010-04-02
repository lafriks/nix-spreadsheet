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
		const string NS_MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
		const string NS_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
		const string NS_DCMI = "http://purl.org/dc/elements/1.1/";
		const string NS_META = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
		const string NS_GRDDL_DATA_VIEW = "http://www.w3.org/2003/g/data-view#";
		const string DATE_TIME_FORMAT = "yyyy-MM-ddTHH:mm:ss";

		protected Stream CreateManifestStream()
		{
			XmlWriterSettings xws = new XmlWriterSettings();
			xws.Indent = false;

			MemoryStream ms = new MemoryStream();
			XmlWriter xml = XmlWriter.Create(ms, xws);
			xml.WriteStartDocument();
			xml.WriteStartElement("manifest", "manifest", NS_MANIFEST);

			// File entry
			xml.WriteStartElement("file-entry", NS_MANIFEST);

			xml.WriteAttributeString("media-type", NS_MANIFEST, "application/vnd.oasis.opendocument.spreadsheet");
			xml.WriteAttributeString("version", NS_MANIFEST, "1.2");
			xml.WriteAttributeString("full-path", NS_MANIFEST, "/");

			xml.WriteEndElement();
			
			// Main files
			foreach ( string file in new string[] {"content.xml", "styles.xml", "meta.xml"} )
			{
				xml.WriteStartElement("file-entry", NS_MANIFEST);
	
				xml.WriteAttributeString("media-type", NS_MANIFEST, "text/xml");
				xml.WriteAttributeString("full-path", NS_MANIFEST, file);
	
				xml.WriteEndElement();				
			}

			xml.WriteEndElement();

			xml.Close();

			ms.Seek(0, SeekOrigin.Begin);

			return ms;
		}

		protected Stream CreateMetaStream()
		{
			XmlWriterSettings xws = new XmlWriterSettings();
			xws.Indent = false;

			MemoryStream ms = new MemoryStream();
			XmlWriter xml = XmlWriter.Create(ms, xws);
			xml.WriteStartDocument();
			xml.WriteStartElement("office", "document-meta", NS_OFFICE);
			
			xml.WriteAttributeString("xmlns", "meta", null, NS_META);
			xml.WriteAttributeString("version", NS_OFFICE, "1.2");
			xml.WriteAttributeString("grddl", "transformation", NS_GRDDL_DATA_VIEW, "http://docs.oasis-open.org/office/1.2/xslt/odf2rdf.xsl");
			
			xml.WriteStartElement("meta", NS_OFFICE);

			xml.WriteElementString("creation-date", NS_META, DateTime.Now.ToString(DATE_TIME_FORMAT));
			xml.WriteElementString("generator", NS_META, "Nix.SpreadSheet/" + this.GetType().Assembly.GetName().Version.ToString());

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
			using (Stream sm = this.CreateMetaStream())
			{
				zip.AddStream("meta.xml", sm);
				sm.Close();
			}
			zip.Close();
		}
	}
}
