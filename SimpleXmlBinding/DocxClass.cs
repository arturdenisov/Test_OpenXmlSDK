using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace OpenXmlSDK
{
    internal class DocxClass
    {

        internal static void GenerateFile()
        {
            var xml =
              new XElement("Customer",
                new XElement("Name", "John Doe2"),
                new XElement("Expiration", "2/1/2010"),
                new XElement("AmountDue", "$129.50"));

            string projectDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            string filePath = projectDirectory + @"\Test.docx";

            using (var wpd = WordprocessingDocument.Open(filePath, true))
            {
                var mainPart = wpd.MainDocumentPart;
                
                var xmlPart = mainPart.AddCustomXmlPart("application/xml"); // you can pass a custom XmlPart or PartId: https://social.msdn.microsoft.com/Forums/office/en-US/95db2ea1-aa48-4b7d-99ae-86b3bad1bdd2/the-type-documentformatopenxmlpackagingcustomxmlpart-cannot-be-used-as-type-parameter-t?forum=oxmlsdk

                using (Stream partStream = xmlPart.GetStream(FileMode.Create, FileAccess.Write))
                {
                    using (StreamWriter outputStream = new StreamWriter(partStream))
                    {
                        outputStream.Write(xml);
                    }
                }

                var taggedContentControls =
                  from sdt in mainPart.Document.Descendants<SdtRun>()
                  let sdtPr = sdt.GetFirstChild<SdtProperties>()
                  let tag = (sdtPr == null ? null : sdtPr.GetFirstChild<Tag>())
                  where tag != null
                  select new
                  {
                      SdtProps = sdtPr,
                      TagName = tag.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value
                  };

                foreach (var taggedContentControl in taggedContentControls)
                {
                    var binding = new DataBinding();
                    binding.XPath = taggedContentControl.TagName;
                    taggedContentControl.SdtProps.Append(binding);
                }

                mainPart.Document.Save();
            }

        }
    }
}
