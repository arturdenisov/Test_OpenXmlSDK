using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace OpenXmlSDK
{
    internal class DocxClass
    {

        internal static void GenerateFile()
        {
    
            var xmlContainer = new XElement("Customers");
            var xmlItem1 = new XElement("Customer1",
                            new XElement("Name", "Artur4"),
                            new XElement("Expiration", "2/1/1999"),
                            new XElement("AmountDue", "$999.99"));
            var xmlItem2 = new XElement("Customer2",
                            new XElement("Name", "Natali4"),
                            new XElement("Expiration", "2/1/2020"),
                            new XElement("AmountDue", "$900.00"));

            xmlContainer.Add(xmlItem1);
            xmlContainer.Add(xmlItem2);


            string projectDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            string filePath = projectDirectory + @"\Test.docx";


            using (var wpd = WordprocessingDocument.Open(filePath, true))
            {
                var mainPart = wpd.MainDocumentPart;


                // (2) Write Custom XML file into .DOCX
                var xmlPart = mainPart.AddCustomXmlPart("application/xml"); // you can pass a custom XmlPart or PartId: https://social.msdn.microsoft.com/Forums/office/en-US/95db2ea1-aa48-4b7d-99ae-86b3bad1bdd2/the-type-documentformatopenxmlpackagingcustomxmlpart-cannot-be-used-as-type-parameter-t?forum=oxmlsdk
                using (Stream partStream = xmlPart.GetStream(FileMode.Create, FileAccess.Write))
                {

                    using (StreamWriter outputStream = new StreamWriter(partStream))
                    {
                        outputStream.Write(xmlContainer);
                    }
                }

                // (3) Correct Document Structure in connections with XmlList.Count() and set correct Tags / Tag-Names
                Body body = mainPart.Document.Descendants<Body>().First();
                Paragraph nodeForClone = mainPart.Document.Descendants<Paragraph>().First();
                Paragraph nodeCloned = (Paragraph)nodeForClone.CloneNode(true);
                body.InsertAfter(nodeCloned, nodeForClone);

                // (4) Rename Tag and Its alias
                var paragraphList = mainPart.Document.Descendants<Paragraph>().ToList();
                var paragraph1 = paragraphList[0].Descendants<SdtRun>().ToList();
                foreach(var stdRun in paragraph1)
                {
                    if(stdRun.Descendants<Tag>().First().Val.ToString().Contains("Name"))
                    { 
                        string tagName = "Customers/Customer1/Name";
                        stdRun.Descendants<Tag>().First().Val = tagName;
                    }
                    if (stdRun.Descendants<Tag>().First().Val.ToString().Contains("Expiration"))
                    {
                        string tagName = "Customers/Customer1/Expiration";
                        stdRun.Descendants<Tag>().First().Val = tagName;
                    }
                    if (stdRun.Descendants<Tag>().First().Val.ToString().Contains("AmountDue"))
                    {
                        string tagName = "Customers/Customer1/AmountDue";
                        stdRun.Descendants<Tag>().First().Val = tagName;
                    }
                }

                var paragraph2 = paragraphList[1].Descendants<SdtRun>().ToList();
                foreach (var stdRun in paragraph2)
                {
                    if (stdRun.Descendants<Tag>().First().Val.ToString().Contains("Name"))
                    {
                        string tagName = "Customers/Customer2/Name";
                        stdRun.Descendants<Tag>().First().Val = tagName;
                    }
                    if (stdRun.Descendants<Tag>().First().Val.ToString().Contains("Expiration"))
                    {
                        string tagName = "Customers/Customer2/Expiration";
                        stdRun.Descendants<Tag>().First().Val = tagName;
                    }
                    if (stdRun.Descendants<Tag>().First().Val.ToString().Contains("AmountDue"))
                    {
                        string tagName = "Customers/Customer2/AmountDue";
                        stdRun.Descendants<Tag>().First().Val = tagName;
                    }
                }

                // (5) Bind Custom XML File
                var taggedContentControls =
                  from sdt in mainPart.Document.Descendants<SdtRun>()
                  let sdtPr = sdt.GetFirstChild<SdtProperties>()
                  let tag = (sdtPr == null ? null : sdtPr.GetFirstChild<Tag>())
                  where tag != null
                  select new
                  {
                      SdtProps = sdtPr,
                      TagName = tag.Val 
                  };

                foreach (var taggedContentControl in taggedContentControls)
                {
                    var binding = new DataBinding();
                    binding.XPath = taggedContentControl.TagName;
                    taggedContentControl.SdtProps.Append(binding);
                }

                // (5) Save DOCX File
                mainPart.Document.Save();
            }
        }
    }
}