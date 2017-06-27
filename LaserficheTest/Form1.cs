using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Laserfiche.RepositoryAccess;
using Laserfiche.DocumentServices;
using System.Web;
using System.Reflection;
using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using System.Xml;

namespace LaserficheTest
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            //create a connection to the Laserfiche repository
            RepositoryRegistration myRegistration = new RepositoryRegistration("TEXAN-APP26", "Laserfiche");
            Session mySession = new Session();


            try
            {
                mySession.LogIn(myRegistration);
                //mySession.LogIn("Laserfiche User Name", " Laserfiche Password", myRegistration);

                //CrawlFolder(mySession, 1);

                //GetDocumentByEntryID(mySession, 1253945);
                InsertDocument(mySession);
                //UpdateDocument(mySession, 1257492);
                //GetAllDocumentsInFolder(mySession);

                //Dictionary<string, object> searchValues = new Dictionary<string, object>();
                ////searchValues.Add("Document", "doc");
                ////searchValues.Add("Type", "typ");
                //searchValues.Add("Cat4", 4);
                //GetAllDocumentsByTemplate(mySession, @"\catapult","Field Types",searchValues);


                //GetAllTemplates(mySession);
                //UpdateTemplate(mySession, 1);
                //TemplateInfo ti = GetTemplateByTemplateId(mySession, 1);
                //Console.WriteLine(ti.Name);
                //TemplateInfo ti = GetTemplateByTemplateName(mySession, "General");
                //Console.WriteLine(ti.Name);



            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Logging off...");
                if (mySession.LogInTime.Year.ToString() != "1")
                {
                    mySession.LogOut();
                }
                mySession = null;
                myRegistration = null;
            }
        }



        private void GetAllDocumentsInFolder(Session session)
        {
            //https://answers.laserfiche.com/questions/89702/SDK-Search-with-Toolkit-RA#90132
            string myFolder = @"Laserfiche\catapult";
            string searchParameters = String.Format("{{LF:Lookin=\"{0}\"}}", myFolder);

            Search lfSearch = new Search(session, searchParameters);
            SearchListingSettings settings = new SearchListingSettings();
            settings.AddColumn(SystemColumn.Id);

            lfSearch.Run();

            SearchResultListing searchResults = lfSearch.GetResultListing(settings);

            foreach (EntryListingRow item in searchResults)
            {
                Int32 docId = (Int32)item[SystemColumn.Id];

                EntryInfo entryInfo = Entry.GetEntryInfo(docId, session);
                if (entryInfo.EntryType == EntryType.Shortcut)
                    entryInfo = Entry.GetEntryInfo(((ShortcutInfo)entryInfo).TargetId, session);

                // Now entry should be the DocumentInfo
                if (entryInfo.EntryType == EntryType.Document)
                {
                    DocumentInfo docInfo = (DocumentInfo)entryInfo;
                    Console.WriteLine(docInfo.Name);

                }

            }

        }


        private void GetAllDocumentsByTemplate(Session session, string folderName, string templateName, Dictionary<string,object> searchValues)
        {

            string searchFolder = String.Format("{{LF:Lookin=\"{0}\"}}", folderName);
            string searchParameters = string.Empty;
            string searchFields = string.Empty;

            foreach (KeyValuePair<string, object> kvPair in searchValues)
            {
                string tempSearch = string.Format(" & {{[{0}]:[{1}]=\"{2}\"}}", templateName,kvPair.Key.ToString(),kvPair.Value.ToString());
                searchFields += tempSearch;
            }

            searchParameters = searchFolder + searchFields;

            Search lfSearch = new Search(session, searchParameters);
            SearchListingSettings settings = new SearchListingSettings();
            settings.AddColumn(SystemColumn.Id);

            lfSearch.Run();

            SearchResultListing searchResults = lfSearch.GetResultListing(settings);

            foreach (EntryListingRow item in searchResults)
            {
                Int32 docId = (Int32)item[SystemColumn.Id];

                EntryInfo entryInfo = Entry.GetEntryInfo(docId, session);
                if (entryInfo.EntryType == EntryType.Shortcut)
                    entryInfo = Entry.GetEntryInfo(((ShortcutInfo)entryInfo).TargetId, session);

                // Now entry should be the DocumentInfo
                if (entryInfo.EntryType == EntryType.Document)
                {
                    DocumentInfo docInfo = (DocumentInfo)entryInfo;
                    Console.WriteLine(docInfo.Name);
                    FieldValueCollection fv = docInfo.GetFieldValues();
                    foreach (KeyValuePair<string, object> fieldValue in fv)
                    {
                        if (fieldValue.Value is Array)
                        {
                            string theValues = string.Empty;
                            foreach (object o in (Array)fieldValue.Value)
                            {
                                if (theValues.Length > 0)
                                    theValues += ", ";
                                theValues += o.ToString();
                            }
                            Console.WriteLine(theValues);
                        }

                    }
                }

            }

        }


        private void GetAllTemplates(Session session)
        {
            foreach (TemplateInfo templateInfo in Template.EnumAll(session))
            {
                if (templateInfo.Name == "Mobile Devices")
                {
                    Console.WriteLine(templateInfo.Name);
                    Console.WriteLine(templateInfo.Id);

                    Laserfiche.RepositoryAccess.FieldInfo f = Field.GetInfo(templateInfo.Fields[0].Name,session);

                }

            }
        }


        public TemplateInfo GetTemplateByTemplateId(Session session, int templateId)
        {
            List<TemplateInfo> templateList = new List<TemplateInfo>();
            foreach (TemplateInfo templateInfo in Template.EnumAll(session))
            {
                templateList.Add(templateInfo);
            }
            var query = from tl in templateList
                        where tl.Id == templateId
                        select tl;

            return query.FirstOrDefault();
        }

        public TemplateInfo GetTemplateByTemplateName(Session session, string templateName)
        {
            List<TemplateInfo> templateList = new List<TemplateInfo>();
            foreach (TemplateInfo templateInfo in Template.EnumAll(session))
            {
                templateList.Add(templateInfo);
            }
            var query = from tl in templateList
                        where tl.Name == templateName
                        select tl;

            return query.FirstOrDefault();


        }

        private void UpdateTemplate(Session session, int templateId)
        {
            TemplateInfo ti = Template.GetInfo(templateId, session);
            ti.Description = "";
            ti.Save();
        }





        //private void InsertDocumentII(Session session, Int32 entryId)
        //{
        //    //https://answers.laserfiche.com/questions/61772/Update-Metadata-Through-Java

        //    EntryInfo entryInfo = Entry.GetEntryInfo(entryId, session);
        //    if (entryInfo.EntryType == EntryType.Shortcut)
        //        entryInfo = Entry.GetEntryInfo(((ShortcutInfo)entryInfo).TargetId, session);

        //    // Now entry should be the DocumentInfo
        //    if (entryInfo.EntryType == EntryType.Document)
        //    {
        //        FieldValueCollection fvc = new FieldValueCollection();
        //        Template template = Template.getByName(entryInfo.getTe.getTemplate(), session);
        //        Field articleField = template.getFields().get("Article");

        //    }

        //}


        private void UpdateDocument(Session session, int entryId)
        {
            try
            {

                EntryInfo entryInfo = Entry.GetEntryInfo(entryId, session);
                if (entryInfo.EntryType == EntryType.Shortcut)
                    entryInfo = Entry.GetEntryInfo(((ShortcutInfo)entryInfo).TargetId, session);

                // Now entry should be the DocumentInfo
                if (entryInfo.EntryType == EntryType.Document)
                {
                    DocumentInfo docInfo = (DocumentInfo)entryInfo;
                    // Process the document here
                    //MessageBox.Show("Retrieved: " + docInfo.Name);

                    FieldValueCollection fv = docInfo.GetFieldValues();
                    fv.Remove("Document");
                    //fv.Add("Date", System.DateTime.Parse("01/01/1950"));
                    fv.Add("Document", "Document updated");
                    //fv.Add("Type", "Type");            
                    //fv.Add("Category", "Category");
                    //fv.Add("Addressee", "Addressee");
                    //fv.Add("Abstract", "Abstract");
                    //fv.Add("Category", "Category");
                    //fv.Add("Subject", "Subject");
                    //fv.Add("Author", "Author");
                    //fv.Add("Priority", "Priority");
                    fv.Remove("Exhibit");
                    fv.Add("Exhibit", "Exhibit updated");
                    //fv.Add("Exhibit", "Exhibit");
                    //fv.Add("Source", "Soource");
                    //fv.Add("Code", "Code");
                    //fv.Add("ssn", "113-22-3333");                          
                    docInfo.SetFieldValues(fv);
                    docInfo.Save();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void InsertDocument(Session session)
        {
            //https://answers.laserfiche.com/questions/49295/Using-RA-to-create-a-document-assign-a-template-and-populate-fields
            try
            {

                FolderInfo parentFolder = Folder.GetFolderInfo("\\catapult", session);
                DocumentInfo document = new DocumentInfo(session);
                document.Create(parentFolder, "jeffKNewDocument - " + DateTime.UtcNow.ToString(), EntryNameOption.None);

                //document.SetTemplate("Field Types");
                //FieldValueCollection fv = new FieldValueCollection();
                //string values = "1,2,3,4";
                //fv.Add("Cat1", items);
                //fv.Add("Cat2", 2);
                //fv.Add("Cat3", 3);
                //fv.Add("Cat4", 4);
                //fv.Add("Cat5", 5);
                //fv.Add("Cat6", System.DateTime.Parse("01/01/1950"));  //date?
                //fv.Add("Cat8", System.DateTime.Parse("2017-01-31T12:12:12"));  //date time?
                //fv.Add("Cat9", System.DateTime.Parse("6/22/2009 07:00:00 AM").ToString("HH:mm"));

                //document.SetTemplate("General");
                //FieldValueCollection fv = new FieldValueCollection();
                //fv.Add("Document", "This is the document name");
                //fv.Add("Type", "This is the document type");
                //fv.Add("Category", "This is the document category");

                document.SetTemplate("Mobile Devices");
                FieldValueCollection fv = new FieldValueCollection();
                fv.Add("First Name", "jeff");
                fv.Add("Last Name", "staubach");
                fv.Add("Line", "111-111-1111");
                fv.Add("Model", "Apple iPhone 5, 16 GB");
                fv.Add("Year", "2017");
                fv.Add("Amount", 21.12);  //date?

                //string theDocumentPath = @"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\CCITT_1.tif";
                //byte[] image = System.IO.File.ReadAllBytes(theDocumentPath);
                //PageInfo newPage = document.AppendPage();
                //System.IO.Stream pageStream = newPage.WritePagePart(PagePart.Image, image.Length);
                //pageStream.Write(image, 0, image.Length);
                //pageStream.Dispose();
                //newPage.Save();
                //image = System.IO.File.ReadAllBytes(theDocumentPath);
                //newPage = document.AppendPage();
                //pageStream = newPage.WritePagePart(PagePart.Image, image.Length);
                //pageStream.Write(image, 0, image.Length);
                //pageStream.Dispose();
                //newPage.Save();
                //image = System.IO.File.ReadAllBytes(theDocumentPath);
                //newPage = document.AppendPage();
                //pageStream = newPage.WritePagePart(PagePart.Image, image.Length);
                //pageStream.Write(image, 0, image.Length);
                //pageStream.Dispose();
                //newPage.Save();



                //string theDocumentPath = @"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\Program.cs";
                //DocumentImporter di = new DocumentImporter();
                //di.Document = document;
                //di.ImportText(theDocumentPath);

                //string theDocumentPath = @"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\CCITT_1.tif";
                //DocumentImporter di = new DocumentImporter();
                //di.Document = document;
                //di.ExtractTextFromEdoc = true;
                //di.ImportEdoc("image/tiff", GenerateStreamFromFile(theDocumentPath));

                //string theDocumentPath = @"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\CCITT_1.tif";
                //DocumentImporter di = new DocumentImporter();
                //di.Document = document;
                //di.ExtractTextFromEdoc = true;
                //di.ImportEdoc("application/pdf", GenerateStreamFromFile(theDocumentPath));

                //string theDocumentPath = @"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\Washington, Sean.pdf";
                //DocumentImporter di = new DocumentImporter();
                //di.Document = document;
                ////di.ImportEdoc("application/pdf",theDocumentPath);
                //di.ImportEdoc("application/pdf", GenerateStreamFromFile(theDocumentPath));

                ////byte[] documentBytes = System.IO.File.ReadAllBytes(@"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\Washington, Sean.pdf");
                //byte[] documentBytes = System.IO.File.ReadAllBytes(@"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\Washington, Sean.pdf");
                //byte[] documentBytes = Encoding.ASCII.GetBytes(richTextBox1.Text);
                //PageInfo newPage = document.AppendPage();
                //System.IO.Stream pageStream = newPage.WritePagePart(PagePart.Image, documentBytes.Length);
                //pageStream.Write(documentBytes, 0, documentBytes.Length);
                //pageStream.Dispose();
                //newPage.Save();


                //http://answers.laserfiche.com/questions/55829/How-to-convert-PDF-Byte-Array-into-PDF-document
                //byte[] documentBytes = System.IO.File.ReadAllBytes(@"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\Washington, Sean.pdf");
                //byte[] documentBytes = Encoding.ASCII.GetBytes(richTextBox1.Text);
                //using (System.IO.Stream edocStream = document.WriteEdoc("application/pdf", documentBytes.Length))
                //{
                //    edocStream.Write(documentBytes, 0, documentBytes.Length);
                //}
                //document.Extension = ".pdf";

                //System.IO.MemoryStream ms = new System.IO.MemoryStream(ASCIIEncoding.Default.GetBytes("Your string here"));
                //using (System.IO.Stream edocStream = document.WriteEdoc("text/plain", ms.ToArray().LongLength))
                //{
                //    edocStream.Write(ms.ToArray(), 0, ms.ToArray().Length);
                //}
                //document.Extension = ".txt";


                //System.IO.MemoryStream ms = new System.IO.MemoryStream(System.IO.File.ReadAllBytes(@"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\Washington, Sean.pdf"));
                //using (System.IO.Stream edocStream = document.WriteEdoc("application/pdf", ms.ToArray().LongLength))
                //{
                //    edocStream.Write(ms.ToArray(), 0, ms.ToArray().Length);
                //}
                //document.Extension = ".pdf";

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(@"C:\Users\jeff.kanarr\Documents\Visual Studio 2015\Projects\LaserficheTest\LaserficheTest\base64.txt");
                Byte[] bytes = Convert.FromBase64String(xmlDoc.SelectSingleNode("//file/content").InnerText);
                using (System.IO.Stream edocStream = document.WriteEdoc("application/pdf", bytes.ToArray().LongLength))
                {
                    edocStream.Write(bytes.ToArray(), 0, bytes.ToArray().Length);
                }
                document.Extension = ".pdf";



                document.SetFieldValues(fv);
                document.Save();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }




        // This method converts the filestream into a byte array so that when it is 
        // used in my ASP.Net project the file can be sent using response.Write
        public static System.IO.Stream GenerateStreamFromFile(string theFilePathAndName)
        {
            System.IO.MemoryStream data = new System.IO.MemoryStream();
            System.IO.Stream str = System.IO.File.OpenRead(theFilePathAndName);

            str.CopyTo(data);
            data.Seek(0, System.IO.SeekOrigin.Begin);
            byte[] buf = new byte[data.Length];
            data.Read(buf, 0, buf.Length);
            return data;
        }

        public static System.IO.Stream GenerateStreamFromString(string s)
        {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            System.IO.StreamWriter writer = new System.IO.StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }


        public static string GetMimeType(string extension, out bool wasFound)
        {
            wasFound = false;
            string mimeType = "application/octet-stream";
            if (!string.IsNullOrEmpty(extension))
            {
                if (!extension.StartsWith("."))
                {
                    extension = "." + extension;
                }

                try
                {
                    bool setMimeType = false;
                    Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(extension.ToLowerInvariant(), false);
                    if (key != null)
                    {
                        string value = System.Convert.ToString(key.GetValue("Content Type", string.Empty));
                        if (!string.IsNullOrEmpty(value))
                        {
                            mimeType = value;
                            wasFound = true;
                        }

                        key.Close();
                    }

                    wasFound = setMimeType;
                }
                catch (System.Security.SecurityException)
                {
                    // No rights to key.
                }
            }

            return mimeType;
        }


        public static class MimeMappingStealer
        {
            // The get mime mapping method info
            private static readonly MethodInfo _getMimeMappingMethod = null;

            /// <summary>
            /// Static constructor sets up reflection.
            /// </summary>
            static MimeMappingStealer()
            {
                // Load hidden mime mapping class and method from System.Web
                var assembly = Assembly.GetAssembly(typeof(HttpApplication));
                Type mimeMappingType = assembly.GetType("System.Web.MimeMapping");
                _getMimeMappingMethod = mimeMappingType.GetMethod("GetMimeMapping",
                    BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public |
                    BindingFlags.NonPublic | BindingFlags.FlattenHierarchy);
            }

            /// <summary>
            /// Exposes the hidden Mime mapping method.
            /// </summary>
            /// <param name="fileName">The file name.</param>
            /// <returns>The mime mapping.</returns>
            public static string GetMimeMapping(string fileName)
            {
                return (string)_getMimeMappingMethod.Invoke(null /*static method*/, new[] { fileName });
            }
        }


        private void GetDocumentByEntryID(Session session, int entryId)
        {
            EntryInfo entryInfo = Entry.GetEntryInfo(entryId, session);
            if (entryInfo.EntryType == EntryType.Shortcut)
                entryInfo = Entry.GetEntryInfo(((ShortcutInfo)entryInfo).TargetId, session);

            // Now entry should be the DocumentInfo
            if (entryInfo.EntryType == EntryType.Document)
            {
                DocumentInfo docInfo = (DocumentInfo)entryInfo;
                FieldValueCollection fv = docInfo.GetFieldValues();
                // Process the document here
                MessageBox.Show("Retrieved: " + docInfo.Name);
            }
        }

    static void CrawlFolder(Session session, int folderId, int depth = 0)
        {
            FolderInfo folder = Folder.GetFolderInfo(folderId, session);

            Console.WriteLine(folder.Path);

            EntryListingSettings settings = new EntryListingSettings();
            settings.AddColumn(SystemColumn.Name);
            settings.AddColumn(SystemColumn.Id);
            settings.AddColumn(SystemColumn.EntryType);
            settings.SetSortColumn(SystemColumn.Name, SortDirection.Ascending);

            List<int> subfolders = new List<int>();

            using (FolderListing listing = folder.OpenFolderListing(settings))
            {
                foreach (var row in listing)
                {
                    EntryType etype = (EntryType)row[SystemColumn.EntryType];
                    int entryId = (int)row[SystemColumn.Id];

                    if (etype == EntryType.Document)
                    {
                        DocumentInfo doc = Document.GetDocumentInfo(entryId, session);
                        Console.WriteLine("  " + doc.Name);
                    }
                    else if (etype == EntryType.Folder || etype == EntryType.RecordSeries)
                    {
                        FolderInfo subfolder = Folder.GetFolderInfo(entryId, session);

                        subfolders.Add(entryId);
                    }
                }
            }

            // Now that the FolderListing is closed, crawl the subfolders
            foreach (int subfolderId in subfolders)
            {
                CrawlFolder(session, subfolderId, depth + 1);
            }
        }


    }
}
