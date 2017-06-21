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
using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

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
                //InsertDocument(mySession);
                UpdateDocument(mySession, 1253945);
                //GetAllDocuments(mySession);


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



        private void GetAllDocuments(Session session)
        {

            string myFolder = @"Laserfiche\z_Training Manual";
            string searchParameters = String.Format("{{LF:Lookin=\"{0}\"}}", myFolder);

            Search lfSearch = new Search(session, searchParameters);
            SearchListingSettings settings = new SearchListingSettings();
            settings.AddColumn(SystemColumn.Id);

            lfSearch.Run();

            SearchResultListing searchResults = lfSearch.GetResultListing(settings);

            foreach (EntryListingRow item in searchResults)
            {
                Int32 docId = (Int32)item[SystemColumn.Id];
                //Console.Write(docId);
                //DocumentInfo docInfo = new DocumentInfo(docId, session);


                EntryInfo entryInfo = Entry.GetEntryInfo(docId, session);
                if (entryInfo.EntryType == EntryType.Shortcut)
                    entryInfo = Entry.GetEntryInfo(((ShortcutInfo)entryInfo).TargetId, session);

                // Now entry should be the DocumentInfo
                if (entryInfo.EntryType == EntryType.Document)
                {
                    DocumentInfo docInfo = (DocumentInfo)entryInfo;
                    // check for docInfo.Name match?
                }

            }

        }


        private void GetAllTemplates(Session session)
        {
            foreach (TemplateInfo templateInfo in Template.EnumAll(session))
            {
                if (templateInfo.Name == "General")
                {
                    Console.WriteLine(templateInfo.Name);
                    Console.WriteLine(templateInfo.Id);
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

                    FieldValueCollection fv = new FieldValueCollection();
                    fv.Add("Date", System.DateTime.Parse("01/01/1950"));
                    fv.Add("Document", "Document");        
                    fv.Add("Type", "Type");            
                    fv.Add("Category", "Category");
                    fv.Add("Addressee", "Addressee");
                    fv.Add("Abstract", "Abstract");
                    fv.Add("Category", "Category");
                    fv.Add("Subject", "Subject");
                    fv.Add("Author", "Author");
                    fv.Add("Priority", "Priority");
                    fv.Add("Exhibit", "Exhibit");
                    fv.Add("Source", "Soource");
                    fv.Add("Code", "Code");
                    fv.Add("ssn", "113-22-3333");                          

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
                document.SetTemplate("General");

                FieldValueCollection fv = new FieldValueCollection();
                //Add the metadata to the collection.  The first parameter is the field name and the
                //second parameter is the value.  NOTE: The value is type Object...
                fv.Add("Document", "This is the document name");        //This field is part of the General template
                fv.Add("Type", "This is the document type");            //So is this field...
                fv.Add("Category", "This is the document category");    //And so is this one...
                fv.Add("Subject", "13-12345");                          //This field is not part of the template...


                document.SetFieldValues(fv);
                document.Save();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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
