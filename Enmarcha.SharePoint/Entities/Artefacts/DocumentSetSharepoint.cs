using Enmarcha.SharePoint.Abstract.Interfaces.Artefacts;
using Microsoft.Office.DocumentManagement.DocumentSets;
using Microsoft.SharePoint;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Enmarcha.SharePoint.Entities.Artefacts
{
    public sealed class DocumentSetSharepoint : IDocumentSet
    {
        public SPList List { get; set; }
        public string Name { get; set; }
        public string FolderPath { get; set; }
        public ILog Logger { get; set; }

        public DocumentSetSharepoint(SPList list, string name, string folderPath, ILog logger)
        {
            List = list;
            Name = name;
            FolderPath = folderPath;
            Logger = logger;
        }

        public DocumentSetSharepoint(SPList list, string name, ILog logger)
        {
            List = list;
            Name = name;
            FolderPath = List.RootFolder.Name;
            Logger = logger;
        }

        public int Create(string description)
        {
            return Create(description, "Document Set");
        }

        public int Create(string description, string contentTypeName)
        {
            try
            {
                int docSetId = -1;
                SPWeb web = List.ParentWeb;
                SPFolder folder = web.GetFolder(string.Concat(web.ServerRelativeUrl, List.Title, FolderPath));
                Hashtable properties = new Hashtable();
                properties.Add("DocumentSetDescription", description);
                DocumentSet docSet = DocumentSet.Create(folder, Name, List.ContentTypes[contentTypeName].Id, properties, true);
                if (docSet != null)
                {
                    docSetId = docSet.Item.ID;
                    SPListItem docSetItem = List.GetItemById(docSetId);
                    if (docSetItem != null)
                    {
                        docSetItem[SPBuiltInFieldId.Author] = List.ParentWeb.CurrentUser;
                        docSetItem[SPBuiltInFieldId.Editor] = List.ParentWeb.CurrentUser;
                        docSetItem.Update();
                    }
                }
                return docSetId;
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Exception Creating document set", exception.Message));
                return -1;
            }
        }

        public bool AddDocument(int id, byte[] file, string fileName, bool overWrite)
        {
            try
            {
                var docSetItem = List.GetItemById(id);
                var spFile = docSetItem.Folder.Files.Add(fileName, file, overWrite);
                return true;
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Exception uploading file to Document Set", exception.Message));
                return false;
            }
        }

        public bool AddDocument(int id, Stream file, string fileName, bool overWrite)
        {
            try
            {
                var docSetItem = List.GetItemById(id);
                var spFile = docSetItem.Folder.Files.Add(fileName, file, overWrite);
                return true;
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Exception uploading file to Document Set", exception.Message));
                return false;
            }
        }

        public void Delete(int id)
        {
            try
            {
                List.Items.DeleteItemById(id);
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Exception deleting Document Set", exception.Message));
            }
        }

        public void DeleteDocument(int id, string fileName)
        {
            try
            {
                var docSetItem = List.Items.GetItemById(id);
                docSetItem.Folder.Files.Delete(fileName);
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Exception deleting document from Document Set", exception.Message));
            }
        }

        public bool Exists()
        {
            throw new NotImplementedException();
        }
    }
}
