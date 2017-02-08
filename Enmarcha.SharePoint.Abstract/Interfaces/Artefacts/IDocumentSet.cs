using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Enmarcha.SharePoint.Abstract.Interfaces.Artefacts
{
    public interface IDocumentSet
    {
        int Create(string description);
        int Create(string description, string contentTypeId);
        void Delete(int id);
        bool AddDocument(int id, Stream file, string fileName, bool overWrite);
        bool AddDocument(int id, byte[] file, string fileName, bool overWrite);
        void DeleteDocument(int id, string fileName);
        bool Exists();
    }
}
