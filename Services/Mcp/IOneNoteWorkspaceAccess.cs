using System;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Mcp
{
    internal interface IOneNoteWorkspaceAccess : IOneNotePageAccess
    {
        string GetHierarchy(string parentId, HierarchyScope scope);
        string FindPages(string rootId, string query);
        string CreatePage(string sectionId);
        string CreateSection(string parentId, string name);
        void Recycle(string id, DateTime expectedModified);
        string GetLink(string id, string objectId);
        void Navigate(string id, string objectId);
        void ExportPdf(string pageId, string path);
    }
}
