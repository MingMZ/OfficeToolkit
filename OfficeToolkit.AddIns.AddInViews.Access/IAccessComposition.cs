using System;
using System.AddIn.Pipeline;

namespace OfficeToolkit.AddIns.AddInViews.Access
{
    [AddInBase()]
    public interface IAccessComposition : IDisposable
    {
        void Open(string filename);
        void Close();

        void SaveObjects(string baseDirectory);
        void LoadObjects(string baseDirectory);
        void ClearObjects();
    }
}
