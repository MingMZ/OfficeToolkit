using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeToolkit.AddIns.HostViews.Access
{
    public interface IAccessComposition : IDisposable
    {
        void Open(string filename);
        void Close();

        void SaveObjects(string baseDirectory);
        void LoadObjects(string baseDirectory);
        void ClearObjects();
    }
}
