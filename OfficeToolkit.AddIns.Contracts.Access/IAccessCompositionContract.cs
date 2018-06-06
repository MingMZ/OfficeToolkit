using System;
using System.AddIn.Contract;
using System.AddIn.Pipeline;

namespace OfficeToolkit.AddIns.Contracts.Access
{
    [AddInContract]
    public interface IAccessCompositionContract : IContract, IDisposable
    {
        void Open(string filename);
        void Close();

        void SaveObjects(string baseDirectory);
        void LoadObjects(string baseDirectory);
        void ClearObjects();
    }
}
