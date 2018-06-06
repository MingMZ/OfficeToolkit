using System;
using System.AddIn.Pipeline;
using OfficeToolkit.AddIns.Contracts.Access;
using OfficeToolkit.AddIns.HostViews.Access;

namespace OfficeToolkit.AddIns.HostSideAdapters.Access
{
    [HostAdapterAttribute()]
    public class AccessCompositionAdapter : IAccessComposition
    {
        private IAccessCompositionContract _contract;
        private ContractHandle _handle;

        public AccessCompositionAdapter(IAccessCompositionContract contract)
        {
            _contract = contract;
            _handle = new ContractHandle(contract);
        }

        public void Open(string filename)
        {
            _contract.Open(filename);
        }

        public void Close()
        {
            _contract.Close();
        }

        public void SaveObjects(string baseDirectory)
        {
            _contract.SaveObjects(baseDirectory);
        }

        public void LoadObjects(string baseDirectory)
        {
            _contract.LoadObjects(baseDirectory);
        }

        public void ClearObjects()
        {
            _contract.ClearObjects();
        }

        public void Dispose()
        {
            _contract.Dispose();
        }
    }
}
