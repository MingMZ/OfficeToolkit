using System;
using System.AddIn.Pipeline;
using OfficeToolkit.AddIns.Contracts.Access;
using OfficeToolkit.AddIns.AddInViews.Access;

namespace OfficeToolkit.AddIns.AddInAdapters.Access
{
    [AddInAdapter()]
    public class AccessCompositionAdapter : ContractBase, IAccessCompositionContract
    {
        private IAccessComposition _view;

        public AccessCompositionAdapter(IAccessComposition view)
        {
            _view = view;
        }

        public void Open(string filename)
        {
            _view.Open(filename);
        }

        public void Close()
        {
            _view.Close();
        }

        public void SaveObjects(string baseDirectory)
        {
            _view.SaveObjects(baseDirectory);
        }

        public void LoadObjects(string baseDirectory)
        {
            _view.LoadObjects(baseDirectory);
        }

        public void ClearObjects()
        {
            _view.ClearObjects();
        }

        public void Dispose()
        {
            _view.Dispose();
        }
    }
}
