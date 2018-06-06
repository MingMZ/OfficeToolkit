using System;
using System.Collections.Generic;
using System.Linq;
using System.AddIn;
using System.IO;
using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using OfficeToolkit.AddIns.AddInViews.Access;

namespace OfficeToolkit.Access2010
{
    [AddIn("Access 2010 Composition", Version = "1.0.0.0")]
    public class AccessComposition : IAccessComposition
    {
        private const string EXTENSION = ".txt";

        private Application _application;
        private Database _database;

        public void Open(string filename)
        {
            if (_application == null)
            {
                _application = new Microsoft.Office.Interop.Access.Application();

                if (_database != null)
                    DisposeDatabase();
            }

            _application.OpenCurrentDatabase(filename, true);

            _database = _application.CurrentDb();
        }

        public void Close()
        {
            DisposeDatabase();
        }

        public void SaveObjects(string baseDirectory)
        {
            foreach (Container container in _database.Containers)
            {
                switch (container.Name)
                {
                    case "Forms":
                        SaveObjects(baseDirectory, container.Name, container, AcObjectType.acForm);
                        break;
                    case "Modules":
                        SaveObjects(baseDirectory, container.Name, container, AcObjectType.acModule);
                        break;
                    case "Reports":
                        SaveObjects(baseDirectory, container.Name, container, AcObjectType.acReport);
                        break;
                    case "Scripts":
                        SaveObjects(baseDirectory, "Macros", container, AcObjectType.acMacro);
                        break;
                }
            }

            DirectoryInfo di = Directory.CreateDirectory(Path.Combine(baseDirectory, "Queries"));
            foreach (QueryDef td in _database.QueryDefs)
            {
                _application.Application.SaveAsText(AcObjectType.acQuery, td.Name, Path.Combine(baseDirectory, di.Name, td.Name + ".txt"));
            }
        }

        private void SaveObjects(string baseDirectory, string subDirectoryName, Container acContainer, AcObjectType acType)
        {
            DirectoryInfo di = Directory.CreateDirectory(Path.Combine(baseDirectory, subDirectoryName));

            foreach (Document document in acContainer.Documents)
            {
                _application.Application.SaveAsText(acType, document.Name, Path.Combine(baseDirectory, di.Name, document.Name + ".txt"));
            }
        }

        public void LoadObjects(string baseDirectory)
        {

            foreach (Container container in _database.Containers)
            {
                switch (container.Name)
                {
                    case "Forms":
                        LoadObjects(baseDirectory, container.Name, container, AcObjectType.acForm);
                        break;
                    case "Modules":
                        LoadObjects(baseDirectory, container.Name, container, AcObjectType.acModule);
                        break;
                    case "Reports":
                        LoadObjects(baseDirectory, container.Name, container, AcObjectType.acReport);
                        break;
                    case "Scripts":
                        LoadObjects(baseDirectory, "Macros", container, AcObjectType.acMacro);
                        break;
                }
            }

            DirectoryInfo di = Directory.CreateDirectory(Path.Combine(baseDirectory, "Queries"));
            FileInfo[] files = di.GetFiles("*" + EXTENSION);
            IList<string> existingObjects = new List<String>();
            foreach (QueryDef q in _database.QueryDefs)
            {
                existingObjects.Add(q.Name);
            }
            string[] conflictObjects = GetConflictObjectName(files, existingObjects.ToArray(), AcObjectType.acQuery);
            foreach (FileInfo f in files)
            {
                _application.Application.LoadFromText(AcObjectType.acQuery, Path.GetFileNameWithoutExtension(f.Name), f.FullName);
            }
        }

        private void LoadObjects(string baseDirectory, string subDirectoryName, Container acContainer, AcObjectType acType)
        {
            // to avoid loading file which have the same name as existing object in the database
            // first compare the names of files and objects, get a list of conflicts
            DirectoryInfo di = Directory.CreateDirectory(Path.Combine(baseDirectory, subDirectoryName));

            FileInfo[] files = di.GetFiles("*" + EXTENSION);

            IList<string> existingObjects = new List<String>();

            foreach (Document d in acContainer.Documents)
            {
                existingObjects.Add(d.Name);
            }

            string[] conflictObjects = GetConflictObjectName(files, existingObjects.ToArray(), acType);

            // then remove those conflicting objects
            ClearObjects(acType, conflictObjects);

            foreach (FileInfo f in files)
            {
                _application.LoadFromText(acType, Path.GetFileNameWithoutExtension(f.Name), f.FullName);
            }
        }

        private string[] GetConflictObjectName(FileInfo[] files, string[] names, AcObjectType acType)
        {
            IEnumerable<string> conflictObjects =
                names.Except(
                    (from f in files select Path.GetFileNameWithoutExtension(f.Name)).ToList()
                );

            return conflictObjects.ToArray();
        }

        public void ClearObjects()
        {
            throw new NotImplementedException();
        }

        private void ClearObjects(AcObjectType t, params string[] names)
        {
            foreach (string name in names)
            {
                _application.DoCmd.DeleteObject(t, name);
            }
        }

        public void Dispose()
        {
            DisposeDatabase();
            DisposeApplication();

            GC.Collect();
        }

        private void DisposeDatabase()
        {
            if (_database != null)
            {
                try
                {
                    _application.CloseCurrentDatabase();
                }
                catch { }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_database);
                    _database = null;
                }
            }
        }

        private void DisposeApplication()
        {
            if (_application != null)
            {
                try
                {
                    _application.Quit();
                }
                catch { }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_application);
                    _application = null;
                }
            }
        }
    }
}
