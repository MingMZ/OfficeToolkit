using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.AddIn.Hosting;
using OfficeToolkit.AddIns.HostViews.Access;

namespace OfficeToolkit.UI.Access
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _addInRoot;



        public MainWindow()
        {
            InitializeComponent();

            _addInRoot = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                "Pipeline");


            string[] warnings = AddInStore.Update(_addInRoot);

            Collection<AddInToken> tokens = AddInStore.FindAddIns(typeof(IAccessComposition), _addInRoot);

            comboBoxProvider.ItemsSource = tokens;
            comboBoxProvider.SelectedIndex = 0;

            //radioButtonLoad.IsChecked = true;
            radioButtonSave.IsChecked = true;
        }

        private void buttonSelectFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.Filter = "All Files (*.*)|*.*|" +
                         "Microsoft Access (*.accdb;*.mdb;*.adp;*.mda)|*.accdb;*.mdb;*.adp;*.mda|" +
                         "Microsoft Access Databases (*.mdb;*.acdb)|*.mdb;*.accdb|" +
                         "Microsoft Access 2007 Databases (*.accdb)|*.accdb|" +
                         "Microsoft Access Projects (*.adp)|*.adp";
            ofd.Multiselect = false;
            ofd.FilterIndex = 2;

            bool? result = ofd.ShowDialog();

            if (result == true)
            {
                textBoxFile.Text = ofd.FileName;
                if (radioButtonSave.IsChecked == true)
                    textBoxFolder.Text = GetDefaultFolder(ofd.FileName);
            }
        }

        private string GetDefaultFolder(string path)
        {
            return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), System.IO.Path.GetFileNameWithoutExtension(path), DateTime.Now.ToString("yyyyMMdd-hhmmss"));
        }
        
        // WTF MS, second WPF major release now and it still managed to missing out some important controls
        private void buttonSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            //System.Windows.Forms.OpenFileDialog d = new System.Windows.Forms.OpenFileDialog();
            //d.CheckFileExists = false;
            //d.CheckPathExists = true;
            //d.Multiselect = false;

            System.Windows.Forms.FolderBrowserDialog d = new System.Windows.Forms.FolderBrowserDialog();

            d.ShowNewFolderButton = false;
            
            var result = d.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                //textBoxFolder.Text = d.FileName;
                textBoxFolder.Text = d.SelectedPath;
            }
        }

        private void buttonExecute_Click(object sender, RoutedEventArgs e)
        {
            AddInToken t = comboBoxProvider.SelectedItem as AddInToken;

            AddInProcess addInProcess = new AddInProcess();

            try
            {
                using (IAccessComposition iac = t.Activate<IAccessComposition>(addInProcess, AddInSecurityLevel.FullTrust))
                {
                    if (radioButtonLoad.IsChecked == true)
                    {
                        iac.Open(textBoxFile.Text);
                        iac.LoadObjects(textBoxFolder.Text);
                        iac.Close();
                    }
                    else if (radioButtonSave.IsChecked == true)
                    {
                        iac.Open(textBoxFile.Text);
                        iac.SaveObjects(textBoxFolder.Text);
                        iac.Close();
                    }
                    else if (radioButtonClear.IsChecked == true)
                    {
                        iac.Open(textBoxFile.Text);
                        iac.ClearObjects();
                        iac.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Message: {0}\r\n\r\nStack Trace: {1}", ex.Message, ex.StackTrace),"Error");
            }
            finally
            {
                if (!addInProcess.IsCurrentProcess)
                    addInProcess.Shutdown();
            }
        }

        #region textBoxFolder drag and drop

        private void textBoxFolder_PreviewDragOver(object sender, DragEventArgs e)
        {
            TextBox control = sender as TextBox;
            e.Effects = DragDropEffects.Copy;

            e.Handled = true;
        }

        private void textBoxFolder_PreviewDrop(object sender, DragEventArgs e)
        {
            TextBox control = sender as TextBox;
            e.Effects = DragDropEffects.None;

            if (control != null & e.Data is DataObject)
            {
                DataObject data = e.Data as DataObject;

                var paths = data.GetFileDropList();

                foreach (string path in paths)
                {
                    if (System.IO.Directory.Exists(path))
                    {
                        control.Text = path;
                    }
                }
            }

            e.Handled = true;
        }
        #endregion

        #region textBoxFile drag and drop
        private void textBoxFile_PreviewDragOver(object sender, DragEventArgs e)
        {
            TextBox control = sender as TextBox;
            e.Effects = DragDropEffects.Copy;

            e.Handled = true;
        }

        private void textBoxFile_PreviewDrop(object sender, DragEventArgs e)
        {
            TextBox control = sender as TextBox;
            e.Effects = DragDropEffects.None;

            if (control != null & e.Data is DataObject)
            {
                DataObject data = e.Data as DataObject;

                var paths = data.GetFileDropList();

                foreach (string path in paths)
                {
                    if (System.IO.File.Exists(path))
                    {
                        control.Text = path;

                        if (radioButtonSave.IsChecked == true)
                            this.textBoxFolder.Text = GetDefaultFolder(path);
                    }
                }
            }

            e.Handled = true;
        }
        #endregion
    }
}
