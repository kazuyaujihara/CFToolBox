using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Threading;
using System.Diagnostics;
using System.ComponentModel;
using System.IO;
using System.Xml.Serialization;
using CfxUtilityGUI.Properties;
using Ujihara.Chemistry.MergeSF;


namespace CfxUtilityGUI
{
    /// <summary>
    /// </summary>
    public partial class CfxToolsMainWindow : Window
    {
        /// <summary>
        /// Interval of refrsh in milli second.
        /// </summary>
        private const int IntervalOfRefresh = 1 * 1000;

        public CfxToolsMainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private const string _E_all_supported = "*.rtf;*.doc;*.docx;*.docm;*.sdf;*.csv;*.txt;*.lst";
        private const string E_all_supported = "All supported files (" + _E_all_supported + ")|" + _E_all_supported;
        private const string E_rtf = "Word files (*.rtf;*.doc;*.docx;*.docm)|*.rtf;*.doc;*.docx;*.docm";
        private const string E_sdf = "SD files (*.sdf)|*.sdf";
        private const string E_all = "All files (*.*)|*.*";
        private const string E_cfx = "ChemFinder files (*.cfx)|*.cfx";
        private const string E_csv = "CSV files (*.csv;*.txt)|:.csv;*.txt";
        private const string E_lst = "Lists (*.lst)|*.lst";
        private const string E_image = "Images (*.gif;*.png;*.jpg;*.tif;*.tiff)|*.gif;*.png;*.jpg;*.tif;*.tiff";

        private void AppendFromCasOnlineItem_Click(object sender, RoutedEventArgs e)
        {
            CreateFromCasOnlineItem_Click(sender, e, true);
        }

        private void CreateFromCasOnlineItem_Click(object sender, RoutedEventArgs e)
        {
            CreateFromCasOnlineItem_Click(sender, e, false);
        }

        private static Ujihara.Chemistry.MergeSF.Program CreateProgram()
        {
            var prog =  new Ujihara.Chemistry.MergeSF.Program();
            prog.LoadSubstanceImageFlag = Properties.Settings.Default.LoadSubstanceImage;
            prog.OverWriteFlag = Properties.Settings.Default.Overwrite;
            return prog;
        }

        private void CreateFromCasOnlineItem_Click(object sender, RoutedEventArgs e, bool append)
        {
            AA_Click(sender, e,
                E_rtf + "|" + E_all,
                append ? null : E_cfx + "|" + E_all, ".cfx",
                (sourcePaths, targetPath) =>
                {
                    var prog = CreateProgram();
                    prog.AppendFlag = append;
                    prog.OutputPath = targetPath;
                    prog.CasOnlineFiles = sourcePaths.ToList();
                    return prog;
                }, append);
        }

        private void AppendFromSciFinderItem_Click(object sender, RoutedEventArgs e)
        {
            CreateFromSciFinderItem_Click(sender, e, true);
        }

        private void CreateFromSciFinderItem_Click(object sender, RoutedEventArgs e)
        {
            CreateFromSciFinderItem_Click(sender, e, false);
        }

        private void CreateFromSciFinderItem_Click(object sender, RoutedEventArgs e, bool append)
        {
            AA_Click(sender, e,
                E_all_supported + "|" + E_rtf + "|" + E_sdf + "|" + E_csv + "|" +  E_all,
                append ? null : E_cfx + "|" + E_all, ".cfx",
                (sourcePaths, targetPath) =>
                {
                    var prog = CreateProgram();
                    prog.AppendFlag = append;
                    prog.OutputPath = targetPath;
                    prog.InputPaths = sourcePaths;
                    return prog;
                }, append);
        }

        private void AppendFromLST_Click(object sender, RoutedEventArgs e)
        {
            CreateFromLST_Click(sender, e, true);
        }

        private void CreateFromLST_Click(object sender, RoutedEventArgs e)
        {
            CreateFromLST_Click(sender, e, false);
        }

        private void CreateFromLST_Click(object sender, RoutedEventArgs e, bool append)
        {
            AA_Click(sender, e,
                E_lst + "|" + E_all,
                append ? null : E_cfx + "|" + E_all, ".cfx",
                (sourcePaths, targetPath) =>
                {
                    var prog = CreateProgram();
                    prog.AppendFlag = append;
                    prog.OutputPath = targetPath;
                    prog.InputPaths = sourcePaths;
                    return prog;
                }, append);
        }

        private void CreateFromCFX_Click(object sender, RoutedEventArgs e)
        {
            CreateFromCFX_Click(sender, e, false);
        }

        private void AppendFromCFX_Click(object sender, RoutedEventArgs e)
        {
            CreateFromCFX_Click(sender, e, true);
        }

        private void CreateFromCFX_Click(object sender, RoutedEventArgs e, bool append)
        {
            AA_Click(sender, e,
                E_cfx + "|" + E_all,
                append ? null : E_cfx + "|" + E_all, ".cfx",
                (sourcePaths, targetPath) =>
                {
                    var prog = CreateProgram();
                    prog.AppendFlag = append;
                    prog.OutputPath = targetPath;
                    prog.InputPaths = sourcePaths;
                    return prog;
                }, append);
        }

        private void CreateFromImage_Click(object sender, RoutedEventArgs e)
        {
            CreateFromImage_Click(sender, e, false);
        }

        private void AppendFromImage_Click(object sender, RoutedEventArgs e)
        {
            CreateFromImage_Click(sender, e, true);
        }

        private void CreateFromImage_Click(object sender, RoutedEventArgs e, bool append)
        {
            AA_Click(sender, e,
                E_image + "|" + E_all,
                append ? null : E_cfx + "|" + E_all, ".cfx",
                (sourcePaths, targetPath) =>
                {
                    var prog = CreateProgram();
                    prog.AppendFlag = append;
                    prog.OutputPath = targetPath;
                    prog.ImageFiles = sourcePaths.ToList();
                    return prog;
                }, append);
        }

        private void KillSelected_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in MainListViewItems().Where(l => l.IsSelected))
            {
                KillProcess(item);
            }
        }

        private static void KillProcess(ListItemViewModel item)
        {
            item.Thread.Abort();
            item.Thread.Join();
            item.TaskID = null;
            item.Action = null;
            item.Progress = "Killed";
        }

        private void DeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            var itemsToDelete = new List<ListItemViewModel>();
            foreach (var item in MainListViewItems().Where(l => l.IsSelected && !l.Thread.IsAlive))
            {
                itemsToDelete.Add(item);
            }

            foreach (var item in itemsToDelete)
            {
                MainListView.Items.Remove(item);
            }
        }

        private void OpenSelected_Click(object sender, RoutedEventArgs e)
        {
            var item = MainListView.SelectedItem as ListItemViewModel;
            if (item == null)
                return;

            if (File.Exists(item.Target))
            {
                string a_cfx = item.Target;
                string a_doc_cfx;

                var fileNameWOExt = System.IO.Path.GetFileNameWithoutExtension(a_cfx);
                var ext = System.IO.Path.GetExtension(a_cfx);
                if (fileNameWOExt.EndsWith("_doc"))
                {
                    a_doc_cfx = a_cfx;
                    a_cfx = fileNameWOExt.Substring(0, fileNameWOExt.Length - "_doc".Length) + ext;
                }
                else
                {
                    a_doc_cfx = fileNameWOExt + "_doc" + ext;
                }
                try
                {
                    System.Diagnostics.Process.Start(a_cfx);
                }
                catch (Exception)
                {
                }
                try
                {
                    System.Diagnostics.Process.Start(a_doc_cfx);
                }
                catch (Exception)
                {
                }
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SetSettings_Click(object sender, RoutedEventArgs e)
        {
            var win = new CfxManipurateSettingsWindow();
            win.SetValuedFromSetting();

            if (win.ShowDialog() != false)
            {
                win.SetSettingFromDialog();
            }
        }

        private void AA_Click(object sender, RoutedEventArgs e,
            string openFileDialogFilter, string saveFileDialogFilter, string defaultSaveExt,
            Func<IEnumerable<string>, string, IRunnable> actionCreator,
            bool isManipurate)
        {
            string targetPath = null;
            string[] sourcePaths = null;

            if (openFileDialogFilter != null)
            {
                OpenFileDialog openFileDialog = null;
                openFileDialog = GUtility.CreateOpenFileDialog(openFileDialogFilter);
                if (openFileDialog.ShowDialog() == false)
                    return;
                sourcePaths = openFileDialog.FileNames;
            }

            if (saveFileDialogFilter != null)
            {
                SaveFileDialog saveFileDialog = null;
                saveFileDialog = GUtility.CreateSaveFileDialog(saveFileDialogFilter, defaultSaveExt);
                if (saveFileDialog.ShowDialog() == false)
                    return;
                targetPath = saveFileDialog.FileName;
            }

            var item = MainListView.SelectedItem as ListItemViewModel;
            {
                if (targetPath == null && item != null && File.Exists(item.Target))
                    targetPath = item.Target;
            }
            if (!isManipurate)
                item = new ListItemViewModel();

            if (item == null)
                return; // not selected

            {
                SynchronizationContext ctx = SynchronizationContext.Current;
                var action = actionCreator(sourcePaths, targetPath);
                
                var thread = new Thread(
                    new ThreadStart(
                        () =>
                        {
                            var mes = "Finished";
                            try
                            {
                                action.Run();
                            }
                            catch (Exception ee)
                            {
                                mes = ee.Message;
                            }
                            ctx.Post(
                                state =>
                                {
                                    item.TaskID = null;
                                    item.Progress = mes;
                                }, null);
                        }));
                item.Thread = thread;
                item.Action = action;
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();

                item.TaskID = thread.ManagedThreadId;
                item.Target = targetPath;
                item.Progress = "Running";
            }
            if (!isManipurate)
                MainListView.Items.Add(item);
        }

        private void AssignLocalCodeNumber_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null, 
                (dummy, target) =>
                {
                    var prog = new Ujihara.Chemistry.CfxUtility.Program();
                    prog.CfxFilePath = target;
                    prog.LocalDatabasePath = Properties.Settings.Default.LocalDbPath;
                    prog.FieldName_LocalID_Input = Properties.Settings.Default.LocalCodeInLocalDb_FieldName;
                    return prog;
                }, true);
        }
        
        private void AddFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = GUtility.CreateOpenFileDialog(E_cfx + "|" + E_all);
            if (openFileDialog.ShowDialog() != false)
            {
                var item = new ListItemViewModel();
                item.Target = openFileDialog.FileName;
                MainListView.Items.Add(item);
            }
        }

        private void GenerateStructureFromName_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null, 
                (dummy, target) =>
                {
                    var prog = new Ujihara.Chemistry.CfxUtility.Program();
                    prog.CfxFilePath = target;
                    prog.GenerateStructureFlag = true;
                    SetProgSetting(prog);
                    return prog;
                }, true);
        }

        private void GenerateStructureFromImage_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null,
                (dummy, target) =>
                {
                    var prog = new Ujihara.Chemistry.CfxUtility.Program();
                    prog.CfxFilePath = target;
                    prog.GenerateStructureFromImage = true;
                    SetProgSetting(prog);
                    return prog;
                }, true);
        }

        private void ScaffordStructure_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null,
                (dummy, target) =>
                {
                    var scaffordCdx = Properties.Settings.Default.ScaffordCdx_FileName ?? "";
                    if (scaffordCdx != "")
                    {
                        var prog = new Ujihara.Chemistry.CfxUtility.Program();
                        prog.CfxFilePath = target;
                        prog.ScaffordCdx = scaffordCdx;
                        SetProgSetting(prog);
                        return prog;
                    }
                    else
                        return null; // TODO: Empty is better than null 
                }, true);
        }

        private void CleanupStructure_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null,
                (dummy, target) =>
                {
                    var prog = new Ujihara.Chemistry.CfxUtility.Program();
                    prog.CfxFilePath = target;
                    prog.CleanupStructureFlag = true;
                    SetProgSetting(prog);
                    return prog;
                }, true);
        }

        private void GenerateSmiles_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null,
                (dummy, target) =>
                {
                    var prog = new Ujihara.Chemistry.CfxUtility.Program();
                    prog.CfxFilePath = target;
                    prog.GenerateSmilesFlag = true;
                    SetProgSetting(prog);
                    return prog;
                }, true);
        }

        private void GenerateSructureFromInChi_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null,
                (dummy, target) =>
                {
                    var prog = new Ujihara.Chemistry.CfxUtility.Program();
                    prog.CfxFilePath = target;
                    prog.GenerateStructureFromInChi = true;
                    SetProgSetting(prog);
                    return prog;
                }, true);
        }

        private void GenerateSructureFromSmiles_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null,
                (dummy, target) =>
                {
                    var prog = new Ujihara.Chemistry.CfxUtility.Program();
                    prog.CfxFilePath = target;
                    prog.GenerateStructureFromSmiles = true;
                    SetProgSetting(prog);
                    return prog;
                }, true);
        }

        private void FillNoStructure_Click(object sender, RoutedEventArgs e)
        {
            AA_Click(sender, e, null, null, null,
                (dummy, target) =>
                {
                    var prog = new Ujihara.Chemistry.CfxUtility.Program();
                    prog.CfxFilePath = target;
                    prog.FillNoStructure = true;
                    SetProgSetting(prog);
                    return prog;
                }, true);
        }

        private static void SetProgSetting(Ujihara.Chemistry.CfxUtility.Program prog)
        {
            prog.SetChemNameFieldNames(Properties.Settings.Default.ChemName_FieldName);
        }
    
        private void HelpAbout_Click(object sender, RoutedEventArgs e)
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            var name = asm.GetName();
            MessageBox.Show(name.Name + " " + name.Version);
        }

        private System.Threading.Timer RefreshTimer { get; set; }

        private void MainListView_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshTimer = new System.Threading.Timer(Refresh, null, IntervalOfRefresh, IntervalOfRefresh);
        }

        private void MainListView_Unloaded(object sender, RoutedEventArgs e)
        {
            if (RefreshTimer != null)
                RefreshTimer.Dispose();
        }

        private void Refresh(object state)
        {
            foreach (var item in MainListViewItems())
            {
                if (item.Thread != null && 
                    item.Thread.ThreadState == System.Threading.ThreadState.Running && 
                    item.Action is IProgressReporter)
                {
                    var reporter = (IProgressReporter)item.Action;
                    item.Progress = reporter.GetReport();
                }
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (MainListViewItems()
                .Where(item => item.Thread.ThreadState == System.Threading.ThreadState.Running).Any())
            {
                var result = MessageBox.Show("Some precesses are running. Are you sure to stop them?", "Caution", MessageBoxButton.OKCancel, MessageBoxImage.Question, MessageBoxResult.Cancel);
                if (result == MessageBoxResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
            }

            foreach (var item in MainListViewItems()
                .Where(item => item.Thread.ThreadState == System.Threading.ThreadState.Running))
            {
                KillProcess(item);
            }
        }

        private IEnumerable<ListItemViewModel> MainListViewItems()
        {
            return MainListView.Items.Cast<ListItemViewModel>();
        }
    }

    public class ListItemViewModel 
        : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public Thread Thread { get; set; }
        public IRunnable Action { get; set; }        

        private bool isSelected;
        public bool IsSelected
        {
            get { return isSelected; }
            set 
            {
                if (value != isSelected)
                {
                    isSelected = value;
                    if (PropertyChanged != null) 
                        PropertyChanged(this, new PropertyChangedEventArgs("IsSelected"));
                }
            }
        }

        private int? taskID;
        public int? TaskID
        {
            get { return taskID; }
            set
            {
                if (value != taskID)
                {
                    taskID = value;
                    if (PropertyChanged != null)
                        PropertyChanged(this, new PropertyChangedEventArgs("TaskID"));
                }
            }
        }

        public string TargetFileName
        {
            get { return System.IO.Path.GetFileName(Target); }
        }

        private string target;
        public string Target
        {
            get { return target; }
            set
            {
                if (value != target)
                {
                    target = value;
                    if (PropertyChanged != null)
                    {
                        PropertyChanged(this, new PropertyChangedEventArgs("Target"));
                        PropertyChanged(this, new PropertyChangedEventArgs("TargetFileName"));
                    }
                }
            }
        } 

        private string progress;
        public string Progress
        {
            get { return progress; }
            set
            {
                if (value != progress)
                {
                    progress = value;
                    if (PropertyChanged != null)
                        PropertyChanged(this, new PropertyChangedEventArgs("Progress"));
                }
            }
        }
    }
}
