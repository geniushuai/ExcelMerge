using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.IO;
using System.Windows;
using Prism.Mvvm;
using FastWpfGrid;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Behaviors;

namespace ExcelMerge.GUI.ViewModels
{
    public class DiffViewModel : BindableBase
    {
        private bool showLocationGridLine;
        public bool ShowLocationGridLine
        {
            get { return showLocationGridLine; }
            set { SetProperty(ref showLocationGridLine, value); }
        }

        private string srcPath;
        public string SrcPath
        {
            get { return srcPath; }
            set
            {
                if (srcPath == value)
                    return;
                SetProperty(ref srcPath, value);
                Settings.EMEnvironmentValue.Set("SRC", value);
                UpdateExecutableFlag();
            }
        }

        private string dstPath;
        public string DstPath
        {
            get { return dstPath; }
            set
            {
                if (dstPath == value)
                    return;
                SetProperty(ref dstPath, value);
                Settings.EMEnvironmentValue.Set("DST", value);
                UpdateExecutableFlag();
            }
        }

        private string targetPath;
        public string TargetPath
        {
            get { return targetPath; }
            set
            {
                if (targetPath == value)
                    return;
                SetProperty(ref targetPath, value);
                Settings.EMEnvironmentValue.Set("TARGET", value);
                UpdateExecutableFlag();
            }
        }

        private List<string> srcFileNames;
        public List<string> SrcFileNames
        {
            get { return srcFileNames; }
            private set { SetProperty(ref srcFileNames, value); }
        }

        private List<string> dstFileNames;
        public List<string> DstFileNames
        {
            get { return dstFileNames; }
            private set { SetProperty(ref dstFileNames, value); }
        }

        private List<string> targetFileNames;
        public List<string> TargetFileNames
        {
            get { return targetFileNames; }
            private set { SetProperty(ref targetFileNames, value); }
        }

        private List<string> srcSheetNames;
        public List<string> SrcSheetNames
        {
            get { return srcSheetNames; }
            private set { SetProperty(ref srcSheetNames, value); }
        }

        private List<string> dstSheetNames;
        public List<string> DstSheetNames
        {
            get { return dstSheetNames; }
            private set { SetProperty(ref dstSheetNames, value); }
        }

        private List<string> targetSheetNames;
        public List<string> TargetSheetNames
        {
            get { return targetSheetNames; }
            private set { SetProperty(ref targetSheetNames, value); }
        }

        private int selectedSrcFileIndex;
        public int SelectedSrcFileIndex
        {
            get { return selectedSrcFileIndex; }
            set { SetProperty(ref selectedSrcFileIndex, value); }
        }

        private int selectedDstFileIndex;
        public int SelectedDstFileIndex
        {
            get { return selectedDstFileIndex; }
            set { SetProperty(ref selectedDstFileIndex, value); }
        }

        private int selectedTargetFileIndex;
        public int SelectedTargetFileIndex
        {
            get { return selectedTargetFileIndex; }
            set { SetProperty(ref selectedTargetFileIndex, value); }
        }

        private int selectedSrcSheetIndex;
        public int SelectedSrcSheetIndex
        {
            get { return selectedSrcSheetIndex; }
            set { SetProperty(ref selectedSrcSheetIndex, value); }
        }

        private int selectedDstSheetIndex;
        public int SelectedDstSheetIndex
        {
            get { return selectedDstSheetIndex; }
            set { SetProperty(ref selectedDstSheetIndex, value); }
        }

        private int selectedTargetSheetIndex;
        public int SelectedTargetSheetIndex
        {
            get { return selectedTargetSheetIndex; }
            set { SetProperty(ref selectedTargetSheetIndex, value); }
        }

        private bool executable;
        public bool Executable
        {
            get { return executable; }
            private set { SetProperty(ref executable, value); }
        }

        private int modifiedCellCount;
        public int ModifiedCellCount
        {
            get { return modifiedCellCount; }
            private set { SetProperty(ref modifiedCellCount, value); }
        }

        private int modifiedRowCount;
        public int ModifiedRowCount
        {
            get { return modifiedRowCount; }
            private set { SetProperty(ref modifiedRowCount, value); }
        }

        private int addedRowCount;
        public int AddedRowCount
        {
            get { return addedRowCount; }
            private set { SetProperty(ref addedRowCount, value); }
        }

        private int removedRowCount;
        public int RemovedRowCount
        {
            get { return removedRowCount; }
            private set { SetProperty(ref removedRowCount, value); }
        }

        private DragAcceptDescription description;
        public DragAcceptDescription Description
        {
            get { return description; }
            private set { SetProperty(ref description, value); }
        }

        public DiffViewModel()
        {
            Description = new DragAcceptDescription();
            Description.DragDrop += DragDrop;
            Description.DragDrop += DragOver;

            SrcPath = string.Empty;
            DstPath = string.Empty;
        }

        public DiffViewModel(string src, string dst, MainWindowViewModel mwv) : this()
        {
            SrcPath = src;
            DstPath = dst;

            mwv.PropertyChanged += Mwv_PropertyChanged;
        }

        public void UpdateDiffSummary(ExcelSheetDiffSummary summary)
        {
            ModifiedCellCount = summary.ModifiedCellCount;
            ModifiedRowCount = summary.ModifiedRowCount;
            AddedRowCount = summary.AddedRowCount;
            RemovedRowCount = summary.RemovedRowCount;
        }

        private void Mwv_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SrcPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        SrcPath = prop.GetValue(vm) as string;
                    }
                }
            }
            else if (e.PropertyName == nameof(DstPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        DstPath = prop.GetValue(vm) as string;
                    }
                }
            }
            else if (e.PropertyName == nameof(TargetPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        TargetPath = prop.GetValue(vm) as string;
                    }
                }
            }
        }

        private void DragDrop(DragEventArgs e)
        {
            var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null || !paths.Any())
                return;

            var target = e.Source as FrameworkElement;
            if (target == null)
                return;

            OnDragDrop(paths, target);
        }

        protected virtual void OnDragDrop(string[] filePath, FrameworkElement target)
        {
            if (filePath.Length > 2)
            {
                SrcPath = filePath[1];
                DstPath = filePath[0];
                TargetPath = filePath[2];
                return;
            }

            var tag = Convert.ToInt32(target.Tag);
            if (tag == 0)
                SrcPath = filePath[0];
            else if (tag == 1)
                DstPath = filePath[0];
            else if (tag == 11)
                TargetPath = filePath[0];
        }

        private void DragOver(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
        }
        private string preSrcPath = "";
        private string preDstPath = "";
        private string preTargetPath = "";
        private void UpdateExecutableFlag()
        {
            var existsSrc = Directory.Exists(SrcPath);
            var existsDst = Directory.Exists(DstPath);
            var existTarget = Directory.Exists(TargetPath);

            if (existsSrc)
            {
                if (preSrcPath != SrcPath)
                {
                    preSrcPath = SrcPath;
                    DirectoryInfo dir = new DirectoryInfo(SrcPath);
                    var t = new List<string>();
                    foreach (FileInfo f  in dir.GetFiles())
                    {
                        t.Add(f.ToString());
                    }
                    SrcFileNames = t;
                    SelectedSrcFileIndex = 0;
                }
            }
            else
            {
                SrcSheetNames = new List<string>();
                SelectedSrcSheetIndex = -1;
                SrcFileNames = new List<string>();
                SelectedSrcFileIndex = -1;
            }

            if (existsDst)
            {
                if (preDstPath != DstPath)
                {
                    preDstPath = DstPath;
                    DirectoryInfo dir = new DirectoryInfo(DstPath);
                    var t = new List<string>();
                    foreach (FileInfo f in dir.GetFiles())
                    {
                        t.Add(f.ToString());
                    }
                    DstFileNames = t;
                    SelectedDstFileIndex = 0;
                }
            }
            else
            {
                DstSheetNames = new List<string>();
                SelectedDstSheetIndex = -1;
                DstFileNames = new List<string>();
                SelectedDstFileIndex = -1;
            }

            if (existTarget )
            {
                if (preTargetPath != TargetPath)
                {
                    preTargetPath = TargetPath;
                    DirectoryInfo dir = new DirectoryInfo(TargetPath);
                    var t = new List<string>();
                    foreach (FileInfo f in dir.GetFiles())
                    {
                        t.Add(f.ToString());
                    }
                    TargetFileNames = t;
                    SelectedTargetFileIndex = 0;
                }
            }
            else
            {
                TargetSheetNames = new List<string>();
                SelectedTargetSheetIndex = -1;
                TargetFileNames = new List<string>();
                SelectedTargetFileIndex = -1;
            }
            Executable = existsSrc && existsDst && existTarget;
        }
    }
}
