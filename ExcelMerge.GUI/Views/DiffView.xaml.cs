using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Practices.Unity;
using FastWpfGrid;
using NetDiff;
using SKCore.Collection;
using ExcelMerge.GUI.ViewModels;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Models;
using ExcelMerge.GUI.Styles;
using ExcelMerge;



namespace ExcelMerge.GUI.Views
{
    public partial class DiffView : UserControl
    {
        private ExcelSheetDiffConfig diffConfig = new ExcelSheetDiffConfig();
        private IUnityContainer container;
        private const string srcKey = "src";
        private const string dstKey = "dst";
        private const string targetKey = "target";
        private ExcelHelperCom targetExcel = null;
        private FastGridControl copyTargetGrid;

        public DiffView()
        {
            InitializeComponent();
            InitializeContainer();
            InitializeEventListeners();

            App.Instance.OnSettingUpdated += OnApplicationSettingUpdated;

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            // In order to enable Ctrl + F immediately after startup.
            ToolExpander.IsExpanded = true;
        }

        private DiffViewModel GetViewModel()
        {
            return DataContext as DiffViewModel;
        }       
        public string RealSrcFilePath
        {
            get {
                if (!Directory.Exists(SrcPathTextBox.Text) || SrcFileCombobox.SelectedIndex < 0)                
                    return "";
                string srcFile = Path.Combine(SrcPathTextBox.Text, SrcFileCombobox.SelectedValue.ToString());
                if (!File.Exists(srcFile))
                    return "";
                return srcFile;

            }
        }
        public string RealDstFilePath
        {
            get
            {
                if (!Directory.Exists(DstPathTextBox.Text) || DstFileCombobox.SelectedIndex < 0)
                    return "";
                string dstFile = Path.Combine(DstPathTextBox.Text, DstFileCombobox.SelectedValue.ToString());
                if (!File.Exists(dstFile))
                    return "";
                return dstFile;

            }
        }
        public string RealTargetFilePath
        {
            get
            {
                if (!Directory.Exists(TargetPathTextBox.Text) || TargetFileCombobox.SelectedIndex < 0)
                    return "";
                string targetFile = Path.Combine(TargetPathTextBox.Text, TargetFileCombobox.SelectedValue.ToString());
                if (!File.Exists(targetFile))
                    return "";
                return targetFile;

            }
        }

        private void InitializeContainer()
        {
            container = new UnityContainer();
            container
                .RegisterInstance(srcKey, SrcDataGrid)
                .RegisterInstance(dstKey, DstDataGrid)                
                .RegisterInstance(srcKey, SrcLocationGrid)
                .RegisterInstance(dstKey, DstLocationGrid)                
                .RegisterInstance(srcKey, SrcViewRectangle)
                .RegisterInstance(dstKey, DstViewRectangle)                
                .RegisterInstance(srcKey, SrcValueTextBox)
                .RegisterInstance(dstKey, DstValueTextBox)
                .RegisterInstance(srcKey, SrcPathTextBox)
                .RegisterInstance(dstKey, DstPathTextBox);
        }

        private void InitializeEventListeners()
        {
            var srcEventHandler = new DiffViewEventHandler(srcKey);
            var dstEventHandler = new DiffViewEventHandler(dstKey);
            var targetEventHandler = new DiffViewEventHandler(targetKey);

            DataGridEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            DataGridEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            DataGridEventDispatcher.Instance.Listeners.Add(targetEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(targetEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(targetEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(targetEventHandler);
        }

        private void OnApplicationSettingUpdated()
        {
            var e = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchApplicationSettingUpdateEvent(e);
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchParentLoadEvent(args);

            ExecuteDiff(isStartup: true);

            // In order to enable Ctrl + F immediately after startup.
            ToolExpander.IsExpanded = false;
        }

        private ExcelSheetDiffConfig CreateDiffConfig(FileSetting srcFileSetting, FileSetting dstFileSetting, FileSetting targetFileSetting, bool isStartup)
        {
            var config = new ExcelSheetDiffConfig();

            config.SrcSheetIndex = SrcSheetCombobox.SelectedIndex;
            config.DstSheetIndex = DstSheetCombobox.SelectedIndex;
            config.TargetSheetIndex = TargetSheetCombobox.SelectedIndex;

            if (srcFileSetting != null)
            {
                if (isStartup)
                    config.SrcSheetIndex = GetSheetIndex(srcFileSetting, SrcSheetCombobox.Items);

                config.SrcHeaderIndex = srcFileSetting.ColumnHeaderIndex;
            }

            if (dstFileSetting != null)
            {
                if (isStartup)
                    config.DstSheetIndex = GetSheetIndex(dstFileSetting, DstSheetCombobox.Items);

                config.DstHeaderIndex = dstFileSetting.ColumnHeaderIndex;
            }

            if (targetFileSetting != null)
            {
                if (isStartup)
                    config.TargetSheetIndex = GetSheetIndex(targetFileSetting, TargetSheetCombobox.Items);

                config.TargetHeaderIndex = targetFileSetting.ColumnHeaderIndex;
            }
            return config;
        }

        private int GetSheetIndex(FileSetting fileSetting, ItemCollection sheetNames)
        {
            if (fileSetting == null)
                return -1;

            var index = fileSetting.SheetIndex;
            if (!string.IsNullOrEmpty(fileSetting.SheetName))
                index = sheetNames.IndexOf(fileSetting.SheetName);

            if (index < 0 || index >= sheetNames.Count)
            {
                MessageBox.Show(Properties.Resources.Msg_OutofSheetRange);
                index = 0;
            }

            return index;
        }

        private void LocationGrid_MouseDown(object sender, MouseEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
        }

        private void LocationGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
                LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
            }
        }

        private void LocationGrid_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseWheelEvent(args, e);
        }

        private void DataGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }

        private void LocationGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }
        private void DataGrid_DoubleClickCell(object sender, FastWpfGrid.DoubleClickEventArgs e)
        {
            var grid = copyTargetGrid = sender as FastGridControl;
            if (grid == null)
                return;            
            var srcRowHeaderText = (SrcDataGrid.Model as DiffGridModel).GetRowHeaderText(SrcDataGrid.CurrentCell.Row.Value);
            var srcColHeaderText = (SrcDataGrid.Model as DiffGridModel).GetColumnHeaderText(SrcDataGrid.CurrentCell.Column.Value);
            var srcRow = SrcDataGrid.CurrentCell.Row;
            var srcCol = SrcDataGrid.CurrentCell.Column;
            var dstRowHeaderText = (DstDataGrid.Model as DiffGridModel).GetRowHeaderText(DstDataGrid.CurrentCell.Row.Value);
            var dstColHeaderText = (DstDataGrid.Model as DiffGridModel).GetColumnHeaderText(DstDataGrid.CurrentCell.Column.Value);
            var dstRow = DstDataGrid.CurrentCell.Row;
            var dstCol = DstDataGrid.CurrentCell.Column;
            if (grid.Name == "SrcDataGrid")
            {
                TryToSelectTargetCell(srcRowHeaderText, srcColHeaderText, srcRow.Value, srcCol.Value);
            }
            else if (grid.Name == "DstDataGrid")
            {
                TryToSelectTargetCell(dstRowHeaderText, dstColHeaderText, dstRow.Value, dstCol.Value);
            }
        }
        private void DataGrid_SelectedCellsChanged(object sender, FastWpfGrid.SelectionChangedEventArgs e)
        {
            var grid = copyTargetGrid = sender as FastGridControl;
            if (grid == null)
                return;

                       
            var name = grid.Name;
            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSelectedCellChangeEvent(args);
            copyTargetGrid = grid;
            if (!SrcDataGrid.CurrentCell.Row.HasValue || !DstDataGrid.CurrentCell.Row.HasValue)
                return;

            if (!SrcDataGrid.CurrentCell.Column.HasValue || !DstDataGrid.CurrentCell.Column.HasValue)
                return;

            if (SrcDataGrid.Model == null || DstDataGrid.Model == null)
                return;

            var srcValue =
                (SrcDataGrid.Model as DiffGridModel).GetCellText(SrcDataGrid.CurrentCell.Row.Value, SrcDataGrid.CurrentCell.Column.Value, true);
            var dstValue =
                (DstDataGrid.Model as DiffGridModel).GetCellText(DstDataGrid.CurrentCell.Row.Value, DstDataGrid.CurrentCell.Column.Value, true);            
            UpdateValueDiff(srcValue, dstValue);

            if (App.Instance.Setting.AlwaysExpandCellDiff)
            {
                var a = new DiffViewEventArgs<RichTextBox>(null, container, TargetType.First);
                ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(a);
            }
        }

        private void TryToSelectTargetCell(string rowHeaderText, string colHeaderText, int rowIndex, int colIndex)
        {
            if (targetExcel == null)
                return;
            var rowResult = targetExcel.GetRowByHeaderText(rowHeaderText, rowIndex);
            var colResult = targetExcel.GetColByHeaderText(colHeaderText, colIndex);
            if (rowResult.Item2 < 0 || colResult.Item2 < 0)
                return;
            targetExcel.ActiveCell(rowResult.Item2, colResult.Item2);
        }

        private void ValueTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container, TargetType.First);
            ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(args);
        }

        private void ValueTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container, TargetType.First);
            ValueTextBoxEventDispatcher.Instance.DispatchLostFocusEvent(args);
        }

        private string GetRichTextString(RichTextBox textBox)
        {
            var textRange = new TextRange(textBox.Document.ContentStart, textBox.Document.ContentEnd);

            return textRange.Text;
        }

        private IEnumerable<DiffResult<string>> DiffCellValue(IEnumerable<string> src, IEnumerable<string> dst)
        {
            var r = DiffUtil.Diff(src, dst);
            r = DiffUtil.Order(r, DiffOrderType.LazyDeleteFirst);
            return DiffUtil.OptimizeCaseDeletedFirst(r);
        }

        private string ConvertWhiteSpaces(string str)
        {
            return new string(str.Select(c =>
            {
                if (Encoding.UTF8.GetByteCount(c.ToString()) == 1)
                    return ' ';
                else
                    return '　';

            }).ToArray());
        }

        private string ConvertWhiteSpaces(char c)
        {
            if (Encoding.UTF8.GetByteCount(c.ToString()) == 1)
                return " ";
            else
                return "　";
        }

        private void DiffModifiedLine(IEnumerable<DiffResult<char>> results, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var splited = results.SplitByRegularity((items, current) => items.Last().Status.Equals(current.Status)).ToList();

            foreach (var sr in splited)
            {
                var status = sr.First().Status;
                if (status == DiffStatus.Equal)
                {
                    ranges.Add(Tuple.Create<string, Color?>(new string(sr.Select(r => r.Obj1).ToArray()), null));
                }
                else if (status == DiffStatus.Modified)
                {
                    var str = new string(sr.Select(r => isSrc ? r.Obj1 : r.Obj2).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.LightOrange));
                }
                else if (status == DiffStatus.Deleted)
                {
                    var str = new string(sr.Select(r => r.Obj1).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.LightGray));
                }
                else if (status == DiffStatus.Inserted)
                {
                    var str = new string(sr.Select(r => r.Obj2).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.Orange));
                }
            }

            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffEqualLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges)
        {
            ranges.Add(Tuple.Create<string, Color?>(lineDiffResult.Obj1, null));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffDeletedLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var str = isSrc ? lineDiffResult.Obj1 : ConvertWhiteSpaces(lineDiffResult.Obj1.ToString());
            ranges.Add(Tuple.Create<string, Color?>(str, isSrc ? EMColor.LightGray : EMColor.LightGray));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffInsertedLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var str = isSrc ? ConvertWhiteSpaces(lineDiffResult.Obj2) : lineDiffResult.Obj2;
            ranges.Add(Tuple.Create<string, Color?>(str, isSrc ? EMColor.LightGray : EMColor.Orange));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void UpdateValueDiff(string srcValue, string dstValue)
        {
            SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();
            DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();

            var srcLines = srcValue.Split('\n').Select(s => s.TrimEnd());
            var dstLines = dstValue.Split('\n').Select(s => s.TrimEnd());

            var lineDiffResults = DiffCellValue(srcLines, dstLines).ToList();

            var srcRange = new List<Tuple<string, Color?>>();
            var dstRange = new List<Tuple<string, Color?>>();
            foreach (var lineDiffResult in lineDiffResults)
            {
                if (lineDiffResult.Status == DiffStatus.Equal)
                {
                    DiffEqualLine(lineDiffResult, srcRange);
                    DiffEqualLine(lineDiffResult, dstRange);
                }
                else if (lineDiffResult.Status == DiffStatus.Modified)
                {
                    var charDiffResults = DiffUtil.Diff(lineDiffResult.Obj1, lineDiffResult.Obj2);
                    charDiffResults = DiffUtil.Order(charDiffResults, DiffOrderType.LazyDeleteFirst);
                    charDiffResults = DiffUtil.OptimizeCaseDeletedFirst(charDiffResults);

                    DiffModifiedLine(charDiffResults.Where(r => r.Status != DiffStatus.Inserted), srcRange, true);
                    DiffModifiedLine(charDiffResults.Where(r => r.Status != DiffStatus.Deleted), dstRange, false);
                }
                else if (lineDiffResult.Status == DiffStatus.Deleted)
                {
                    DiffDeletedLine(lineDiffResult, srcRange, true);
                    DiffDeletedLine(lineDiffResult, dstRange, false);
                }
                else if (lineDiffResult.Status == DiffStatus.Inserted)
                {
                    DiffInsertedLine(lineDiffResult, srcRange, true);
                    DiffInsertedLine(lineDiffResult, dstRange, false);
                }
            }

            foreach (var r in srcRange)
            {
                var bc = r.Item2.HasValue ? new SolidColorBrush(r.Item2.Value) : new SolidColorBrush();
                SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Add(new Run(r.Item1) { Background = bc });
            }

            foreach (var r in dstRange)
            {
                var bc = r.Item2.HasValue ? new SolidColorBrush(r.Item2.Value) : new SolidColorBrush();
                DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Add(new Run(r.Item1) { Background = bc });
            }
        }

        private void DiffButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteDiff();
        }

        private ExcelSheetReadConfig CreateReadConfig()
        {
            var setting = ((App)Application.Current).Setting;

            return new ExcelSheetReadConfig()
            {
                TrimFirstBlankRows = setting.SkipFirstBlankRows,
                TrimFirstBlankColumns = setting.SkipFirstBlankColumns,
                TrimLastBlankRows = setting.TrimLastBlankRows,
                TrimLastBlankColumns = setting.TrimLastBlankColumns,
            };
        }

        private Dictionary<string, int> ReadTargetWorkBooks(int targetSheetIndex = 0)
        {
            var config = CreateReadConfig();
            Dictionary<string, int> rowKeyCache = new Dictionary<string, int>();
            ExcelWorkbook twb = ExcelWorkbook.Create(RealTargetFilePath, config, targetSheetIndex);
            int continueBlankCount = 0;
            foreach(var rowInfo in twb.Sheets[twb.Sheets.Keys.First()].Rows)
            {
                int rowIndex = rowInfo.Key;
                ExcelRow row = rowInfo.Value;
                if(row.Cells.Count>0)
                {
                    string data = row.Cells[0].Value.ToString().Trim();
                    if(data == "")
                    {
                        continueBlankCount++;
                    }
                    else
                    {
                        continueBlankCount = 0;
                    }
                    if (!rowKeyCache.ContainsKey(data))
                    {
                        rowKeyCache[data] = rowIndex;
                    }
                }
                if (continueBlankCount > ExcelHelperCom.ContinueRowBlankUpperLimit)
                    break;
            }
            return rowKeyCache;                        
        }
        private Tuple<ExcelWorkbook, ExcelWorkbook> ReadWorkbooks(int srcSheetIndex = 0, int dstSheetIndex=0)
        {
            ExcelWorkbook swb = null;
            ExcelWorkbook dwb = null;
            var srcPath = RealSrcFilePath;
            var dstPath = RealDstFilePath;            
            ProgressWindow.DoWorkWithModal(progress =>
            {
                progress.Report(Properties.Resources.Msg_ReadingFiles);

                var config = CreateReadConfig();
                /*
                config.TrimFirstBlankColumns = true;
                config.TrimLastBlankColumns = true;
                config.TrimFirstBlankRows = true;
                config.TrimLastBlankRows = true;
                */
                swb = ExcelWorkbook.Create(srcPath, config, srcSheetIndex);
                dwb = ExcelWorkbook.Create(dstPath, config, dstSheetIndex);                
            });

            return Tuple.Create(swb, dwb);
        }

        private Tuple<FileSetting, FileSetting> FindFileSettings(bool isStartup)
        {
            FileSetting srcSetting = null;
            FileSetting dstSetting = null;
            FileSetting targetSetting = null;
            var srcPath = SrcPathTextBox.Text;
            var dstPath = DstPathTextBox.Text;
            var targetPath = TargetPathTextBox.Text;
            if (!IgnoreFileSettingCheckbox.IsChecked.Value)
            {
                srcSetting =
                    FindFilseSetting(Path.GetFileName(srcPath), SrcSheetCombobox.SelectedIndex, SrcSheetCombobox.SelectedItem.ToString(), isStartup);

                dstSetting =
                    FindFilseSetting(Path.GetFileName(dstPath), DstSheetCombobox.SelectedIndex, DstSheetCombobox.SelectedItem.ToString(), isStartup);

                targetSetting =
                    FindFilseSetting(Path.GetFileName(targetPath), TargetSheetCombobox.SelectedIndex, TargetSheetCombobox.SelectedItem.ToString(), isStartup);

                diffConfig = CreateDiffConfig(srcSetting, dstSetting, targetSetting, isStartup);
            }
            else
            {
                diffConfig = new ExcelSheetDiffConfig();

                diffConfig.SrcSheetIndex = Math.Max(SrcSheetCombobox.SelectedIndex, 0);
                diffConfig.DstSheetIndex = Math.Max(DstSheetCombobox.SelectedIndex, 0);
                diffConfig.TargetSheetIndex = Math.Max(TargetSheetCombobox.SelectedIndex, 0);
            }

            return Tuple.Create(srcSetting, dstSetting);
        }

        private ExcelSheetDiff ExecuteDiff(ExcelSheet srcSheet, ExcelSheet dstSheet)
        {
            ExcelSheetDiff diff = null;
            ProgressWindow.DoWorkWithModal(progress =>
            {
                progress.Report(Properties.Resources.Msg_ExtractingDiff);
                diff = ExcelSheet.Diff(srcSheet, dstSheet, diffConfig);
            });

            return diff;
        }

        private void ExecuteDiff(bool isStartup = false)
        {
                        
            if (!File.Exists(RealSrcFilePath) || !File.Exists(RealDstFilePath) || !File.Exists(RealTargetFilePath))
                return;                
            if (targetExcel != null)
            {
                try { 
                    targetExcel.ExitApp();
                }
                catch
                {

                }
            targetExcel = null;        
            }
            Dictionary<string, int> keyCache =  ReadTargetWorkBooks();
            targetExcel = new ExcelHelperCom();
            if(!targetExcel.OpenExcel(RealTargetFilePath, keyCache))
            {
                MessageBox.Show("打开excel失败");
                return;
            }
            targetExcel.ShowExcel();
            targetExcel.ActivateSheet(TargetSheetCombobox.SelectedIndex+1);
            targetExcel.ActiveCell(1, 1);           
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchPreExecuteDiffEvent(args);

            var workbooks = ReadWorkbooks(SrcSheetCombobox.SelectedIndex, DstSheetCombobox.SelectedIndex);
            var srcWorkbook = workbooks.Item1;
            var dstWorkbook = workbooks.Item2;            

            var fileSettings = FindFileSettings(isStartup);
            var srcFileSetting = fileSettings.Item1;
            var dstFileSetting = fileSettings.Item2;            

            SrcSheetCombobox.SelectedIndex = diffConfig.SrcSheetIndex;
            DstSheetCombobox.SelectedIndex = diffConfig.DstSheetIndex;
            TargetSheetCombobox.SelectedIndex = diffConfig.TargetSheetIndex;

            var srcSheet = srcWorkbook.Sheets[SrcSheetCombobox.SelectedItem.ToString()];
            var dstSheet = dstWorkbook.Sheets[DstSheetCombobox.SelectedItem.ToString()];
            //var targetSheet = targetWorkbook.Sheets[TargetSheetCombobox.SelectedItem.ToString()];

            if (srcSheet.Rows.Count > 50000 || dstSheet.Rows.Count > 50000 )
                MessageBox.Show(Properties.Resources.Msg_WarnSize);

            var diff = ExecuteDiff(srcSheet, dstSheet);
            SrcDataGrid.Model = new DiffGridModel(diff, DiffType.Source);
            SrcDataGrid.Model.SelectedRowCountLimit = 100;
            SrcDataGrid.Model.SelectedColumnCountLimit = 100;
            DstDataGrid.Model = new DiffGridModel(diff, DiffType.Dest);
            DstDataGrid.Model.SelectedRowCountLimit = 100;
            DstDataGrid.Model.SelectedColumnCountLimit = 100;

            args = new DiffViewEventArgs<FastGridControl>(SrcDataGrid, container);
            DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, srcFileSetting);

            args = new DiffViewEventArgs<FastGridControl>(DstDataGrid, container);
            DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, dstFileSetting);

            args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, ShowOnlyDiffRadioButton.IsChecked.Value);
            DataGridEventDispatcher.Instance.DispatchPostExecuteDiffEvent(args);

            var summary = diff.CreateSummary();
            GetViewModel().UpdateDiffSummary(summary);             

            //book.Activate = book.Sheets[TargetSheetCombobox.SelectedIndex];
            if (!App.Instance.KeepFileHistory)
                App.Instance.UpdateRecentFiles(SrcPathTextBox.Text, DstPathTextBox.Text, TargetPathTextBox.Text);

            if (App.Instance.Setting.NotifyEqual && !summary.HasDiff)
                MessageBox.Show(Properties.Resources.Message_NoDiff);

            if (App.Instance.Setting.FocusFirstDiff)
                MoveNextModifiedCell();
        }

        private FileSetting FindFilseSetting(string fileName, int sheetIndex, string sheetName, bool isStartup)
        {
            var results = new List<FileSetting>();
            foreach (var setting in App.Instance.Setting.FileSettings)
            {
                if (setting.UseRegex)
                {
                    var regex = new System.Text.RegularExpressions.Regex(setting.Name);

                    if (regex.IsMatch(fileName))
                        results.Add(setting);
                }
                else
                {
                    if (setting.ExactMatch)
                    {
                        if (setting.Name == fileName)
                            results.Add(setting);
                    }
                    else
                    {
                        if (fileName.Contains(setting.Name))
                            results.Add(setting);
                    }
                }
            }

            if (isStartup)
                return results.FirstOrDefault(r => r.IsStartupSheet) ?? results.FirstOrDefault() ?? null;

            return results.FirstOrDefault(r => r.SheetName == sheetName) ?? results.FirstOrDefault(r => r.SheetIndex == sheetIndex) ?? null;
        }

        private void SetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderChagneEvent(args);
                }
            }
        }

        private void ResetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderResetEvent(args);
                }
            }
        }

        private void SetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderChangeEvent(args);
                }
            }
        }

        private void ResetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderResetEvent(args);
                }
            }
        }

        private void SwapButton_Click(object sender, RoutedEventArgs e)
        {
            Swap();
        }

        private void Swap()
        {
            var srcTmp = SrcSheetCombobox.SelectedIndex;
            var dstTmp = DstSheetCombobox.SelectedIndex;

            var tmp = SrcPathTextBox.Text;
            SrcPathTextBox.Text = DstPathTextBox.Text;
            DstPathTextBox.Text = tmp;

            diffConfig.SrcSheetIndex = dstTmp;
            diffConfig.DstSheetIndex = srcTmp;

            ExecuteDiff();
        }

        private void DiffByHeaderSrc_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = SrcDataGrid.CurrentCell.Row.HasValue ? SrcDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.SrcHeaderIndex= headerIndex;

            ExecuteDiff();
        }

        private void DiffByHeaderDst_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = DstDataGrid.CurrentCell.Row.HasValue ? DstDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.DstSheetIndex = headerIndex;

            ExecuteDiff();
        }

        private void ShowAllRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, false);
        }

        private void ShowOnlyDiffRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, true);
        }

        private bool ValidateDataGrids()
        {
            return SrcDataGrid.Model != null && DstDataGrid.Model != null;
        }

        private void ValuteTextBox_ScrollChanged(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container);
            ValueTextBoxEventDispatcher.Instance.DispatchScrolledEvent(args, (ScrollChangedEventArgs)e);
        }

        private void NextModifiedCellButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextModifiedCell();
        }

        private void MoveNextModifiedCell()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextModifiedCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevModifiedCellButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevModifiedCell();
        }

        private void MovePrevModifiedCell()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousModifiedCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextModifiedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextModifiedRow();
        }

        private void MoveNextModifiedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextModifiedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevModifiedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevModifiedRow();
        }

        private void MovePrevModifiedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousModifiedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextAddedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextAddedRow();
        }

        private void MoveNextAddedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextAddedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevAddedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevAddedRow();
        }

        private void MovePrevAddedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousAddedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextRemovedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextRemovedRow();
        }

        private void MoveNextRemovedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextRemovedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevRemovedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevRemovedRow();
        }

        private void MovePrevRemovedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousRemovedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevMatchCellButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevMatchCell();
        }

        private void MovePrevMatchCell()
        {
            if (!ValidateDataGrids())
                return;

            var text = SearchTextCombobox.Text;
            if (string.IsNullOrEmpty(text))
                return;

            var history = App.Instance.Setting.SearchHistory.ToList();
            if (history.Contains(text))
                history.Remove(text);

            history.Insert(0, text);
            history = history.Take(10).ToList();

            App.Instance.Setting.SearchHistory = new ObservableCollection<string>(history);
            App.Instance.Setting.Save();

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousMatchCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell, text,
                ExactMatchCheckBox.IsChecked.Value, CaseSensitiveCheckBox.IsChecked.Value, RegexCheckBox.IsChecked.Value, ShowOnlyDiffRadioButton.IsChecked.Value);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextMatchCellButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextMatchCell();
        }

        private void MoveNextMatchCell()
        {
            if (!ValidateDataGrids())
                return;

            var text = SearchTextCombobox.Text;
            if (string.IsNullOrEmpty(text))
                return;

            var history = App.Instance.Setting.SearchHistory.ToList();
            if (history.Contains(text))
                history.Remove(text);

            history.Insert(0, text);
            history = history.Take(10).ToList();

            App.Instance.Setting.SearchHistory = new ObservableCollection<string>(history);
            App.Instance.Setting.Save();

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextMatchCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell, text,
                ExactMatchCheckBox.IsChecked.Value, CaseSensitiveCheckBox.IsChecked.Value, RegexCheckBox.IsChecked.Value, ShowOnlyDiffRadioButton.IsChecked.Value);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }
        private void CopyCurrentRow(FastGridControl copyTargetGrid, bool replace=false)
        {
            if (copyTargetGrid == null)
                return;

            var model = copyTargetGrid.Model as DiffGridModel;
            if (model == null)
                return;
            bool currentRowReplace = replace;
            foreach (var currentCell in copyTargetGrid.SelectedCells)
            {
                string colHeaderText = model.GetColumnHeaderText(currentCell.Column.Value);
                string rowHeaderText = model.GetRowHeaderText(currentCell.Row.Value);
                int rowCount = model.RowCount;
                int columnCount = model.ColumnCount;
                var rowIndexInTarget = targetExcel.GetRowByHeaderText(rowHeaderText, currentCell.Row.Value);
                if (rowIndexInTarget.Item1 == 1)
                {
                    //MessageBox.Show("第一列中已经存在：" + rowHeaderText + "将在其前一行插入复制内容");
                    targetExcel.ActiveCell(rowIndexInTarget.Item2, 1);
                    currentRowReplace = true && replace;                    
                }
                else
                {
                    //MessageBox.Show("第一列中:" + rowHeaderText + "不存在,将找最接近的前一行插入复制内容");
                    targetExcel.ActiveCell(rowIndexInTarget.Item2 + 1, 1);
                    currentRowReplace = false;
                }
                int newRowIndex = rowIndexInTarget.Item2;
                if (!currentRowReplace)
                    newRowIndex = targetExcel.InsertRow();
                for (int colIndex = 0; colIndex < columnCount; ++colIndex)
                {
                    string currentColHeaderText = model.GetColumnHeaderText(colIndex);
                    string currentCellText = model.GetCellText(currentCell.Row.Value, colIndex);
                    string targetCellText = targetExcel.ReadData(newRowIndex, colIndex + 1);
                    if(currentCellText != targetCellText)
                        targetExcel.WriteData(currentCellText, newRowIndex, colIndex + 1);
                }                
            }
        }
        private void CopyCurrentCell(FastGridControl copyTargetGrid)
        {
            if (copyTargetGrid == null)
                return;            
            var model = copyTargetGrid.Model as DiffGridModel;
            if (model == null)
                return;
            foreach (var currentCell in copyTargetGrid.SelectedCells)
            {
                string colHeaderText = model.GetColumnHeaderText(currentCell.Column.Value);
                string rowHeaderText = model.GetRowHeaderText(currentCell.Row.Value);
                int rowCount = model.RowCount;
                int columnCount = model.ColumnCount;
                var rowIndexInTarget = targetExcel.GetRowByHeaderText(rowHeaderText, currentCell.Row.Value);
                var colIndexInTarget = targetExcel.GetColByHeaderText(colHeaderText, currentCell.Column.Value);
                if (rowIndexInTarget.Item2 < 0 || colIndexInTarget.Item2 < 0)
                {
                    MessageBox.Show("没有在Excel表中找到" + colHeaderText + ":" + rowHeaderText);
                    return;
                }
                targetExcel.WriteData(model.GetCellText(copyTargetGrid.SelectedCells.First(), true), rowIndexInTarget.Item2, colIndexInTarget.Item2);
            }
        }
        private void CopyToClipboardSelectedCells(FastGridControl copyTargetGrid,string separator)
        {
            if (copyTargetGrid == null)
                return;

            var model = copyTargetGrid.Model as DiffGridModel;
            if (model == null)
                return;

            var tsv = string.Join(Environment.NewLine,
               copyTargetGrid.SelectedCells
              .GroupBy(c => c.Row.Value)
              .OrderBy(g => g.Key)
              .Select(g => string.Join(separator, g.Select(c => model.GetCellText(c, true)))));

            Clipboard.SetDataObject(tsv);
        }

        private void UserControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Right:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextModifiedCell();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Left:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevModifiedCell();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Down:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextModifiedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Up:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevModifiedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.L:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextRemovedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.O:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevRemovedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.K:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextAddedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.I:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevAddedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.F8:
                    {
                        MovePrevMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.F9:
                    {
                        MoveNextMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.F:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ToolExpander.IsExpanded = true;
                            SearchTextCombobox.Focus();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Z:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                        {                            
                            CopyCurrentRow(SrcDataGrid, false);
                        }
                        else if(Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
                        {
                            CopyCurrentSrcCell_Click(null, null);
                        }
                        else if (Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt))
                        {
                            CopyCurrentRow(SrcDataGrid,true);
                        }
                        e.Handled = true;
                    }
                    break;
                case Key.Y:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                        {
                            CopyCurrentRow(DstDataGrid, false);
                        }
                        else if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
                        {
                            CopyCurrentDstCell_Click(null, null);
                        }
                        else if(Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt))
                        {
                            CopyCurrentRow(DstDataGrid, true);
                        }
                        e.Handled = true;
                    }
                    break;
                case Key.B:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ShowLog();
                            e.Handled = true;
                        }
                    }
                    break;
            }
        }

        private void ShowLog()
        {
            var log = BuildCellBaseLog();

            (App.Current.MainWindow as MainWindow).WriteToConsole(log);
        }

        private void BuildCellBaseLog_Click(object sender, RoutedEventArgs e)
        {
            ShowLog();
        }

        private string BuildCellBaseLog()
        {
            var srcModel = SrcDataGrid.Model as DiffGridModel;
            if (srcModel == null)
                return string.Empty;

            var dstModel = DstDataGrid.Model as DiffGridModel;
            if (dstModel == null)
                return string.Empty;

            var builder = new StringBuilder();

            var selectedCells = SrcDataGrid.SelectedCells;

            var modifiedLogFormat = App.Instance.Setting.LogFormat;
            var addedLogFormat = App.Instance.Setting.AddedRowLogFormat;
            var removedLogFormat = App.Instance.Setting.RemovedRowLogFormat;

            foreach (var row in SrcDataGrid.SelectedCells.GroupBy(c => c.Row))
            {
                var rowHeaderText = srcModel.GetRowHeaderText(row.Key.Value);
                if (string.IsNullOrEmpty(rowHeaderText))
                    rowHeaderText = dstModel.GetRowHeaderText(row.Key.Value);

                if (dstModel.IsAddedRow(row.Key.Value, true))
                {
                    var log = addedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText));

                    builder.AppendLine(log);

                    continue;
                }

                if (dstModel.IsRemovedRow(row.Key.Value, true))
                {
                    var log = removedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText));

                    builder.AppendLine(log);

                    continue;
                }

                foreach (var cell in row)
                {
                    if (cell.Row.Value == srcModel.ColumnHeaderIndex)
                        continue;

                    var srcText = srcModel.GetCellText(cell, true);
                    var dstText = dstModel.GetCellText(cell, true);
                    if (srcText == dstText)
                        continue;

                    var colHeaderText = srcModel.GetColumnHeaderText(cell.Column.Value);

                    if (string.IsNullOrEmpty(colHeaderText))
                        colHeaderText = dstModel.GetColumnHeaderText(cell.Column.Value);

                    if (string.IsNullOrEmpty(srcText))
                        srcText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(dstText))
                        dstText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(rowHeaderText))
                        rowHeaderText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(colHeaderText))
                        colHeaderText = Properties.Resources.Word_Blank;

                    var log = modifiedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText))
                        .Replace("${COL}", RemoveMultiLine(colHeaderText))
                        .Replace("${LEFT}", RemoveMultiLine(srcText))
                        .Replace("${RIGHT}", RemoveMultiLine(dstText));

                    builder.AppendLine(log);
                }
            }

            return builder.ToString();
        }

        private string RemoveMultiLine(string log)
        {
            return log.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        }

        private void CopyCurrentSrcRow_Click(object sender, RoutedEventArgs e)
        {
            CopyCurrentRow(SrcDataGrid);
            CopyToClipboardSelectedCells(SrcDataGrid, "\t");
        }

        private void CopyCurrentDstRow_Click(object sender, RoutedEventArgs e)
        {
            CopyCurrentRow(DstDataGrid);
            CopyToClipboardSelectedCells(DstDataGrid, "\t");
        }

        private void ReplaceCurrentDstRow_Click(object sender, RoutedEventArgs e)
        {
            CopyCurrentRow(DstDataGrid, true);
            CopyToClipboardSelectedCells(DstDataGrid, "\t");
        }
        private void ReplaceCurrentSrcRow_Click(object sender, RoutedEventArgs e)
        {
            CopyCurrentRow(SrcDataGrid,true);
            CopyToClipboardSelectedCells(SrcDataGrid, "\t");
        }
        private void CopyCurrentSrcCell_Click(object sender, RoutedEventArgs e)
        {
            CopyCurrentCell(SrcDataGrid);
            CopyToClipboardSelectedCells(SrcDataGrid, ",");
        }
        private void CopyCurrentDstCell_Click(object sender, RoutedEventArgs e)
        {
            CopyCurrentCell(DstDataGrid);
            CopyToClipboardSelectedCells(DstDataGrid,",");
        }

        private void SrcSheetCombobox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }
        private void matchItem(ComboBox sender, string selectedValue)
        {
            string candidateItem = "";
            foreach (string item in sender.ItemsSource)
            {
                if (item == selectedValue)
                {
                    sender.SelectedValue = item;
                    break;
                }
                for(int i=0;i<item.Length&&i<selectedValue.Length;++i)
                {
                    if( selectedValue[i] =='.' || item[i] == '.')
                    {
                        candidateItem = item;
                        break;
                    }
                    if(item[i]!= selectedValue[i])
                    {
                        break;
                    }
                }
            }
            if (candidateItem != "")
            {
                sender.SelectedValue = candidateItem;
            }
        }
        private void SrcFileCombobox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var path = SrcPathTextBox.Text;
            var comb = sender as ComboBox;
            if (comb.SelectedIndex < 0)
            {
                return;
            }
            string selectedValue = comb.Items[comb.SelectedIndex].ToString();
            var newPath = Path.Combine(path, selectedValue);
            var exists = File.Exists(newPath);
            if(exists)
            {
                matchItem(DstFileCombobox, selectedValue);
                matchItem(TargetFileCombobox, selectedValue);
                try
                {                    
                    SrcSheetCombobox.ItemsSource = ExcelWorkbook.GetSheetNames(newPath).ToList();                    
                    SrcSheetCombobox.SelectedIndex = 0;
                }
                catch
                {
                    SrcSheetCombobox.SelectedIndex = -1;
                    SrcSheetCombobox.ItemsSource = new List<string>();
                }
            }

        }

        private void DstFileCombobox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var path = DstPathTextBox.Text;
            var comb = sender as ComboBox;
            if (comb.SelectedIndex < 0)
            {
                return;
            }
            string selectedValue = comb.Items[comb.SelectedIndex].ToString();
            var newPath = Path.Combine(path, selectedValue);
            var exists = File.Exists(newPath);
            if (exists)
            {
                try
                {
                    DstSheetCombobox.ItemsSource = ExcelWorkbook.GetSheetNames(newPath).ToList();
                    DstSheetCombobox.SelectedIndex = 0;
                }
                catch
                {
                    DstSheetCombobox.SelectedIndex = -1;
                    DstSheetCombobox.ItemsSource = new List<string>();
                }
            }
        }

        private void TargetFileCombobox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var path = TargetPathTextBox.Text;
            var comb = sender as ComboBox;
            if (comb.SelectedIndex < 0)
            {
                return;
            }
            string selectedValue = comb.Items[comb.SelectedIndex].ToString();
            var newPath = Path.Combine(path, selectedValue);
            var exists = File.Exists(newPath);
            if (exists)
            {
                try
                {
                    TargetSheetCombobox.ItemsSource = ExcelWorkbook.GetSheetNames(newPath).ToList();
                    TargetSheetCombobox.SelectedIndex = 0;
                }
                catch
                {
                    TargetSheetCombobox.SelectedIndex = -1;
                    TargetSheetCombobox.ItemsSource = new List<string>();
                }
            }
        }
    }
}
