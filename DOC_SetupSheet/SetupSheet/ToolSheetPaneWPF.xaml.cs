using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using PowerMILL;
using Delcam.Plugins.Framework;
using Delcam.Plugins.Events;
using IOPath = System.IO.Path;
using System.Globalization;
using System.Linq;

namespace SetupSheet
{
    public partial class ToolSheetPaneWPF : UserControl
    {
        private IPluginCommunicationsInterface oComm;
        private string selectedMachine = "";
        private string flipAxis = "";

        public ToolSheetPaneWPF(IPluginCommunicationsInterface comms)
        {
            InitializeComponent();
            oComm = comms;
            PowerMILLAutomation.SetVariables(oComm);

            comms.EventUtils.Subscribe(new EventSubscription("EntityCreated", EntityCreated));
            comms.EventUtils.Subscribe(new EventSubscription("EntityDeleted", EntityDeleted));
            comms.EventUtils.Subscribe(new EventSubscription("ProjectClosed", ProjectClosed));
            comms.EventUtils.Subscribe(new EventSubscription("ProjectOpened", ProjectOpened));
            comms.EventUtils.Subscribe(new EventSubscription("EntityRenamed", EntityRenamed));

            LoadProjectComponents();
            if (MachineComboBox.Items.Count > 0)
                MachineComboBox.SelectedIndex = 0;
        }

        public void PreInitialise(string locale) { }
        public void ProcessCommand(string command) { }
        public void ProcessEvent(string eventData) { }
        public void SerializeProjectData(string path, bool saving) { }
        public void Uninitialise() { GC.Collect(); }

        private void LoadProjectComponents()
        {
            listNCProgs.Items.Clear();
            List<string> ncprogs = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.NCPrograms);
            foreach (string ncprog in ncprogs)
                listNCProgs.Items.Add(ncprog);
        }

        void ProjectOpened(string eventName, Dictionary<string, string> eventArguments) { LoadProjectComponents(); }

        void EntityCreated(string eventName, Dictionary<string, string> eventArguments)
        {
            if (eventArguments["EntityType"] == "Ncprogram")
                listNCProgs.Items.Add(eventArguments["Name"]);
        }

        void EntityRenamed(string eventName, Dictionary<string, string> eventArguments)
        {
            if (eventArguments["EntityType"] == "Ncprogram")
            {
                string orig = eventArguments["PreviousName"];
                string newName = eventArguments["Name"];
                if (listNCProgs.Items.Contains(orig)) { listNCProgs.Items.Remove(orig); listNCProgs.Items.Add(newName); }
                if (listNCProgsSelected.Items.Contains(orig)) { listNCProgsSelected.Items.Remove(orig); listNCProgsSelected.Items.Add(newName); }
            }
        }

        void EntityDeleted(string eventName, Dictionary<string, string> eventArguments)
        {
            if (eventArguments["EntityType"] == "Ncprogram")
            {
                string name = eventArguments["Name"];
                if (listNCProgs.Items.Contains(name)) listNCProgs.Items.Remove(name);
                else if (listNCProgsSelected.Items.Contains(name)) listNCProgsSelected.Items.Remove(name);
            }
        }

        void ProjectClosed(string eventName, Dictionary<string, string> eventArguments) { listNCProgs.Items.Clear(); }

        private void AddNCProg_Click(object sender, RoutedEventArgs e)
        {
            if (listNCProgs.SelectedItems.Count > 0)
            {
                for (int i = listNCProgs.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object item = listNCProgs.SelectedItems[i];
                    listNCProgsSelected.Items.Insert(0, item);
                    listNCProgs.Items.Remove(item);
                }
            }
        }

        private void RemoveNCProg_Click(object sender, RoutedEventArgs e)
        {
            if (listNCProgsSelected.SelectedItems.Count > 0)
            {
                for (int i = listNCProgsSelected.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object item = listNCProgsSelected.SelectedItems[i];
                    listNCProgs.Items.Insert(0, item);
                    listNCProgsSelected.Items.Remove(item);
                }
            }
        }

        private void MachineComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MachineComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                selectedMachine = selectedItem.Content.ToString();
                if (selectedMachine.ToString().Contains("переворот X") || selectedMachine.ToString().Contains("(X)"))
                    flipAxis = "X";
                else if (selectedMachine.ToString().Contains("переворот Y") || selectedMachine.ToString().Contains("(Y)"))
                    flipAxis = "Y";
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string projectPath = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)").Trim();

                if (string.IsNullOrWhiteSpace(projectPath))
                {
                    MessageBox.Show("Не удалось получить путь к проекту.");
                    return;
                }

                List<string> ncList = new List<string>();
                foreach (var item in listNCProgsSelected.Items)
                    ncList.Add(item.ToString());

                if (ncList.Count == 0)
                {
                    MessageBox.Show("Выберите хотя бы одну NC программу.");
                    return;
                }

                double xmin = double.Parse(PowerMILLAutomation.ExecuteEx("print $block.limits.xmin").Replace(",", "."), CultureInfo.InvariantCulture);
                double xmax = double.Parse(PowerMILLAutomation.ExecuteEx("print $block.limits.xmax").Replace(",", "."), CultureInfo.InvariantCulture);
                double ymin = double.Parse(PowerMILLAutomation.ExecuteEx("print $block.limits.ymin").Replace(",", "."), CultureInfo.InvariantCulture);
                double ymax = double.Parse(PowerMILLAutomation.ExecuteEx("print $block.limits.ymax").Replace(",", "."), CultureInfo.InvariantCulture);
                double zmin = double.Parse(PowerMILLAutomation.ExecuteEx("print $block.limits.zmin").Replace(",", "."), CultureInfo.InvariantCulture);
                double zmax = double.Parse(PowerMILLAutomation.ExecuteEx("print $block.limits.zmax").Replace(",", "."), CultureInfo.InvariantCulture);

                double blockX = xmax - xmin;
                double blockY = ymax - ymin;
                double blockZ = zmax - zmin;

                HashSet<string> tools = new HashSet<string>();
                foreach (string ncprog in ncList)
                {
                    List<string> toolpaths = PowerMILLAutomation.GetNCProgToolpathes(ncprog);
                    foreach (string tp in toolpaths)
                    {
                        try
                        {
                            string tpEsc = tp.Replace("'", "''");
                            string toolName = PowerMILLAutomation.ExecuteEx($"print $entity('toolpath';'{tpEsc}').tool.name").Trim();
                            string toolNumRaw = PowerMILLAutomation.ExecuteEx($"print $entity('toolpath';'{tpEsc}').tool.number").Trim();

                            int toolNum;
                            if (!int.TryParse(toolNumRaw.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out toolNum))
                                toolNum = -1;

                            if (!string.IsNullOrWhiteSpace(toolName))
                            {
                                if (toolNum >= 0)
                                    tools.Add($"{toolName} - T{toolNum}");
                                else
                                    tools.Add(toolName);
                            }
                        }
                        catch { }
                    }
                }

                double totalTimeMinutes = 0.0;
                foreach (string ncprog in ncList)
                {
                    try
                    {
                        string timeMin = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';'" + ncprog + "').Statistics.TotalTime");
                        timeMin = timeMin.Trim().Replace(",", ".");

                        if (double.TryParse(timeMin, NumberStyles.Any, CultureInfo.InvariantCulture, out double minutes))
                        {
                            totalTimeMinutes += minutes;
                        }
                    }
                    catch
                    {
                        try
                        {
                            List<string> toolpaths = PowerMILLAutomation.GetNCProgToolpathes(ncprog);
                            foreach (string tp in toolpaths)
                            {
                                try
                                {
                                    string tpEsc = tp.Replace("'", "''");
                                    string tpTime = PowerMILLAutomation.ExecuteEx($"print $entity('toolpath';'{tpEsc}').statistics.totaltime");
                                    tpTime = tpTime.Trim().Replace(",", ".");

                                    if (double.TryParse(tpTime, NumberStyles.Any, CultureInfo.InvariantCulture, out double seconds))
                                    {
                                        totalTimeMinutes += seconds / 60.0;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
                int totalHours = (int)(totalTimeMinutes / 60);
                int totalMins = (int)(totalTimeMinutes % 60);
                string time = $"{totalHours} ч {totalMins} мин";

                string article = ArticleTextBox.Text;
                string material = MaterialTextBox.Text;
                string positionHoles = PositionHolesTextBox.Text;
                string comments = Comments_TextBox.Text;

                string machineName = selectedMachine;
                int bracketIndex = machineName.IndexOf(" (");
                if (bracketIndex > 0)
                    machineName = machineName.Substring(0, bracketIndex);

                StringBuilder setupText = new StringBuilder();
                if (OneSideRadio.IsChecked == true)
                {
                    setupText.AppendLine("1 сторона X0Y0 - Центр, Z0 - От стола");
                }
                else
                {
                    setupText.AppendLine("1 сторона X0Y0 - Центр, Z0 - От верха");
                    setupText.AppendLine($"Перевернуть по оси {flipAxis}");
                    setupText.AppendLine("2 сторона X0Y0 - Центр, Z0 - От стола");
                }
                string setupTextValue = setupText.ToString().Trim();

                string folder = IOPath.Combine(projectPath, "Техкарта");
                Directory.CreateDirectory(folder);

                string materialText = material + " " + blockZ + " мм";
                string stockSize = $"{blockX:F0}x{blockY:F0}x{blockZ:F0}мм";
                string screenshotPath = IOPath.Combine(folder, "screenshot.png");

                DocumentGenerator.GenerateTechCard(
                    projectPath,
                    article,
                    materialText,
                    machineName,
                    stockSize,
                    ncList,
                    tools.ToList(),
                    time,
                    screenshotPath,
                    comments,
                    positionHoles,
                    Environment.UserName,
                    setupTextValue
                );

                MessageBox.Show("Файлы успешно созданы!", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public static FileInfo GetNewestFile(DirectoryInfo directory)
        {
            return directory.GetFiles()
                .Union(directory.GetDirectories().Select(d => GetNewestFile(d)))
                .OrderByDescending(f => (f == null ? DateTime.MinValue : f.LastWriteTime))
                .FirstOrDefault();
        }

        private void Get_Picture_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (listNCProgsSelected.Items.Count == 0)
                {
                    MessageBox.Show("Please select at least 1 NC Program");
                    return;
                }

                string NCProg = listNCProgsSelected.Items[0].ToString();
                string projectPath = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)");
                string snapshotPath = projectPath + "\\SetupSheets_files\\snapshots";

                PowerMILLAutomation.ExecuteEx("KEEP SNAPSHOT NCPROGRAM '" + NCProg + "' CURRENT");
                System.Threading.Thread.Sleep(500);

                FileInfo picture = GetNewestFile(new DirectoryInfo(snapshotPath));
                if (picture == null)
                {
                    MessageBox.Show("Screenshot not found");
                    return;
                }

                string targetFolder = IOPath.Combine(projectPath, "Техкарта");
                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                string targetPath = IOPath.Combine(targetFolder, "screenshot.png");
                if (File.Exists(targetPath))
                    File.Delete(targetPath);
                File.Copy(picture.FullName, targetPath);

                BitmapImage image = new BitmapImage();
                using (var stream = File.OpenRead(targetPath))
                {
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.StreamSource = stream;
                    image.EndInit();
                }
                Setup_Picture.Source = image;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void Del_Picture_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string projectPath = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)");
                string screenshotPath = IOPath.Combine(projectPath, "Техкарта", "screenshot.png");
                if (File.Exists(screenshotPath))
                {
                    File.Delete(screenshotPath);
                    Setup_Picture.Source = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        // ===== Drag & Drop - перетаскивание между списками =====
        private void listNCProgs_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed && listNCProgs.SelectedItems.Count > 0)
            {
                DragDrop.DoDragDrop(listNCProgs, listNCProgs.SelectedItems.Cast<object>().ToList(), DragDropEffects.Move);
            }
        }

        private void listNCProgsSelected_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed && listNCProgsSelected.SelectedItems.Count > 0)
            {
                DragDrop.DoDragDrop(listNCProgsSelected, listNCProgsSelected.SelectedItems.Cast<object>().ToList(), DragDropEffects.Move);
            }
        }

        private void listNCProgs_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Move;
            e.Handled = true;
        }

        private void listNCProgsSelected_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Move;
            e.Handled = true;
        }

        private void listNCProgs_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(List<object>)))
            {
                var items = (List<object>)e.Data.GetData(typeof(List<object>));
                foreach (var item in items)
                {
                    if (listNCProgsSelected.Items.Contains(item))
                    {
                        listNCProgsSelected.Items.Remove(item);
                        listNCProgs.Items.Add(item);
                    }
                }
            }
        }

        private void listNCProgsSelected_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(List<object>)))
            {
                var items = (List<object>)e.Data.GetData(typeof(List<object>));
                foreach (var item in items)
                {
                    if (listNCProgs.Items.Contains(item))
                    {
                        listNCProgs.Items.Remove(item);
                        listNCProgsSelected.Items.Add(item);
                    }
                }
            }
        }

        // ===== Кнопки изменения порядка ⬆ =====
        private void MoveUpBtn_Click(object sender, RoutedEventArgs e)
        {
            if (listNCProgsSelected.SelectedItem != null && listNCProgsSelected.SelectedIndex > 0)
            {
                int index = listNCProgsSelected.SelectedIndex;
                object item = listNCProgsSelected.SelectedItem;

                listNCProgsSelected.Items.RemoveAt(index);
                listNCProgsSelected.Items.Insert(index - 1, item);
                listNCProgsSelected.SelectedIndex = index - 1;
            }
        }

        private void MoveDownBtn_Click(object sender, RoutedEventArgs e)
        {
            if (listNCProgsSelected.SelectedItem != null &&
                listNCProgsSelected.SelectedIndex < listNCProgsSelected.Items.Count - 1)
            {
                int index = listNCProgsSelected.SelectedIndex;
                object item = listNCProgsSelected.SelectedItem;

                listNCProgsSelected.Items.RemoveAt(index);
                listNCProgsSelected.Items.Insert(index + 1, item);
                listNCProgsSelected.SelectedIndex = index + 1;
            }
        }

        private void ConvertToUsableName(string originalName, out string newName)
        {
            newName = originalName;
            foreach (char c in IOPath.GetInvalidFileNameChars())
                newName = newName.Replace(c, '_');
            newName = newName.Replace("*", "  ");
        }
    }
}