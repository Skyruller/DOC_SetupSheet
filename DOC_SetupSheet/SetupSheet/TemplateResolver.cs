using System.IO;
using System.Reflection;
using System.Windows;

namespace SetupSheet
{
    internal static class TemplateResolver
    {
        public static string Resolve(string name)
        {
            string dllDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            string[] probe =
            {
                Path.Combine(dllDir, name),
                Path.Combine(@"C:\Program Files (x86)\Autodesk\Excel SetupSheet", name),
                Path.Combine(@"C:\Program Files\Autodesk\Excel SetupSheet", name),
            };

            foreach (var p in probe)
                if (File.Exists(p)) return p;

            var msg = $"Шаблон не найден: {name}.\nУказать путь вручную?";
            if (MessageBox.Show(msg, "Excel SetupSheet",
                    MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                    FileName = name
                };
                if (dlg.ShowDialog() == true) return dlg.FileName;
            }

            throw new FileNotFoundException($"Template not found: {name}");
        }
    }
}
