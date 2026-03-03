using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace SetupSheet
{
    public static class DocumentGenerator
    {
        public static void GenerateTechCard(
            string projectPath,
            string article,
            string material,
            string machines,
            string stockSize,
            List<string> ncPrograms,
            List<string> tools,
            string time,
            string screenshotPath,
            string comments,
            string positionHoles = "",
            string author = "",
            string setupText = ""
        )
        {
            Word.Application wordApp = null;
            Word.Document doc = null;
            string templatePath = "";
            string outputFolder = "";
            string outputPath = "";
            string localScreenshotPath = null;
            bool useTempSave = false;

            try
            {
                string iniPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcellSetupSheet",
                    "Template.ini"
                );

                if (File.Exists(iniPath))
                {
                    templatePath = File.ReadAllLines(iniPath, Encoding.UTF8)
                                       .LastOrDefault()?.Trim() ?? "";
                }

                if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
                {
                    MessageBox.Show($"Шаблон не найден:\n{templatePath}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!templatePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("Неверный формат шаблона. Ожидается .docx файл", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                outputFolder = Path.Combine(projectPath, "Техкарта");
                Directory.CreateDirectory(outputFolder);

                outputPath = Path.Combine(outputFolder, "Техкарта.docx");

                if (IsNetworkPath(projectPath))
                {
                    useTempSave = true;
                }

                if (!string.IsNullOrWhiteSpace(screenshotPath) && File.Exists(screenshotPath))
                {
                    string tempScreenshotsDir = Path.Combine(Path.GetTempPath(), "TechCardScreenshots");
                    Directory.CreateDirectory(tempScreenshotsDir);
                    localScreenshotPath = Path.Combine(tempScreenshotsDir, "screenshot.png");

                    try
                    {
                        File.Copy(screenshotPath, localScreenshotPath, true);
                    }
                    catch
                    {
                        localScreenshotPath = screenshotPath;
                    }
                }

                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                doc = wordApp.Documents.Add(templatePath);

                ReplaceAll(doc, "{article}", article ?? "");
                ReplaceAll(doc, "{material}", material ?? "");
                ReplaceAll(doc, "{machines}", machines ?? "");
                ReplaceAll(doc, "{stockSize}", stockSize ?? "");
                ReplaceAll(doc, "{projectPath}", projectPath ?? "");
                ReplaceAll(doc, "{time}", time ?? "");
                ReplaceAll(doc, "{comments}", comments ?? "");
                ReplaceAll(doc, "{ncList}", JoinLines(ncPrograms));
                ReplaceAll(doc, "{toolsList}", JoinLines(tools));
                ReplaceAll(doc, "{positionHoles}", positionHoles ?? "");
                ReplaceAll(doc, "{setupText}", setupText ?? "");

                string authorValue = string.IsNullOrWhiteSpace(author) ? Environment.UserName : author;
                ReplaceAll(doc, "{author}", authorValue);

                string screenshotToUse = localScreenshotPath ?? screenshotPath;

                if (!string.IsNullOrWhiteSpace(screenshotToUse) && File.Exists(screenshotToUse))
                {
                    try
                    {
                        ReplaceScreenshot(doc, "{screenshot}", screenshotToUse);
                    }
                    catch
                    {
                        ReplaceAll(doc, "{screenshot}", "");
                        ReplaceAll(doc, "{ screenshot }", "");
                    }
                }
                else
                {
                    ReplaceAll(doc, "{screenshot}", "");
                    ReplaceAll(doc, "{ screenshot }", "");
                }

                if (useTempSave)
                {
                    string tempPath = Path.Combine(Path.GetTempPath(), "TechCard_" + Guid.NewGuid().ToString("N") + ".docx");

                    TrySaveDocx(doc, tempPath);

                    doc.Close(true);
                    ReleaseCom(doc);
                    doc = null;

                    System.Threading.Thread.Sleep(500);

                    wordApp.Quit();
                    ReleaseCom(wordApp);
                    wordApp = null;

                    System.Threading.Thread.Sleep(500);

                    Directory.CreateDirectory(outputFolder);

                    int retries = 0;
                    while (retries < 3)
                    {
                        try
                        {
                            if (File.Exists(tempPath))
                            {
                                File.Copy(tempPath, outputPath, true);
                                File.Delete(tempPath);
                                break;
                            }
                        }
                        catch
                        {
                            retries++;
                            System.Threading.Thread.Sleep(500);
                        }
                    }
                }
                else
                {
                    TrySaveDocx(doc, outputPath);

                    doc.Close(true);
                    ReleaseCom(doc);
                    doc = null;

                    wordApp.Quit();
                    ReleaseCom(wordApp);
                    wordApp = null;
                }
            }
            catch (COMException comEx)
            {
                try
                {
                    string tempPath = Path.Combine(Path.GetTempPath(), "TechCard_" + Guid.NewGuid().ToString("N") + ".docx");

                    if (doc != null)
                    {
                        TrySaveDocx(doc, tempPath);

                        doc.Close(true);
                        ReleaseCom(doc);
                        doc = null;

                        if (wordApp != null)
                        {
                            System.Threading.Thread.Sleep(500);
                            wordApp.Quit();
                            ReleaseCom(wordApp);
                            wordApp = null;
                        }

                        Directory.CreateDirectory(outputFolder);
                        File.Copy(tempPath, outputPath, true);
                        File.Delete(tempPath);
                    }
                }
                catch (Exception fallbackEx)
                {
                    MessageBox.Show($"Ошибка COM:\n{comEx.Message}\nFallback:\n{fallbackEx.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка:\n{ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (doc != null) { try { doc.Close(false); } catch { } ReleaseCom(doc); }
                if (wordApp != null) { try { wordApp.Quit(); } catch { } ReleaseCom(wordApp); }

                if (!string.IsNullOrEmpty(localScreenshotPath) && File.Exists(localScreenshotPath))
                {
                    try
                    {
                        System.Threading.Thread.Sleep(500);
                        File.Delete(localScreenshotPath);
                        string dir = Path.GetDirectoryName(localScreenshotPath);
                        if (Directory.Exists(dir) && Directory.GetFiles(dir).Length == 0)
                            Directory.Delete(dir);
                    }
                    catch { }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static bool IsNetworkPath(string path)
        {
            try
            {
                if (path.StartsWith(@"\\")) return true;

                string driveLetter = Path.GetPathRoot(path);
                if (!string.IsNullOrEmpty(driveLetter) && driveLetter.Length >= 2)
                {
                    DriveInfo drive = new DriveInfo(driveLetter);
                    return drive.DriveType == DriveType.Network;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static void TrySaveDocx(Word.Document doc, string fullPath)
        {
            object fileName = fullPath;
            object fileFormat = Word.WdSaveFormat.wdFormatXMLDocument;
            doc.SaveAs2(ref fileName, ref fileFormat);
            doc.Saved = true;
        }

        private static void ReplaceAll(Word.Document doc, string placeholder, string value)
        {
            if (doc == null) return;

            string[] placeholders = new string[]
            {
                placeholder,
                "{ " + placeholder.Trim(new char[] { '{', '}', ' ' }) + " }"
            };

            Word.Range range = null;
            try
            {
                foreach (Word.Range storyRange in doc.StoryRanges)
                {
                    range = storyRange;
                    while (range != null)
                    {
                        foreach (string ph in placeholders)
                        {
                            if (!string.IsNullOrWhiteSpace(ph))
                                FindReplaceInRange(range, ph, value ?? "");
                        }
                        range = range.NextStoryRange;
                    }
                }
            }
            catch { }
            finally { if (range != null) ReleaseCom(range); }
        }

        private static void FindReplaceInRange(Word.Range range, string findText, string replaceText)
        {
            if (range == null) return;
            Word.Find find = null;
            try
            {
                find = range.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = findText;
                find.Replacement.Text = replaceText;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindContinue;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;
                object replaceAll = Word.WdReplace.wdReplaceAll;
                find.Execute(Replace: ref replaceAll);
            }
            finally { if (find != null) ReleaseCom(find); }
        }

        private static void ReplaceScreenshot(Word.Document doc, string placeholder, string screenshotPath)
        {
            if (doc == null) return;

            System.Diagnostics.Debug.WriteLine($"Попытка вставить скриншот из: {screenshotPath}");
            System.Diagnostics.Debug.WriteLine($"Файл существует: {File.Exists(screenshotPath)}");

            string[] placeholders = new string[]
            {
                placeholder,
                "{ " + placeholder.Trim(new char[] { '{', '}', ' ' }) + " }"
            };

            bool inserted = false;
            foreach (Word.Range storyRange in doc.StoryRanges)
            {
                Word.Range r = storyRange;
                while (r != null)
                {
                    foreach (string ph in placeholders)
                    {
                        if (!string.IsNullOrWhiteSpace(ph))
                        {
                            System.Diagnostics.Debug.WriteLine($"Ищем плейсхолдер: '{ph}'");
                            if (TryInsertPictureInRange(r, ph, screenshotPath))
                            {
                                System.Diagnostics.Debug.WriteLine("Скриншот вставлен успешно!");
                                inserted = true;
                                return;
                            }
                        }
                    }
                    r = r.NextStoryRange;
                }
            }

            foreach (string ph in placeholders)
            {
                if (!string.IsNullOrWhiteSpace(ph) && TryInsertPictureInRange(doc.Content, ph, screenshotPath))
                {
                    inserted = true;
                    return;
                }
            }

            if (!inserted)
            {
                System.Diagnostics.Debug.WriteLine("Не удалось вставить скриншот!");
            }
        }

        private static bool TryInsertPictureInRange(Word.Range range, string placeholder, string screenshotPath)
        {
            Word.Find find = null;
            try
            {
                find = range.Find;
                find.ClearFormatting();
                find.Text = placeholder;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;

                bool found = find.Execute();
                if (!found) return false;

                Word.Cell cell = null;
                try
                {
                    cell = range.Cells[1];
                }
                catch { }

                if (cell != null)
                {
                    cell.Range.Text = "";
                    cell.Range.InlineShapes.AddPicture(
                        FileName: screenshotPath,
                        LinkToFile: false,
                        SaveWithDocument: true
                    );
                }
                else
                {
                    range.Text = "";
                    range.InlineShapes.AddPicture(
                        FileName: screenshotPath,
                        LinkToFile: false,
                        SaveWithDocument: true
                    );
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка вставки скриншота: {ex.Message}");
                return false;
            }
            finally
            {
                if (find != null) ReleaseCom(find);
            }
        }

        private static string JoinLines(List<string> list)
        {
            if (list == null || list.Count == 0) return "";
            return string.Join("\r\n", list.Where(s => !string.IsNullOrWhiteSpace(s)));
        }

        private static bool IsFileLocked(string path)
        {
            try
            {
                using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    return false;
            }
            catch { return true; }
        }

        private static void ReleaseCom(object o)
        {
            try { if (o != null && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o); }
            catch { }
        }
    }
}