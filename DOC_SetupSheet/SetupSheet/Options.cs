using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SetupSheet
{
    public partial class Options : Form
    {
        public string FolderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ExcellSetupSheet";

        public Options()
        {
            InitializeComponent();

            string Path = FolderPath + "\\Template.ini";

            if (File.Exists(Path))
            {
                using (var reader = new StreamReader(Path, Encoding.UTF8))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        Template.Text = line;
                    }
                }
            }
            else
            {
                Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string directory = System.IO.Path.GetDirectoryName(Path);
                Path = directory + "\\TechCard_Template.docx";
                Template.Text = Path;
            }
        }

        private void OK_Click(object sender, EventArgs e)
        {
            string Path = FolderPath + "\\Template.ini";

            if (string.IsNullOrWhiteSpace(Template.Text))
            {
                MessageBox.Show("Please select a template");
            }
            else
            {
                if (File.Exists(Path)) File.Delete(Path);
                if (!Directory.Exists(FolderPath)) Directory.CreateDirectory(FolderPath);
                File.WriteAllText(Path, Template.Text, Encoding.UTF8);
                this.Close();
            }
        }

        private void Browse_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Select Word template";
            dlg.Filter = "Word files (*.docx)|*.docx|All Files (*.*)|*.*";
            dlg.FilterIndex = 1;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Template.Text = dlg.FileName;
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}