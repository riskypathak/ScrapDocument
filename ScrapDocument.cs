using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ScrapDocument
{
    public partial class ScrapDocument : Form
    {
        IList<string> FileNames = new List<string>();
        string xml = string.Empty;

        public ScrapDocument()
        {
            InitializeComponent();
            cmbTemplate.DataSource = (new TemplateXml()).GetAllTemplateTypes();
            cmbTemplate.DisplayMember = "value";
        }

        #region Events

        private void Select_Click(object sender, EventArgs e)
        {
            FileNames.Clear();
            DialogResult result = this.fbdDirectory.ShowDialog();

            if (result == DialogResult.OK)
            {
                string foldername = txtDirectory.Text = this.fbdDirectory.SelectedPath;
                foreach (string file in Directory.GetFiles(foldername))
                {
                    if (file.Trim().EndsWith("doc") || file.Trim().EndsWith("docx"))
                    {
                        FileNames.Add(file);
                        txtLog.Text += string.Format("\r\nFile added: {0}\r\n*************", file);
                    }
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            string logDetails = string.Empty;

            if (FileNames.Count > 0)
            {
                IList<Article> ArticlesData = new List<Article>();

                foreach (string fileName in FileNames)
                {
                    try
                    {
                        Article articleData = BusinessLogic.ExtractArticleData(fileName, Convert.ToString(cmbTemplate.SelectedItem));
                        ArticlesData.Add(articleData);
                        txtLog.Text += string.Format("\r\nArticle Data extracted for  {0}\r\n********************", fileName);
                    }
                    catch (Exception exception)
                    {
                        txtLog.Text += string.Format("\r\nFile {0} could not be processed\r\n********************", fileName);
                        logDetails += "\r\n***********Error Detail Summary***********\r\n" + "Error logged for file " + fileName + "\r\n" + exception.ToString() + "\r\n";
                    }
                }

                xml = BusinessLogic.GenerateXML(ArticlesData);
                if (!string.IsNullOrWhiteSpace(xml))
                {
                    DialogResult result = fbdSaveFolder.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        var destinationFolderName = fbdSaveFolder.SelectedPath;
                        if (Directory.Exists(destinationFolderName))
                        {
                            File.WriteAllText(destinationFolderName + "/" + "xml.xml", xml);
                            txtLog.Text += string.Format("\r\nXml file saved\r\n");

                            BusinessLogic.WritetoExcel(ArticlesData, destinationFolderName);
                            txtLog.Text += string.Format("\r\nExcel sheet created\r\n");

                            logDetails += "\r\n**********************Log Detail Summary**********************\r\n" + txtLog.Text;
                            System.IO.StreamWriter file = new System.IO.StreamWriter(destinationFolderName + "\\LogMonitor.txt");
                            file.WriteLine(logDetails);
                            file.Close();

                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Kindly give a valid directory path");
            }


        }

        private void txtLog_TextChanged(object sender, EventArgs e)
        {
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        private void txtDirectory_TextChanged(object sender, EventArgs e)
        {
            txtLog.Clear();
            FileNames.Clear();
        }

        #endregion

    }
}
