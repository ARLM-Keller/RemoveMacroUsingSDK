using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.DocumentFormat.OpenXml.Packaging;
using System.Xml;

namespace RemoveMacroUsingSDK
{
    public partial class RemoveMacro : Form
    {
        public RemoveMacro()
        {
            InitializeComponent();
        }

        private string GetOpenPath()
        {
            OpenFileDialog sfd = new OpenFileDialog();
            sfd.AddExtension = true;
            //Get only Docx file
            sfd.Filter = "Word Document (*.docm)|*.docm|Presentation Document (*.pptx)|*.pptx|Excel Document (*.xlsx)|*.xlsx";
            sfd.CheckPathExists = true;
            sfd.DefaultExt = ".docm";
            sfd.ShowDialog();
            return sfd.FileName;
            // return the filename and the path in which the user wants to create the file
        }


        private void RemoveMacro_Load(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //string filename = @"c:\macro.docm";
            string filename = GetOpenPath();

            WordprocessingDocument wordDoc = WordprocessingDocument.Open(filename, true);
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;

            const string relTypeVBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

            foreach (IdPartPair var in mainPart.Parts)
            {
                if (var.OpenXmlPart.RelationshipType.ToString() == relTypeVBA)
                {
                    OpenXmlPart macroPart = (OpenXmlPart)var.OpenXmlPart;
                    mainPart.DeletePart(macroPart);
                    break;
                }
            }

            //delete the main part and recreate it.
            XmlDocument mainpartdoc = new XmlDocument();
            mainpartdoc.Load(mainPart.GetStream());

            wordDoc.DeletePart(mainPart);
            MainDocumentPart main = wordDoc.AddMainDocumentPart();
            StreamWriter mainstream = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write));
            mainpartdoc.Save(mainstream);

            wordDoc.ChangeDocumentType(WordprocessingDocumentType.Document);
            wordDoc.Close();


            string newfilename = @"c:\MacroFree.docx";

            File.Move(filename, newfilename);


            MessageBox.Show("done");
        }
    }
}
    
