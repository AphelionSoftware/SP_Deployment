//original namespaces
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//added namespaces
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        const string Nspace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        
        private void button1_Click(object sender, EventArgs e)
        {
            string sourcefile = "";
            string myfile = "";
            string mydirectory = "";
           
            // Ensure that the textboxes are not blank.
            if ((string.IsNullOrEmpty(this.Path.Text)))
            {
                MessageBox.Show("Please specify the full directory path.");
                return;
            }

            if ((string.IsNullOrEmpty(this.ConnectionsPath.Text)))
            {
                MessageBox.Show("Please specify the name of the excel file.");
                return;
            }

            if ((string.IsNullOrEmpty(this.Location.Text)))
            {
                MessageBox.Show("Please specify the new location of the odc file.");
                return;
            }

            if ((!string.IsNullOrEmpty(this.Path.Text) | !string.IsNullOrEmpty(this.ConnectionsPath.Text) | !string.IsNullOrEmpty(this.Location.Text)))
            {
                myfile = this.ConnectionsPath.Text;
                mydirectory = this.Path.Text;
                sourcefile = System.IO.Path.Combine(mydirectory, myfile);
                UpdateConnectPath(sourcefile, myfile, mydirectory);
            }
        }

        private void UpdateConnectPath(string sourcefile, string myfile, string mydirectory)
        {
            //Variables
            string newpath = System.IO.Path.Combine(mydirectory, "Temp");
            //string sourcefile = System.IO.Path.Combine(mydirectory, myfile);
            string destfile = System.IO.Path.Combine(newpath, myfile);
            string xmlpath = @"xl\connections.xml";

            //Create new subdirectoy and copy file there
            System.IO.Directory.CreateDirectory(newpath);
            System.IO.File.Copy(sourcefile, destfile, true);

            //Convert .xlsx to .zip
            string result = System.IO.Path.ChangeExtension(destfile, ".zip");
            File.Move(destfile, System.IO.Path.ChangeExtension(destfile, ".zip"));

            //Unzip contents
            ZipFile.ExtractToDirectory(result, newpath);

            //Change xml value
            string path = System.IO.Path.Combine(newpath, xmlpath);
            var xDoc = XDocument.Load(path);
            var ns = XNamespace.Get(Nspace);

            var odcFile = xDoc.Root.Elements(ns + "connection")
                                   .FirstOrDefault(x => (int)x.Attribute("id") == 1)
                                   .Attribute("odcFile");

            odcFile.Value = this.Location.Text;

            xDoc.Save(path);
            
            //open the document
            using (SpreadsheetDocument wkb = SpreadsheetDocument.Open(sourcefile, true))
            {
                //remove the current xlsx Connection part
                //wkb.WorkbookPart.DeletePart(wkb.WorkbookPart.ConnectionsPart);
                ConnectionsPart Part = wkb.WorkbookPart.ConnectionsPart;

                //using the new XML snippet
                using (StreamReader read = new StreamReader(path))
                {
                    string contents = read.ReadToEnd();

                    //commit the new part to the spreadsheet
                    using (StreamWriter write = new StreamWriter(Part.GetStream(FileMode.Create)))
                    {
                        write.Write(contents);
                    }
                }
                
                // Close workbook and exit.
                wkb.Close();
                System.IO.Directory.Delete(newpath, true);
                MessageBox.Show("Operation Complete");
                System.Windows.Forms.Application.Exit();

            }

	}
  }
}

