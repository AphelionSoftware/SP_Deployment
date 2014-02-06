using Microsoft.VisualBasic;
using System;
using System.IO;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;
using System.Diagnostics;

namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        const string spreadsheetmlNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        
        private void button1_Click(object sender, EventArgs e)
        {
            string strSourceFile = "";
           
            // Ensure that the textboxes are not blank.
            if ((string.IsNullOrEmpty(this.Path.Text)))
            {
                MessageBox.Show("Please select the file path.");
                return;
            }

            if ((string.IsNullOrEmpty(this.Location.Text)))
            {
                MessageBox.Show("Please specify the location of the external data.");
                return;
            }

            if ((!string.IsNullOrEmpty(this.Path.Text) | !string.IsNullOrEmpty(this.Location.Text)))
            {
                strSourceFile = this.Path.Text;
                UpdateConnectPath(strSourceFile);
            }
        }

        private void UpdateConnectPath(string strSourceFile)
        {
            
            string path = "connections.xml";
            var xDoc = XDocument.Load(path);
            var ns = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var odcFile = xDoc.Root.Elements(ns + "connection")
                                   .FirstOrDefault(x => (int)x.Attribute("id") == 1)
                                   .Attribute("odcFile");

            odcFile.Value = this.Location.Text;

            xDoc.Save(path);
            
            //open the document
            using (SpreadsheetDocument wkb = SpreadsheetDocument.Open(strSourceFile, true))
            {
                //remove the current xlsx Connection part
                wkb.WorkbookPart.DeletePart(wkb.WorkbookPart.ConnectionsPart);
                ConnectionsPart newPart = wkb.WorkbookPart.AddNewPart<ConnectionsPart>();

                //using the new XML snippet
                using (StreamReader read = new StreamReader(path))
                {
                    string contents = read.ReadToEnd();

                    //commit the new part to the spreadsheet
                    using (StreamWriter write = new StreamWriter(newPart.GetStream(FileMode.Create)))
                    {
                        write.Write(contents);
                    }
                }
                
                // Close workbook and exit.
                wkb.Close();
                MessageBox.Show("Operation Complete");
                System.Windows.Forms.Application.Exit();
            }

	}
  }
}

