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

            if ((string.IsNullOrEmpty(this.oldLocation.Text) | string.IsNullOrEmpty(this.Location.Text)))
            {
                MessageBox.Show("Please specify the location of the external data.");
                return;
            }

            if ((!string.IsNullOrEmpty(this.Path.Text) | !string.IsNullOrEmpty(this.oldLocation.Text) | !string.IsNullOrEmpty(this.Location.Text)))
            {
                strSourceFile = this.Path.Text;
                UpdateConnectPath(strSourceFile);
            }
        }

        private void UpdateConnectPath(string strSourceFile)
        {
            
            XmlNode oxmlNode = null;
            
            try
            {
                // Open the workbook.
                SpreadsheetDocument wkb = SpreadsheetDocument.Open(strSourceFile, true);

                // Manage namespaces to perform Xml XPath queries.
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("sh", spreadsheetmlNamespace);

                // Get the connections part from the package.
                XmlDocument xdoc = new XmlDocument(nt);
                
                // Load the XML in the part into an XmlDocument instance.
                xdoc.Load(wkb.WorkbookPart.ConnectionsPart.GetStream());

                // Find the odcFile attribute.
                oxmlNode = xdoc.SelectSingleNode("/sh:connections/sh:connection/@odcFile", nsManager);

                // Replace the old path with the new path.
                oxmlNode.Value = Strings.Replace(oxmlNode.Value, this.oldLocation.Text, this.Location.Text);
                xdoc.Save(wkb.WorkbookPart.ConnectionsPart.GetStream());
                               
                // Close workbook and exit.
                wkb.Close();
                MessageBox.Show("Operation Complete");
                System.Windows.Forms.Application.Exit();
               
		} 
        catch 
        {
			// Some files have no external connections so ignore any errors.
		}

	}
  }
}

