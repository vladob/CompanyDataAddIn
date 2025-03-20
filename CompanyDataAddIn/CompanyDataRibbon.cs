using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace CompanyDataAddIn
{
    public partial class CompanyDataRibbon
    {
        private void CompanyDataRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void FetchDataButton_Click(object sender, RibbonControlEventArgs e)
        {
            string cin = CinTextBox.Text;
            if (!string.IsNullOrEmpty(cin))
            {
                // Call the method to fetch and process data
                Globals.ThisAddIn.FetchAndInsertData(cin);
            }
            else
            {
                MessageBox.Show("Please enter a valid CIN (IČO).");
            }
        }
    }
}
