using System.Windows;
using System.Windows.Controls;

namespace CompanyDataAddIn
{
    public partial class RibbonControl : UserControl
    {
        public RibbonControl()
        {
            InitializeComponent();
        }

        private void FetchDataButton_Click(object sender, RoutedEventArgs e)
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