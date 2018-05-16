using System;
using System.Windows.Forms;

namespace XMLParse1
{
    public partial class ExcelHeaderReuirements : Form
    {
        public ExcelHeaderReuirements()
        {
            InitializeComponent();
        }

        private void ExcelHeaderReuirements_Load(object sender, EventArgs e)
        {
            textBox1.Text = 
                "Excel must be sorted in ascending order of 'Left'.\r\n"+ 
                "Column order does not matter.\r\n" +
                "Header names matter and must match those below, case insensitive.\r\n"+
                "'Left' and 'Right' are not required in the Excel file. (assuming data is sorted)\r\n\r\n"+
                "level\r\ncategory hierarchy: name\r\ncategory hierarchy: id\r\nleft\r\nright";
        }
    }
}
