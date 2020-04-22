using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Vocabulary
{
    public partial class RecordNewWord : Form
    {
        public RecordNewWord()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RecordWord();
            InputEntxt.Text = "";
            InputCntxt.Text = "";
        }

        public void RecordWord()
        {
            Excel excel = new Excel("D:\\Vocabulary.xlsx", 1);
            int excelRowNum = excel.RowNum();
            excel.WriteCell(excelRowNum + 1, 1, InputEntxt.Text);
            excel.WriteCell(excelRowNum + 1, 2, InputCntxt.Text);
            excel.CloseExcel();
        }
    }
}
