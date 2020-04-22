using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;


namespace Vocabulary
{
    public partial class Form1 : Form
    {
        private ReadWord stratgo = new ReadWord();
        
        public Form1()
        {
            InitializeComponent();
        }
        
        private void Start_Click(object sender, EventArgs e) // 產生要背的5個單字
        {
            Word[] Readwords = MakeNewWord();
            GetEnlbl1.Text = Readwords[0].Enword;
            GetCnlbl1.Text = Readwords[0].Cnword;
            GetEnlbl2.Text = Readwords[1].Enword;
            GetCnlbl2.Text = Readwords[1].Cnword;
            GetEnlbl3.Text = Readwords[2].Enword;
            GetCnlbl3.Text = Readwords[2].Cnword;
            GetEnlbl4.Text = Readwords[3].Enword;
            GetCnlbl4.Text = Readwords[3].Cnword;
            GetEnlbl5.Text = Readwords[4].Enword;
            GetCnlbl5.Text = Readwords[4].Cnword;

        }

        private void Record_Click(object sender, EventArgs e)  //記錄新的單字
        {
            RecordNewWord _form = new RecordNewWord();
            _form.Show();
            _form = null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string path = "D:";
            if (Directory.Exists(path) == false)
            {
                Directory.CreateDirectory(path);
            }

            string pathExecl = path + "\\Vocabulary.xlsx";
            if (System.IO.File.Exists(pathExecl)==false)   //判斷Excel是否存在
                CreatNewFile();
        }

        //public void OpenFile()
        //{
        //    Excel excel = new Excel("D:\\coding_practice\\C#diary\\Vocabulary_Excel-2th\\Vocabulary\\folder\\123.xlsx", 1);
        //    excel.WriteCell(1, 0, "Apple");
        //    excel.SaveFile();
        //    excel.CloseExcel();
        //    //  excel.SaveAsFile(@"Texr2.xlsx");
        //}

        public void CreatNewFile()
        {
            Excel excel = new Excel();
            excel.CreartNewFile();
            //excel.CreatNewSheet();  //建立新sheet
            excel.SaveAsFile("D:\\Vocabulary.xlsx");
            excel.CloseExcel();
        }

        public void ReadRange(out int excelRowNum, out int excelColNum) //檢查所開啟的excel，目前有效使用了多少"行"、"列" 
        {
            Excel excel = new Excel("D:\\Vocabulary.xlsx", 1);
            excelRowNum = excel.RowNum();
            excelColNum = excel.ColNum();
            excel.CloseExcel();
        }

        public Word[] MakeNewWord()
        {
            Excel excel = new Excel("D:\\Vocabulary.xlsx", 1);
            Word[] words = new Word[5]; //一次產生5個不同單字

            Random rnd = new Random(); //產生亂數，此亂數代表Excel中要拉出第幾列的數值

            int sumExcelRowNum = excel.RowNum();
            int[] arrayrnd = new int[5]; //用來暫存所產生的亂數，稱'指向execl'矩陣
            int arraylength = 0;
            int NewRnd;

            while (arraylength != 5)   //產生亂數，此亂數代表Excel中要拉出第幾列的數值，因為不想重複，所以要檢測亂數有沒有重複
            {
                bool ChectNewRnd = false; //多一個bool來判別亂數使否有重複
                NewRnd = rnd.Next(1, sumExcelRowNum + 1); 
                for (int k = 0; k < 5; k++)  //檢查新產生的亂數有沒有重複
                {
                    if (arrayrnd[k] == NewRnd) //找到該亂數已經存在於暫存的亂數矩陣裡
                    {
                        ChectNewRnd = false;  //因為有重複，所以該次失敗且跳出for迴圈，重新產生一個新亂數
                        break;
                    }
                    else
                    {
                        ChectNewRnd = true;
                    }
                }

                if (ChectNewRnd == true)  //到此若bool是true，則表示該亂數並沒有重複，所以放進'指向excel'矩陣
                {
                    arrayrnd[arraylength] = NewRnd;
                    arraylength++;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                words[i] = stratgo.MakeWord(excel.ReadCell(arrayrnd[i], 1), excel.ReadCell(arrayrnd[i], 2));
            }

            excel.CloseExcel();

            return words;
        }

    }
}
