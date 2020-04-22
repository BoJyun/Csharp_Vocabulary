using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel; // Add excel library

namespace Vocabulary
{
    class Excel   //資源:youtube.com/user/TheBospear/playlists
    {
        string path;
        _Application excel = new Microsoft.Office.Interop.Excel.Application(); //設定Excel應用程序
        Workbook wb; //設定Excel檔案
        Worksheet ws; //設定工作表

        public Excel() //建構子 for 建立新excel
        {

        }

        public Excel(string path, int sheet) // 建構子 for Open and Read Excel
        {
            this.path = path;
            this.wb = excel.Workbooks.Open(path);
            this.ws = wb.Worksheets[sheet];
        }

        public string ReadCell(int i,int j)  // 讀取Excel 的值
        {
           // i++;
           // j++;
            if (ws.Cells[i, j].Value != null)    //Value 和 Value2 要探索清楚其不同之處
                return ws.Cells[i, j].Value.ToString();
            else
                return "";    
        }

        public string[,] ReadRange(int starti,int starty,int endi,int endy)  //讀取excel範圍
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[endi - starti, endy - starty]; //矩陣是從[0,0]開始

            for (int p = 1; p <= endi - starti; p++)
                for (int q = 1; q <= endy - starty; q++)
                    returnstring[p - 1, q - 1] = holder[p, q].ToString();

            return returnstring;
        }

        public void WriteCell(int i,int j,string s)   //Write Excel Files
        {
           // i++;
           // j++;
            ws.Cells[i, j].Value = s;
        }

        public void SaveFile()  //Save Excel Files
        {
            wb.Save();
        }

        public void SaveAsFile(string path)  //SaveAs Excel Files
        {
            wb.SaveAs(path);
        }

        public void CloseExcel() //關閉excel 記得都要關，不然會一直開著，可能導致唯獨
        {
            wb.Close();
        }

        public void CreartNewFile()  //建立新excel，要多建立一個建構子
        {                               
           this.wb = excel.Workbooks.Add();  //wb要用this指定告知，因為'建構子'那邊沒有指定
           this.ws = wb.Worksheets[1];
        }

        public void CreatNewSheet() //建立新sheet，建立在原ws之後，所以上面的CreatNewFile的ws要用 this指定
        {                           
            wb.Worksheets.Add(After: ws);
        }

        public int RowNum()   //判別該excel目前有效使用了多少"行"
        {
            int rowNum = ws.UsedRange.Cells.Rows.Count;
            return rowNum;
        }

        public int ColNum()   //判別該excel目前有效使用了多少"列"
        {
            int colNum = ws.UsedRange.Cells.Columns.Count;
            return colNum;
        }
    }
}
