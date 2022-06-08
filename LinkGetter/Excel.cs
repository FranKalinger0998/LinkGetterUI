using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace LinkGetter
{
    class Excel
    {
        Uri path;
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(Uri path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path.ToString());
            ws = wb.Worksheets[sheet];
        }
        public string readCell(int i, int j)
        {

            i++; j++;
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return "";
        }
        public void writeToCell(int i, int j, string s)
        {
            i++; j++;
            ws.Cells[i, j].Value2 = s;
        }
        public void writeALinkToCell(int i, int j, string s)
        {
            i++; j++;
            ws.Cells[i, j].Hyperlink = new Uri(s);
        }
        public void save()
        {
            wb.Save();
        }
        public void wbclose()
        {
            wb.Close();
        }


    }
}



