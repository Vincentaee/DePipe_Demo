using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExReaderConsole
{
    class ExReader
    {
        
        string exlocatiion;
        Excel.Application xlApp;
        Excel.Workbooks xlWorkbooks;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;

        int rowCount;
        int colCount;
        int startRow;
        int endRow;

        //DE part
        public List<Tuple<double, double>> excaRange = new List<Tuple<double, double>>();   // X Y
        public List<Tuple<int, double>> excaLevel = new List<Tuple<int, double>>();   // 階數 深度
        public List<Tuple<int, double, int, string, int>> supLevel = new List<Tuple<int, double, int, string, int>>();
        public List<double> centralCol = new List<double>();
        public double wall_width;

        //MH part
        public List<List<string>> MHdata = new List<List<string>>();

        public ExReader() {
            
        }
           
        public void SetData(string file, int page)
        {            
            xlApp = new Excel.Application();
            xlWorkbooks = xlApp.Workbooks;
            xlWorkbook = xlWorkbooks.Open(file);

            xlWorksheet = xlWorkbook.Sheets[page];
            xlRange = xlWorksheet.UsedRange;
            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;
        }

        void SetPage(int page)
        {
          
            xlWorksheet = xlWorkbook.Sheets[page];
            xlRange = xlWorksheet.UsedRange;
            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;
        }
        public void Change_color(Tuple<int, int> tuple)
        {
            xlWorksheet.Cells[tuple.Item1, tuple.Item2].Interior.Color = Excel.XlRgbColor.rgbRed;
            
        }
        public void Save_excel(string file)
        {
            
            xlWorkbook.SaveAs(String.Format(@"{0}",file));
            
        }
        public void PassDE()
        {
            var pos = this.FindAddress("連續壁厚度");
            wall_width = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;

            pos = this.FindAddress("開挖範圍");
            int i = 1;
            do
            {
                var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                excaRange.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            pos = this.FindAddress("開挖階數");
            i = 1;
            do
            {
                var data = Tuple.Create((int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                excaLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            pos = this.FindAddress("支撐階數");
            i = 1;
            do
            {
                var data = Tuple.Create((int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2,
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2);
                supLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            SetPage(2);
            pos = this.FindAddress("中間樁");
            for(int j = 1; j != 7; j++)
            {
                double d = xlRange.Cells[pos.Item1 + 1, pos.Item2 + j].Value2;
                centralCol.Add(d);
            }
            SetPage(1);
        }

        public void PassSSD()
        {

            startRow = 3;
            endRow = startRow;

            while (xlRange.Cells[endRow, 2].Value2 != null)
                endRow++;
            Console.WriteLine(startRow + " " + endRow + " " + colCount);


            for (int i = startRow; i != endRow; i++)
            {
                List<string> data = new List<string>();

                for (int j = 1; j <= colCount; j++)
                {

                    if (xlRange.Cells[i, j].Value2 != null)
                        data.Add(xlRange.Cells[i, j].Value2.ToString());
                    else
                        data.Add("");
                }
                MHdata.Add(data);
            }
        }

        public void PassMH()
        {
            startRow = 3;
            endRow = startRow;
            
            while (xlRange.Cells[endRow, 2].Value2 != null )
                endRow++;
            Console.WriteLine(startRow + " " + endRow+" " +colCount);


            for (int i =startRow; i!=endRow; i++)
            {
                List<string> data = new List<string>();
                
                for (int j = 1; j<= colCount; j++)
                {
                    
                    if (xlRange.Cells[i, j].Value2 != null)
                        data.Add(xlRange.Cells[i, j].Value2.ToString());
                    else
                        data.Add("");
                }
                MHdata.Add(data);
            }
        }

        public Tuple<int,int> FindAddress(string name)
        {
            Excel.Range address;
            address = xlRange.Find(name, MatchCase: true);
            var pos = Tuple.Create(address.Row, address.Column);
            return pos;
        }


        public void CloseEx()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlWorkbooks);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            
        }


    }

    
}
