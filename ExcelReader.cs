using System;
using System.Collections.Generic;

using System.Text;
using System.Threading.Tasks;


using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excell
{
    class ExcelReader
    {
        string filePath = "";

        Application xlApp = new Application();
        Workbook xlWorkbook;
        Worksheet xlWorksheet;
        Range usedRange;

        List<string> headers = new List<string>();
       

        public ExcelReader(string filePath, int sheetNumber)
        {
            this.filePath = filePath;

            this.xlWorkbook = xlApp.Workbooks.Open(this.filePath);
            this.xlWorksheet = xlWorkbook.Sheets[sheetNumber];
            usedRange = this.xlWorksheet.UsedRange;

            headers = getHeaders();
            
        }

        

        private List<string> getRow(int rowNumber)
        {
            List<string> row = new List<string>();

            if (rowNumber > 0)
            {
                
                int colCount = usedRange.Columns.Count;

                for (int i = 1; i <= colCount; i++)
                {
                    if (usedRange.Cells[rowNumber, i] != null && usedRange.Cells[rowNumber, i].Value2 != null)
                    {
                        row.Add(usedRange.Cells[rowNumber, i].Value2.ToString());
                    }

                }

            }

            return row;

        }
               

        private List<string> getHeaders()
        {
            return getRow(1);
        }
           
        public List<Field> getFields(int row)
        {
            List<Field> list = new List<Field>();
           
            
            for (int i = 2; i < headers.Count; i++)
            {
                if (usedRange.Cells[row, i + 1] != null && usedRange.Cells[row, i + 1].Value2 != null)
                {
                    Field field = new Field(usedRange.Cells[row, 1].Value2.ToString(),
                                            usedRange.Cells[row, 2].Value2.ToString(),
                                            headers[i],
                                            usedRange.Cells[row, i + 1].Value2.ToString());
                    list.Add(field);
                }


            }

            return list;
            
        }

        private void CleanMemory() {

          GC.Collect();
          GC.WaitForPendingFinalizers();

        }
        
        public void Release()
        {

            CleanMemory();
                       
            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(xlWorksheet);
                       
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
                      
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

        }

    }
}
