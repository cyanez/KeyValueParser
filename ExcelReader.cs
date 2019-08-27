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

        List<string> headers = new List<string>();
       

        public ExcelReader(string filePath, int sheetNumber)
        {
            this.filePath = filePath;

            this.xlWorkbook = xlApp.Workbooks.Open(this.filePath);
            this.xlWorksheet = xlWorkbook.Sheets[sheetNumber];

            headers = getHeaders();
        }

        public Range getUsedRange()
        {
            return this.xlWorksheet.UsedRange;
        }

        public List<string> getRow(int rowNumber)
        {
            List<string> row = new List<string>();

            if (rowNumber > 0)
            {
                Range usedRange = getUsedRange();
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
           
            Range usedRange = getUsedRange();

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

    }
}
