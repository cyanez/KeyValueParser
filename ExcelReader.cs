using System;
using System.Collections.Generic;
using System.Text;

using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excell
{
    class ExcelReader
    {
        private Application xlApp = new Application();
        private Workbook xlWorkbook;
        private Worksheet xlWorksheet;
        private Range usedRange;
        public int rowsCount;

        List<string> headers = new List<string>();

        private List<string> getRow(int rowNumber) 
        {
          List<string> row = new List<string>();

          if (rowNumber > 0) 
          {
          
            for (int i = 1; i <= usedRange.Columns.Count; i++)
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

        public bool IsRowEmpty(int rowNumber)
        {

          List<string> row = getRow(rowNumber);

          if(row.Count != 0)
          {
            return false;
          } else {
            return true;
          }

        }

        public ExcelReader(string filePath)
        {
           this.xlWorkbook = xlApp.Workbooks.Open(filePath);          
        }

        public void LoadWorkSheet(int workSheetNumber)
        {
          this.xlWorksheet = xlWorkbook.Sheets[workSheetNumber];
          this.usedRange = this.xlWorksheet.UsedRange;
          this.rowsCount = this.usedRange.Rows.Count;

          headers = getHeaders();
        }

        public List<string> GetWorkSheetsNames() {

          List<string> workSheetsNames = new List<string>();
          
          for(int i = 1; i <= this.xlWorkbook.Worksheets.Count; i++) {

            Worksheet currentWorkSheet = this.xlWorkbook.Worksheets[i];
            workSheetsNames.Add(currentWorkSheet.Name);

          }
            
          return workSheetsNames;
        }        
           
        public List<Field> getRowFields(int row)
        {
            List<Field> list = new List<Field>();
                         
            for (int i = 2; i < headers.Count; i++)
            { 
                if (usedRange.Cells[row, i + 1] != null && usedRange.Cells[row, i + 1].Value2 != null)
                {
                    Field field = new Field(usedRange.Cells[row, 1].Value2 != null ? usedRange.Cells[row, 1].Value2.ToString() : "" , 
                                            usedRange.Cells[row, 2].Value2 != null ? usedRange.Cells[row, 2].Value2.ToString() : "",
                                            headers[i],
                                            usedRange.Cells[row, i + 1].Value2.ToString());
                    list.Add(field);
                }

            }

            return list;            
        }

        private void CleanMemory() 
        { 
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
