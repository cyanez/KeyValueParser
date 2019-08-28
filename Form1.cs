using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace Excell
{
    public partial class frmParser : Form
    {
        string filePath = string.Empty;

        ExcelReader excel;

        public frmParser()
        {
            InitializeComponent();
        }

        private void BtnOpenExcelFile_Click(object sender, EventArgs e) {

          using (OpenFileDialog openExcelFileDialog = new OpenFileDialog()) {

            openExcelFileDialog.Title = "Seleccionar archivo Excel";
            openExcelFileDialog.InitialDirectory = @"c:\";
            openExcelFileDialog.Filter = "Excel files |*.xlsx;*,xlsx";

            if (openExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
              filePath = openExcelFileDialog.FileName;

              FillcbxWorkSheets();
            }

          }

        }

        private void FillcbxWorkSheets()
        {
           excel = new ExcelReader(this.filePath);
           List<string> workSheetsNames = excel.GetWorkSheetsNames();
           cbxWorkSheets.DataSource = workSheetsNames;
        }

        private int GetSelectedWorkSheet() 
        {
          return cbxWorkSheets.SelectedIndex + 1;
        }

        private void Button1_Click(object sender, EventArgs e)
            {
              if (filePath != string.Empty)
              {
                int workSheetSelectedNumber = GetSelectedWorkSheet();

                LoadFields(workSheetSelectedNumber);        
              }  else {
                MessageBox.Show("Es necesario seleccionar el archivo excel para procesarlo!!! ");
              }     

        }

        private bool isValidRow(List<Field> row, int rowNumber) 
        {
           if(row[0].UID == string.Empty)  
           {
              rtxtWarnings.AppendText("Warning: " + "El campo UID se encuentra en blanco del renglón número: " + rowNumber.ToString() + " \n");
              return false;
           } else if(row[0].ObjectType == string.Empty)
           {
              rtxtWarnings.AppendText("Warning: " + "El campo ObjectType se encuentra del renglón número: " + rowNumber.ToString() + " \n");
              return false;
           } else {
              return true;
           }
          
        }

        
        public void LoadFields(int sheetNumber)
        {
            excel.LoadWorkSheet(sheetNumber);
                       
           int emptyRowsCount = 0;
           int rowsAddedCount = 0;

            for (int i = 2; i < excel.rowsCount; i++)
            {
              if (excel.IsRowEmpty(i)) 
              {
                emptyRowsCount++;

                rtxtWarnings.AppendText("El renglón número: " + i.ToString() + " se encentra en blanco \n");

                if (emptyRowsCount == 5) 
                {
                  break;
                }
              } else
              {
                emptyRowsCount = 0;
                
                List<Field> row = excel.getRowFields(i);
                if (isValidRow(row, i))
                {
                  foreach (Field field in row)
                  {
                    field.AddField();                   
                  }

                  rowsAddedCount++;
                  rtxtNotifications.Text = "Se agregó correctamente el renglón número: " + i.ToString() +
                                    "\n" + "\n" +
                                    rowsAddedCount.ToString() + " Renglones agregados...";

                }            
                
              }               

            }           
            
            MessageBox.Show("Se agregó la tabla");        
        }

        private void FrmParser_FormClosed(object sender, FormClosedEventArgs e) {
          excel.Release();
        }
  }
}
