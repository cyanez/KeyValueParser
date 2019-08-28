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
    public partial class Form1 : Form
    {
    string filePath = string.Empty;    
    
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnOpenExcelFile_Click(object sender, EventArgs e) {

          using (OpenFileDialog openExcelFileDialog = new OpenFileDialog()) {

            openExcelFileDialog.Title = "Seleccionar archivo Excel";
            openExcelFileDialog.InitialDirectory = @"c:\";
            openExcelFileDialog.Filter = "Excel files |*.xlsx;*,xlsx";

            if (openExcelFileDialog.ShowDialog() == DialogResult.OK) {
              filePath = openExcelFileDialog.FileName;
            }

          }

        }

        private void Button1_Click(object sender, EventArgs e)
            {
              if (filePath != string.Empty) {
                LoadFields(filePath, 2, 2);
              }  else {
                MessageBox.Show("Es necesario seleccionar el archivo excel para procesarlo!!! ");
              }     

        }

        public void LoadFields(string fileName, int sheetNumber, int columns)
        {
            ExcelReader excel = new ExcelReader(fileName, sheetNumber);

            int count = columns; // excel.RowsCount;

            for (int i = 0; i < count; i++)
            {
                List<Field> headers = excel.getFields(i + 2);

                foreach (Field field in headers)
                {
                    field.AddField();
                }

            }
            
            excel.Release();

            MessageBox.Show("Se agregó la tabla");
        
        }

   


  }
}
