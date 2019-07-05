using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;


namespace POCDD
{
    public partial class Dashboard : Form
    {
        public Dashboard()
        {
            InitializeComponent();
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {

        }

        private void fpSpread1_CellClick(object sender, FarPoint.Win.Spread.CellClickEventArgs e)
        {
            //if (fpSpread2_Sheet1.IsSelected(0, 3))
            //{
            //    fpSpread2_Sheet1.Cells.Get(3, 4).Value = 211100;
            //}
        }

        private void fpSpread2_CellClick(object sender, FarPoint.Win.Spread.CellClickEventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void exportToExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            DialogResult result = this.saveFileDialog1.ShowDialog();
            
            if (result == DialogResult.OK)
            {
               

                    fpSpread2.SaveExcel(this.saveFileDialog1.FileName  , FarPoint.Win.Spread.Model.IncludeHeaders.BothCustomOnly);
                
            }
            
           // fpSpread2.Open("C:\\spreadfile.xml"); 
        }

        private void fpSpread2_ComboCloseUp(object sender, FarPoint.Win.Spread.EditorNotifyEventArgs e)
        {
           
            //    if (fpSpread2_Sheet1.IsSelected(0, 3))
            //{
            //    fpSpread2_Sheet1.Cells.Get(3, 4).Value = 211100;
            //}
            //else if (fpSpread2_Sheet1.IsSelected(1, 3))
            //{
            //    fpSpread2_Sheet1.Cells.Get(3, 4).Value = 9999;
            //}
            //else if (fpSpread2_Sheet1.IsSelected(0, 4))
            //{
            //    fpSpread2_Sheet1.Cells.Get(3, 4).Value = 1111;
            //}
            //FarPoint.Win.Spread.CellType.ComboBoxCellType comboBoxCellType1 = (FarPoint.Win.Spread.CellType.ComboBoxCellType)fpSpread2.ActiveSheet.ActiveCell.CellType;
            //FarPoint.Win.Spread.CellType.ComboBoxCellType comboBoxCellType2 = (FarPoint.Win.Spread.CellType.ComboBoxCellType)fpSpread2.ActiveSheet.ActiveCell.CellType;
            ////c.StopEditing();
            ////SendKeys.Send("\r");
            ////FarPoint.Win.Spread.CellType.ComboBoxCellType comboBoxCellType1 = new FarPoint.Win.Spread.CellType.ComboBoxCellType();
            //if (!comboBoxCellType1.Equals(""))
            //    {
            //    fpSpread2_Sheet1.Cells.Get(3, 2).Value = 200;
            //}

            //if (!comboBoxCellType2.Equals(""))
            //{
            //    fpSpread2_Sheet1.Cells.Get(3, 4).Value = 211100;
            //}

        }
       
        private void fpSpread2_ComboSelChange(object sender, FarPoint.Win.Spread.EditorNotifyEventArgs e)
        {
            //if (fpSpread2_Sheet1.IsSelected(0, 3))
            //{
            //    fpSpread2_Sheet1.Cells.Get(3, 4).Value = 211100;
            //}
            //else if (fpSpread2_Sheet1.IsSelected(1, 3))
            //{
            //    fpSpread2_Sheet1.Cells.Get(3, 4).Value = 9999;
            //}
            //else if (fpSpread2_Sheet1.IsSelected(0, 4))
            //{
            //    fpSpread2_Sheet1.Cells.Get(3, 4).Value = 1111;
            //}
        }

        private void fpSpread2_ActiveSheetChanged(object sender, EventArgs e)
        {
            if (fpSpread2.ActiveSheetIndex == 1)
            {
                fpSpread2_Sheet1.Cells.Get(3, 4).Value = 211100;
            }
            else if (fpSpread2.ActiveSheetIndex == 2)
            {
                fpSpread2_Sheet1.Cells.Get(3, 4).Value = 1111;
            }
            
        }

        private void exportToPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
           

            //saveFileDialog1.Title = "Save Spread to PDF Format";
            //saveFileDialog1.Filter = "Pdf File |*.pdf";
            ////if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            ////{

            ////    FarPoint.Win.Spread.PrintInfo pi = new FarPoint.Win.Spread.PrintInfo();
            ////    pi.PrintToPdf = true;
            ////    pi.PdfWriteMode = FarPoint.Win.Spread.PdfWriteMode.New;
            ////    pi.PdfWriteTo = FarPoint.Win.Spread.PdfWriteTo.File;
            ////    pi.PdfFileName = saveFileDialog1.FileName;
            ////    fpSpread2.Sheets[fpSpread2.ActiveSheetIndex].PrintInfo = pi;
            ////    fpSpread2.PrintSheet(fpSpread2.ActiveSheetIndex);

            ////}
            //Workbook workbook = new Workbook();
            //////workbook.LoadFromFile("Sample.xlsx");
            ////Worksheet sheet = workbook.Worksheets[fpSpread2.ActiveSheetIndex];
            ////sheet.SaveToPdf("toPDF.pdf");
            ////System.Diagnostics.Process.Start("toPDF.pdf");

            ////fpSpread2.Save(this.saveFileDialog1.FileName, true);
            //workbook.SaveToFile("result.pdf", Spire.Xls.FileFormat.PDF);
        }
    }
}
