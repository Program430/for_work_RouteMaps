using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using System.Xml.Linq;
using Spire.Xls.Core;
using Spire.Pdf.Exporting.XPS.Schema;
using System.Runtime.Remoting.Messaging;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using OfficeOpenXml.Drawing;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {  
        public Form1()
        {       
            InitializeComponent(); 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            File.Delete("Duplicatesheet1.xlsx");
            File.Copy("Duplicatesheet.xlsx", "Duplicatesheet1.xlsx");

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //provide file path
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filePath = openFileDialog1.FileName;

            string systemfile = "System1.xlsx";
            ExcelPackage excelPackage = new ExcelPackage(systemfile);
            using (ExcelPackage package = new ExcelPackage(filePath))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                int i = 0;
                for (int row = 2 ; row <= rowCount; row = row + 4)
                {
                    for (int row1 = 0; row1 < 4; row1++)
                    {
                            string l1,l2,l3,l4,l5,l6;

                            try
                            {
                                l1 = worksheet.Cells[row + row1, 4].Value.ToString();
                                l2 = worksheet.Cells[row + row1, 3].Value.ToString();
                                l3 = worksheet.Cells[row + row1, 8].Value.ToString();
                                l4 = worksheet.Cells[row + row1, 6].Value.ToString();
                                l5 = worksheet.Cells[row + row1, 5].Value.ToString();
                                l6 = worksheet.Cells[row + row1, 11].Value.ToString();
                            }
                            catch
                            {
                                l1 = l2 = l3 = l4 = l5 = l6 = "";
                            }

                            //установить значения листа нового

                            ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets[0];

                            //строка колонка
                            worksheet1.Cells[1 + row1 * 10, 1].Value = l1;
                            worksheet1.Cells[2 + row1 * 10, 1].Value = l2;
                            worksheet1.Cells[2 + row1 * 10, 11].Value = l3;
                            worksheet1.Cells[6 + row1 * 10, 4].Value = l4;
                            worksheet1.Cells[7 + row1 * 10, 4].Value = l4;
                            worksheet1.Cells[4 + row1 * 10, 7].Value = l5;
                        switch (l6)
                            {
                                case "06":
                                case "0б":
                                    l6 = "Люлька";
                                    break;

                                case "00":
                                    l6 = "Общ Вид";
                                    break;

                                case "01":
                                    l6 = "РамаОп";
                                    break;
                                case "02":
                                    l6 = "рамаПов";
                                    break;
                                case "03":
                                    l6 = "Стрела";
                                    break;
                                case "04":
                                    l6 = "Электр";
                                    break;
                                case "05":
                                    l6 = "Гидравл";
                                    break;
                                case "0700":
                                    l6 = "Г.цилОбщВида";
                                    break;
                                case "0701":
                                    l6 = "Г.ЦилРамыОп";
                                    break;
                                case "703":
                                    l6 = "Г.цилСтр";
                                    break;
                                case "706":
                                    l6 = "Г.цилЛюльки";
                                    break;
                                case "702":
                                    l6 = "Г.ЦилРамыПов";
                                    break;
                                default:
                                    l6 = "ДАННЫЕ НЕ РАСПОЗНАНЫ";
                                    break;
                                    
                            }
                            worksheet1.Cells[2 + row1 * 10, 13].Value = l6;
                     
                    }
                    excelPackage.Save();

                    //Load the sample Excel
                    Workbook workbook = new Workbook();
                    workbook.LoadFromFile(systemfile);
                    Workbook workbook2 = new Workbook();
                    workbook2.LoadFromFile("Duplicatesheet1.xlsx");
                    //Add worksheet and set its name
                    workbook2.Worksheets.Add("Copy-"+ i.ToString());

                    //copy worksheet to the new added worksheets
                    workbook2.Worksheets[i+1].CopyFrom(workbook.Worksheets[0]);
                     
                    //Save the Excel workbook.
                    workbook2.SaveToFile("Duplicatesheet1.xlsx");

                    //напечатать лист
                    Workbook wb = new Workbook();

                    //Load an Excel file
                    wb.LoadFromFile(systemfile);
                    
                    i++;
                }
            }
            Process.Start("C:/Program Files/Microsoft Office/root/Office16/EXCEL", "Duplicatesheet1.xlsx");
        }
    }
}
