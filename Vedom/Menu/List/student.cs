using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Vedom.Menu.List
{
    public partial class student : Form
    {
        private DataTable dt;
        public student()
        {
            InitializeComponent();
        }

        private void student_Load(object sender, EventArgs e)
        {
            LoadDataFromExcel();
        }

        private void LoadDataFromExcel()
        {
            dataGridView1.Visible = false;
            label1.Visible = true;
            string fileName = "vedom.xlsx";
            string studentsSheetName = "студенты";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

                // Проверяем существует ли файл
                if (!File.Exists(filePath))
                {
                    // Если файл не существует, создаем новый
                    workbook = excelApp.Workbooks.Add();
                    workbook.SaveAs(filePath);
                }
                else
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при открытии файла: " + ex.Message);
            }

            if (workbook != null)
            {
                Excel.Worksheet worksheet = null;

                // Проверяем существует ли лист "студенты"
                bool studentsSheetExists = false;
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == studentsSheetName)
                    {
                        worksheet = sheet;
                        studentsSheetExists = true;
                        break;
                    }
                }

                // Если лист "студенты" не существует, создаем его
                if (!studentsSheetExists)
                {
                    worksheet = workbook.Sheets.Add();
                    worksheet.Name = studentsSheetName;
                    workbook.Save(); // Сохраняем изменения в файле
                }

                if (worksheet != null)
                {                   
                    // Создаем DataTable для хранения данных
                    DataTable dt = new DataTable();
                    dt.Columns.Add("№");
                    dt.Columns.Add("ФИО");

                    // Загружаем данные из листа Excel в DataTable
                    for (int i = 2; i <= worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        DataRow row = dt.NewRow();
                        row["№"] = worksheet.Cells[i, 1].Value;                      
                        row["ФИО"] = worksheet.Cells[i, 2].Value;
                        dt.Rows.Add(row);
                    }

                    // Отображаем данные в DataGridView
                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns["ФИО"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;                              
                }

                workbook.Close();
                excelApp.Quit();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Visible = true;
            label1.Visible = false;
            dataGridView1.AllowUserToAddRows = false;
        }

       

        private void save_Click(object sender, EventArgs e)
        {
            string fileName = "vedom.xlsx";
            string studentsSheetName = "студенты";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

                // Проверяем существует ли файл
                if (!File.Exists(filePath))
                {
                    // Если файл не существует, создаем новый
                    workbook = excelApp.Workbooks.Add();
                    workbook.SaveAs(filePath);
                }
                else
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при открытии файла: " + ex.Message);
            }

            if (workbook != null)
            {
                Excel.Worksheet worksheet = null;

                // Проверяем существует ли лист "студенты"
                bool studentsSheetExists = false;
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == studentsSheetName)
                    {
                        worksheet = sheet;
                        studentsSheetExists = true;
                        break;
                    }
                }

                // Если лист "студенты" не существует, создаем его
                if (!studentsSheetExists)
                {
                    worksheet = workbook.Sheets.Add();
                    worksheet.Name = studentsSheetName;
                    workbook.Save(); // Сохраняем изменения в файле
                }
                
                worksheet.Cells[1, 1] = "№";
                worksheet.Cells[1, 2] = "ФИО";

                if (worksheet != null)
                {
                    // Заполняем столбец "№" от 1 до 25
                    for (int i = 1; i <= 25; i++)
                    {
                        worksheet.Cells[i + 1, 1] = i;
                    }

                    // Получаем данные из DataGridView
                    DataTable dt = (DataTable)dataGridView1.DataSource;

                    // Записываем данные в столбец "ФИО"
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        worksheet.Cells[i + 2, 2] = dt.Rows[i]["ФИО"];
                    }

                    // Сохраняем изменения в файле
                    workbook.Save();
                }

                workbook.Close();
                excelApp.Quit();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            MessageBox.Show("Данные сохранены в Excel файл!");
        }

       
    }
}
