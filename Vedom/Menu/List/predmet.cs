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
    public partial class predmet : Form
    {
        public predmet()
        {
            InitializeComponent();
        }

        private void predmet_Load(object sender, EventArgs e)
        {
            LoadDataFromExcel();
        }

        DataTable dt = new DataTable();

        private void LoadDataFromExcel()
        {
            string fileName = "vedom.xlsx";
            string studentsSheetName = "дисциплины";
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
                    
                    dt.Columns.Add("Семестр");
                    dt.Columns.Add("Название");
                    dt.Columns.Add("Тип оценивания");
                    dt.Columns.Add("Преподаватель");
                    

                    // Добавление данных о семестре для каждой записи
                    for (int i = 2; i <= worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        DataRow row = dt.NewRow();                      
                        row["Семестр"] = worksheet.Cells[i, 1].Value;
                        row["Название"] = worksheet.Cells[i, 2].Value;
                        row["Тип оценивания"] = worksheet.Cells[i, 3].Value;
                        row["Преподаватель"] = worksheet.Cells[i, 4].Value;

                        // Пример установки значения семестра для каждой записи
                        
                        dt.Rows.Add(row);
                    }

                    // Отображаем данные в DataGridView
                    dataGridView1.DataSource = dt;
                }

                workbook.Close();
                excelApp.Quit();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            dataGridView1.Columns[0].Width = 80;


            List<string> semestersList = new List<string>();
            semestersList.Add("");
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow) // проверяем, что это не новая строка
                {
                    string semester = row.Cells["Семестр"].Value.ToString();
                    if (!string.IsNullOrEmpty(semester) && !semestersList.Contains(semester))
                    {
                        semestersList.Add(semester);
                    }
                }
            }
            comboBox1.DataSource = semestersList;
            if (Properties.Settings.Default.semsestSave != null)
            {
                comboBox1.SelectedItem = Properties.Settings.Default.semsestSave;
            }

        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataView dv = dt.DefaultView; // Создаем DataView для фильтрации данных
            if (comboBox1.SelectedItem == null || comboBox1.SelectedItem.ToString() == "")
            {
                // Если ни один элемент не выбран или выбрано пустое значение в ComboBox1, отображаем все записи
                dv.RowFilter = ""; // Сбрасываем фильтр
            }
            else
            {
                // Создаем строку для фильтрации по выбранным значениям из ComboBox1
                StringBuilder filter = new StringBuilder();
                filter.Append("Семестр IN (");

                filter.Append("'" + comboBox1.SelectedItem.ToString() + "',");

                // Удаляем последнюю запятую
                filter.Remove(filter.Length - 1, 1);
                filter.Append(")");

                // Применяем фильтр к DataView
                dv.RowFilter = filter.ToString();
                Properties.Settings.Default.semsestSave = comboBox1.SelectedItem.ToString();
                Properties.Settings.Default.Save();
            }

            // Обновляем источник данных DataGridView с учетом фильтра
            dataGridView1.DataSource = dv.ToTable();
        }

        private void save_Click(object sender, EventArgs e)
        {
            string fileName = "vedom.xlsx";
            string studentsSheetName = "дисциплины";
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

                worksheet.Cells[1, 1] = "Семестр";
                worksheet.Cells[1, 2] = "Название";
                worksheet.Cells[1, 3] = "Тип оценивания";
                worksheet.Cells[1, 4] = "Преподаватель";


                if (worksheet != null)
                {
                    // Получаем данные из DataGridView
                    DataTable dt = (DataTable)dataGridView1.DataSource;

                    // Записываем данные в лист Excel
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1] = dt.Rows[i]["Семестр"];
                        worksheet.Cells[i + 2, 2] = dt.Rows[i]["Название"];
                        worksheet.Cells[i + 2, 3] = dt.Rows[i]["Тип оценивания"];
                        worksheet.Cells[i + 2, 4] = dt.Rows[i]["Преподаватель"];
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
