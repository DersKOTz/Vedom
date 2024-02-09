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
    public partial class propusk : Form
    {
        public propusk()
        {
            InitializeComponent();
        }

        private void propusk_Load(object sender, EventArgs e)
        {
            LoadDataFromExcel();
        }

        private void LoadDataFromExcel()
        {
          
            string fileName = "vedom.xlsx";
            string studentsSheetName = "студенты";
            string attendanceSheetName = "прогулы";
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
                Excel.Worksheet studentsSheet = null;
                Excel.Worksheet attendanceSheet = null;

                // Проверяем существует ли лист "прогулы"
                bool attendanceSheetExists = false;
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == attendanceSheetName)
                    {
                        attendanceSheet = sheet;
                        attendanceSheetExists = true;
                        break;
                    }
                }

                // Если лист "прогулы" не существует, создаем его
                if (!attendanceSheetExists)
                {
                    attendanceSheet = workbook.Sheets.Add();
                    attendanceSheet.Name = attendanceSheetName;
                    workbook.Save(); // Сохраняем изменения в файле
                }

                // Проверяем существует ли лист "студенты"
                bool studentsSheetExists = false;
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == studentsSheetName)
                    {
                        studentsSheet = sheet;
                        studentsSheetExists = true;
                        break;
                    }
                }

                // Если лист "студенты" не существует, создаем его
                if (!studentsSheetExists)
                {
                    studentsSheet = workbook.Sheets.Add();
                    studentsSheet.Name = studentsSheetName;
                    workbook.Save(); // Сохраняем изменения в файле
                }

                if (studentsSheet != null && attendanceSheet != null)
                {
                    // Создаем DataTable для хранения данных
                    DataTable dt = new DataTable();
                    dt.Columns.Add("№");
                    dt.Columns.Add("ФИО");

                    for (int i = 1; i <= 31; i++)
                    {
                        dt.Columns.Add(i.ToString());
                    }

                    dt.Columns.Add("Всего");
                    dt.Columns.Add("Уваж.");
                    dt.Columns.Add("Неуваж.");

                    // Загружаем данные из листа "студенты" Excel в DataTable
                    for (int i = 2; i <= studentsSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        DataRow row = dt.NewRow();
                        row["№"] = studentsSheet.Cells[i, 1].Value;
                        row["ФИО"] = studentsSheet.Cells[i, 2].Value;
                        dt.Rows.Add(row);
                    }

                    // Загружаем данные из листа "прогулы" Excel в DataTable для столбцов 1-31
                    for (int i = 2; i <= attendanceSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        DataRow row = dt.Rows[i - 2];
                        for (int j = 1; j <= 31; j++)
                        {
                            row[j.ToString()] = attendanceSheet.Cells[i, j + 2].Value;
                        }

                        row["Всего"] = attendanceSheet.Cells[i, 34].Value; // Предполагается, что "Всего" находится в столбце 34
                        row["Уваж."] = attendanceSheet.Cells[i, 35].Value; // Предполагается, что "Уваж." находится в столбце 35
                        row["Неуваж."] = attendanceSheet.Cells[i, 36].Value; // Предполагается, что "Неуваж." находится в столбце 36
                    }

                    // Отображаем данные в DataGridView
                    dataGridView1.DataSource = dt;
                    // dataGridView1.Columns["ФИО"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }

                workbook.Close();
                excelApp.Quit();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[0].ReadOnly = true; 
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
        }

        private void save_Click(object sender, EventArgs e)
        {
            string fileName = "vedom.xlsx";
            string studentsSheetName = "прогулы";
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

                // Проверяем существует ли лист "прогулы"
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

                // Если лист "прогулы" не существует, создаем его
                if (!studentsSheetExists)
                {
                    worksheet = workbook.Sheets.Add();
                    worksheet.Name = studentsSheetName;
                    workbook.Save(); // Сохраняем изменения в файле
                }

                worksheet.Cells[1, 1] = "№";
                worksheet.Cells[1, 2] = "ФИО";

                // Записываем заголовки для столбцов с 1 по 31
                for (int i = 1; i <= 31; i++)
                {
                    worksheet.Cells[1, i + 2] = i.ToString();
                }

                worksheet.Cells[1, 34] = "Всего";
                worksheet.Cells[1, 35] = "Уваж.";
                worksheet.Cells[1, 36] = "Неуваж.";

                if (worksheet != null)
                {
                    // Получаем данные из DataGridView
                    DataTable dt = (DataTable)dataGridView1.DataSource;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1] = dt.Rows[i]["№"];
                        worksheet.Cells[i + 2, 2] = dt.Rows[i]["ФИО"];

                        // Записываем данные в столбцы с 1 по 31 и вычисляем сумму
                        int total = 0;
                        for (int j = 1; j <= 31; j++)
                        {
                            object value = dt.Rows[i][j.ToString()];
                            // Проверяем, является ли значение DBNull
                            if (value != DBNull.Value)
                            {
                                worksheet.Cells[i + 2, j + 2] = value;
                                // Выполняем преобразование к типу Int32 только для не-DBNull значений
                                total += Convert.ToInt32(value);
                            }
                            else
                            {
                                // Если значение DBNull, записываем 0 или другое значение по умолчанию
                                worksheet.Cells[i + 2, j + 2] = 0; // Или другое значение по умолчанию
                            }
                        }

                        // Записываем сумму в столбец "Всего"
                        worksheet.Cells[i + 2, 34] = total;

                        // Записываем данные для столбцов "Уваж." и "Неуваж."
                        worksheet.Cells[i + 2, 35] = dt.Rows[i]["Уваж."];
                        worksheet.Cells[i + 2, 36] = dt.Rows[i]["Неуваж."];
                        
                    }
                    //
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
