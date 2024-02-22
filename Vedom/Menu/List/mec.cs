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
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Vedom.Menu.List
{
    public partial class mec : Form
    {
        public mec()
        {
            InitializeComponent();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "MMMM yyyy";
            dateTimePicker1.ShowUpDown = true;
        }

        private void mec_Load(object sender, EventArgs e)
        {
            DateTime selectedDate = dateTimePicker1.Value;
            string selectedMonthYear = selectedDate.ToString("MMMM yyyy");
            LoadSemesterComboBoxItems();
            LoadDataFromExcel(selectedMonthYear);
        }


        private void LoadSemesterComboBoxItems()
        {
            if (Properties.Settings.Default.semestr != null)
            {
                foreach (string item in Properties.Settings.Default.semestr)
                {
                    if (!string.IsNullOrWhiteSpace(item)) // Проверяем, что строка не пустая или не состоит из пробелов
                    {
                        comboBox1.Items.Add(item);
                    }
                }
            }
        }


        private void LoadDataFromExcel(string selectedMonthYear)
        {
            string fileName = "vedom.xlsx";
            string studentsSheetName = "студенты";
            string attendanceSheetName = "Прогулы " + selectedMonthYear;
            string mecSheetName = "Ведомость " + selectedMonthYear;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

                if (!File.Exists(filePath))
                {
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
                Excel.Worksheet mecSheet = null;

                bool selectedMonthYearExists = WorksheetExists(workbook, attendanceSheetName);

                if (selectedMonthYearExists)
                {
                    attendanceSheet = workbook.Sheets[attendanceSheetName];
                    mecSheet = workbook.Sheets[mecSheetName];
                }
                else
                {
                    mecSheet = workbook.Sheets.Add();
                    mecSheet.Name = mecSheetName;
                    workbook.Save();
                }

                if (!WorksheetExists(workbook, studentsSheetName))
                {
                    studentsSheet = workbook.Sheets.Add();
                    studentsSheet.Name = studentsSheetName;
                    workbook.Save();
                }
                else
                {
                    studentsSheet = workbook.Sheets[studentsSheetName];
                }

                if (studentsSheet != null && attendanceSheet != null)
                {
                    int currentSemester = Convert.ToInt32(Properties.Settings.Default.semsestSave);

                    DataTable dt = new DataTable();
                    dt.Columns.Add("№");
                    dt.Columns.Add("ФИО");

                    Excel.Worksheet disciplinesSheet = workbook.Sheets["Дисциплины"];
                    for (int i = 2; i <= disciplinesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        string subjectName = disciplinesSheet.Cells[i, 2].Value?.ToString();
                        int semester = Convert.ToInt32(disciplinesSheet.Cells[i, 1].Value);
                        if (!string.IsNullOrEmpty(subjectName) && semester == currentSemester)
                            dt.Columns.Add(subjectName);
                    }

                    dt.Columns.Add("Всего");
                    dt.Columns.Add("Уваж.");
                    dt.Columns.Add("Неуваж.");

                    for (int i = 2; i <= studentsSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        DataRow row = dt.NewRow();
                        row["№"] = studentsSheet.Cells[i, 1].Value;
                        row["ФИО"] = studentsSheet.Cells[i, 2].Value;

                        foreach (DataColumn column in dt.Columns)
                        {
                            string columnName = column.ColumnName;
                            int columnIndex = -1;

                            if (mecSheet != null && mecSheet.Cells != null)
                            {
                                int lastColumn = mecSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                                for (int j = 1; j <= lastColumn; j++)
                                {
                                    if (mecSheet.Cells[1, j]?.Value?.ToString() == columnName)
                                    {
                                        columnIndex = j;
                                        break;
                                    }
                                }
                            }

                            if (columnIndex != -1 && mecSheet.Cells[i, columnIndex].Value != null)
                            {
                                row[columnName] = mecSheet.Cells[i, columnIndex].Value.ToString();
                            }
                        }

                        row["Всего"] = attendanceSheet.Cells[i, 34].Value;
                        row["Уваж."] = attendanceSheet.Cells[i, 35].Value;
                        row["Неуваж."] = attendanceSheet.Cells[i, 36].Value;
                        dt.Rows.Add(row);
                    }

                    dataGridView1.DataSource = dt;
                }

                workbook.Close();
                excelApp.Quit();
            }
        }

        private bool WorksheetExists(Excel.Workbook workbook, string worksheetName)
        {
            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == worksheetName)
                {
                    return true;
                }
            }
            return false;
        }

        private void ClearDataGridView()
        {
            // Установка источника данных в null перед очисткой
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            Properties.Settings.Default.semsestSave = comboBox1.SelectedItem.ToString();
            Properties.Settings.Default.Save();
            ClearDataGridView();
            DateTime selectedDate = dateTimePicker1.Value;
            string selectedMonthYear = selectedDate.ToString("MMMM yyyy");
            LoadDataFromExcel(selectedMonthYear);
        }

        private void save_Click(object sender, EventArgs e)
        {
            string fileName = "vedom.xlsx";
            ExportToExcel(dataGridView1, fileName);
        } // 

        private void ExportToExcel(DataGridView dataGridView, string fileName)
        {
            string studentsSheetName = "студенты"; // исправлено
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

                if (!File.Exists(filePath))
                {
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

                // Получаем выбранную дату из DateTimePicker
                DateTime selectedDate = dateTimePicker1.Value;

                // Формируем название листа по месяцу и году
                string monthYearSheetName = "Ведомость " + selectedDate.ToString("MMMM yyyy", CultureInfo.CreateSpecificCulture("ru-RU"));

                // Проверяем существует ли лист для текущего месяца и года
                bool monthYearSheetExists = false;
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == monthYearSheetName)
                    {
                        worksheet = sheet;
                        monthYearSheetExists = true;
                        break;
                    }
                }

                if (!monthYearSheetExists)
                {
                    worksheet = workbook.Sheets.Add();
                    worksheet.Name = monthYearSheetName;
                    workbook.Save(); // Сохраняем изменения в файле
                }
                else
                {
                    // Если лист для текущего месяца и года существует, устанавливаем его в качестве worksheet
                    worksheet = workbook.Sheets[monthYearSheetName];
                }

                // Если лист "прогулы" не существует, создаем его
                if (worksheet == null)
                {
                    worksheet = workbook.Sheets.Add();
                    worksheet.Name = studentsSheetName;
                    workbook.Save(); // Сохраняем изменения в файле
                }

                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView.Columns[i].HeaderText;
                }

                // Запись данных
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                    }
                }


                // Сохраняем изменения в файле
                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                MessageBox.Show("Данные сохранены в Excel файл!");
            }
        }

        private void dateTimePicker1_Leave(object sender, EventArgs e)
        {
            DateTime selectedDate = dateTimePicker1.Value;
            string selectedMonthYear = selectedDate.ToString("MMMM yyyy");
            LoadDataFromExcel(selectedMonthYear);
        }
    }
}
