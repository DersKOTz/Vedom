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
using System.Globalization;

namespace Vedom.Menu.List
{
    public partial class sem : Form
    {
        public sem()
        {
            InitializeComponent();
        }

        private void sem_Load(object sender, EventArgs e)
        {
            LoadSemesterComboBoxItems();
            LoadDataFromExcel();
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
        private void LoadDataFromExcel()
        {
            string fileName = "vedom.xlsx";
            string studentsSheetName = "студенты";
            string vedomSheetName = "Ведомость семестр №" + Properties.Settings.Default.semsestSave;
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
                Excel.Worksheet disciplinesSheet = null;

                // Проверка наличия листа студентов
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

                // Проверка наличия листа ведомости
                Excel.Worksheet vedomSheet = null;
                if (!WorksheetExists(workbook, vedomSheetName))
                {
                    vedomSheet = workbook.Sheets.Add();
                    vedomSheet.Name = vedomSheetName;
                    workbook.Save();
                }
                else
                {
                    vedomSheet = workbook.Sheets[vedomSheetName];
                }

                // Получение текущего семестра из настроек                  
                int currentSemester = Convert.ToInt32(Properties.Settings.Default.semsestSave);

                DataTable dt = new DataTable();
                dt.Columns.Add("№");
                dt.Columns.Add("ФИО");

                disciplinesSheet = workbook.Sheets["Дисциплины"];

                // Добавление столбцов для каждого предмета с текущим семестром
                // Добавление столбцов для каждого предмета с текущим семестром
                for (int i = 2; i <= disciplinesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                {
                    string subjectName = disciplinesSheet.Cells[i, 2].Value?.ToString(); // Получаем название предмета из второго столбца
                    string secondLine = disciplinesSheet.Cells[i, 3].Value?.ToString(); // Получаем значение для второй строки из третьего столбца
                    string thirdLine = disciplinesSheet.Cells[i, 4].Value?.ToString(); // Получаем значение для третьей строки из четвёртого столбца

                    int semester = Convert.ToInt32(disciplinesSheet.Cells[i, 1].Value); // Получаем номер семестра
                    if (!string.IsNullOrEmpty(subjectName) && semester == currentSemester)
                    {
                        // Создаем двустрочный заголовок
                        string columnHeader = $"{secondLine}\n{thirdLine}\n{subjectName}";

                        // Добавляем столбец с двустрочным заголовком в таблицу данных только если семестр равен текущему
                        dt.Columns.Add(columnHeader);
                    }
                }


                dt.Columns.Add("Всего");
                dt.Columns.Add("Уваж.");
                dt.Columns.Add("Неуваж.");

                // Заполнение данных студентов
                for (int i = 2; i <= studentsSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                {
                    DataRow row = dt.NewRow();
                    row["№"] = studentsSheet.Cells[i, 1].Value;
                    row["ФИО"] = studentsSheet.Cells[i, 2].Value;

                    foreach (DataColumn column in dt.Columns)
                    {
                        string columnName = column.ColumnName;
                        int columnIndex = -1;

                        // Находим индекс столбца с названием columnName в листе Excel
                        for (int j = 1; j <= vedomSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column; j++)
                        {
                            if (vedomSheet.Cells[1, j].Value?.ToString() == columnName)
                            {
                                columnIndex = j;
                                break;
                            }
                        }

                        // Если столбец найден, записываем его значение в DataTable
                        if (columnIndex != -1 && vedomSheet.Cells[i, columnIndex].Value != null)
                        {
                            row[columnName] = vedomSheet.Cells[i, columnIndex].Value.ToString();
                        }

                    }

                    // row["Всего"] = attendanceSheet.Cells[i, 34].Value;
                    // row["Уваж."] = attendanceSheet.Cells[i, 35].Value;
                    // row["Неуваж."] = attendanceSheet.Cells[i, 36].Value;
                    dt.Rows.Add(row);
                }

                dataGridView1.DataSource = dt;
                workbook.Close();
                excelApp.Quit();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;

            if (Properties.Settings.Default.semsestSave != null)
            {
                comboBox1.SelectedItem = Properties.Settings.Default.semsestSave;
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
            LoadDataFromExcel();
        }

        private void save_Click(object sender, EventArgs e)
        {
            string fileName = "vedom.xlsx";
            ExportToExcel(dataGridView1, fileName);
        }



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

                // Формируем название листа по месяцу и году
                string monthYearSheetName = "Ведомость семестр №" + Properties.Settings.Default.semsestSave;

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
                    // Удаление содержимого
                    worksheet.Cells.ClearContents();

                    // Удаление форматирования
                    worksheet.Cells.ClearFormats();
                }

                // Если лист "прогулы" не существует, создаем его
                if (worksheet == null)
                {
                    worksheet = workbook.Sheets.Add();
                    worksheet.Name = studentsSheetName;
                    workbook.Save(); // Сохраняем изменения в файле
                }





                int startRow = 7;
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    string[] headerLines = dataGridView.Columns[i].HeaderText.Split('\n'); // Разбиваем двухстрочный заголовок на отдельные строки

                    for (int j = 0; j < headerLines.Length; j++)
                    {
                        worksheet.Cells[startRow + j - 3, i + 1] = headerLines[j]; // Записываем каждую строку заголовка в отдельную ячейку
                    }
                }


                // Запись данных
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        worksheet.Cells[i + startRow + 1, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                    }
                }



                // 1 2 3 и тд
                worksheet.Cells[6, 2].Value = "Дисциплины";
                int columnCount = dataGridView.Columns.Count;
                Excel.Range range1 = worksheet.Range[worksheet.Cells[7, 1], worksheet.Cells[7, columnCount]];
                for (int i = 1; i <= columnCount; i++)
                {
                    range1.Cells[1, i].Value = i;
                }


                // пропуски объед
                /*
                int lastColumn = dataGridView.Columns.Count;
                int startColumn = lastColumn - 2;
                Excel.Range rangeToMerge2 = worksheet.Range[worksheet.Cells[6, startColumn], worksheet.Cells[6, lastColumn]];
                rangeToMerge2.Merge();
                rangeToMerge2.Value = "Пропуски";
                rangeToMerge2.Columns.AutoFit();
                rangeToMerge2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                */



                // Сохраняем изменения в файле
                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                MessageBox.Show("Данные сохранены в Excel файл!");
            }
        }
    }
}
