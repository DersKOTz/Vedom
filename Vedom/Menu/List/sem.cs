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
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Bibliography;

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
            dataGridView1.Visible = false;
            LoadSemesterComboBoxItems();
            LoadDataFromExcel();
            dataGridView1.Visible = true;
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
                if (Properties.Settings.Default.semsestSave != null)
                {
                    comboBox1.SelectedItem = Properties.Settings.Default.semsestSave;
                }
                else
                {
                    comboBox1.SelectedIndex = 0;
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
                List<string> subjects = new List<string>();
                for (int i = 2; i <= disciplinesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                {
                    string subjectName = disciplinesSheet.Cells[i, 2].Value?.ToString(); // Получаем название предмета из второго столбца
                    string secondLine = disciplinesSheet.Cells[i, 3].Value?.ToString(); // Получаем значение для второй строки из третьего столбца
                    string thirdLine = disciplinesSheet.Cells[i, 4].Value?.ToString(); // Получаем значение для третьей строки из четвёртого столбца
                    int semester = Convert.ToInt32(disciplinesSheet.Cells[i, 1].Value);
                    if (!string.IsNullOrEmpty(subjectName) && semester == currentSemester)
                    {
                        subjects.Add(subjectName);
                        dt.Columns.Add(subjectName); // Добавляем столбец для каждого предмета
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

                    foreach (string subject in subjects)
                    {
                        int columnIndex = FindColumnIndex(vedomSheet, subject);
                        if (columnIndex != -1)
                        {
                            row[subject] = vedomSheet.Cells[i + 6, columnIndex].Value; // !1
                        }
                    }


                    // Записать сумму в row["Всего"]
                    Excel.Worksheet[] sheets = new Excel.Worksheet[12];
                    for (int month = 1; month <= 12; month++)
                    {
                        string monthName = new DateTime(DateTime.Now.Year, month, 1).ToString("MMMM yyyy");
                        string sheetName = "Прогулы " + monthName + " " + Properties.Settings.Default.semsestSave;

                        Excel.Worksheet sheet;
                        try
                        {
                            // Попытка получить лист по имени
                            sheet = (Excel.Worksheet)workbook.Sheets[sheetName];
                        }
                        catch (System.Runtime.InteropServices.COMException ex)
                        {
                            // Если лист не существует, создаем новый лист
                            sheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, workbook.Sheets[workbook.Sheets.Count], Type.Missing, Type.Missing);
                            sheet.Name = sheetName;
                            // Дополнительные действия при создании нового листа, если необходимо
                        }

                        // Сохраняем лист в массив
                        sheets[month - 1] = sheet;
                    }

                    double[] valuesFromSheets = new double[12];
                    double[] valuesFromSheets2 = new double[12];
                    double[] valuesFromSheets3 = new double[12];
                    for (int month = 1; month <= 12; month++)
                    {
                        // Получить значение из соответствующей ячейки на листе для текущего месяца
                        object cellValue = sheets[month - 1].Cells[i, 34].Value;
                        object cellValue2 = sheets[month - 1].Cells[i, 35].Value;
                        object cellValue3 = sheets[month - 1].Cells[i, 36].Value;
                        // Console.WriteLine("Значение в ячейке " + month + ": " + (cellValue != null ? cellValue.ToString() : "пусто"));

                        // Проверить, что значение не равно null
                        if (cellValue != null && cellValue is double)
                        {
                            // Преобразовать значение в double и сохранить в массиве
                            valuesFromSheets[month - 1] = (double)cellValue;
                        }
                        else
                        {
                            // Обработка случая, когда значение ячейки не является double или равно null
                            // Например, можно использовать значение по умолчанию или выполнить другие действия
                            valuesFromSheets[month - 1] = 0; // значение по умолчанию
                        }
                        if (cellValue2 != null && cellValue2 is double)
                        {
                            // Преобразовать значение в double и сохранить в массиве
                            valuesFromSheets2[month - 1] = (double)cellValue2;
                        }
                        else
                        {
                            // Обработка случая, когда значение ячейки не является double или равно null
                            // Например, можно использовать значение по умолчанию или выполнить другие действия
                            valuesFromSheets2[month - 1] = 0; // значение по умолчанию
                        }
                        if (cellValue3 != null && cellValue3 is double)
                        {
                            // Преобразовать значение в double и сохранить в массиве
                            valuesFromSheets3[month - 1] = (double)cellValue3;
                        }
                        else
                        {
                            // Обработка случая, когда значение ячейки не является double или равно null
                            // Например, можно использовать значение по умолчанию или выполнить другие действия
                            valuesFromSheets3[month - 1] = 0; // значение по умолчанию
                        }
                    }

                    double sum = 0;
                    for (int month = 0; month < 12; month++)
                    {
                        sum += valuesFromSheets[month];
                    }
                    row["Всего"] = sum;

                    double sum2 = 0;
                    for (int month = 0; month < 12; month++)
                    {
                        sum2 += valuesFromSheets2[month];
                    }
                    row["Уваж."] = sum2;

                    double sum3 = 0;
                    for (int month = 0; month < 12; month++)
                    {
                        sum3 += valuesFromSheets3[month];
                    }
                    row["Неуваж."] = sum3;


                    dt.Rows.Add(row);
                }

                dataGridView1.DataSource = dt;
                // удаляем пустые листы
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    Excel.Range usedRange = sheet.UsedRange;
                    // Проверка на пустоту листа
                    if (usedRange.Rows.Count == 1 && usedRange.Columns.Count == 1 && string.IsNullOrEmpty(usedRange.Cells[1, 1].Value))
                    {
                        // Если лист пуст, удалить его
                        sheet.Delete();
                    }
                }


                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }



            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;


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

        private int FindColumnIndex(Excel.Worksheet sheet, string columnName)
        {
            if (sheet != null && sheet.Cells != null)
            {
                int lastColumn = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                for (int j = 1; j <= lastColumn; j++)
                {
                    if (sheet.Cells[6, j]?.Value?.ToString() == columnName)
                    {
                        return j;
                    }
                }
            }
            return -1; // Столбец с заданным заголовком не найден
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

                Excel.Worksheet disciplinesSheet = null;
                disciplinesSheet = workbook.Sheets["Дисциплины"];
                int currentSemester = Convert.ToInt32(Properties.Settings.Default.semsestSave);



                List<string> subjects = new List<string>();
                List<string> subjects2 = new List<string>();
                for (int i = 2; i <= disciplinesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                {
                    string subjectName = disciplinesSheet.Cells[i, 2].Value?.ToString(); // Получаем название предмета из второго столбца
                    string secondLine = disciplinesSheet.Cells[i, 3].Value?.ToString(); // Получаем значение для второй строки из третьего столбца
                    string thirdLine = disciplinesSheet.Cells[i, 4].Value?.ToString(); // Получаем значение для третьей строки из четвёртого столбца
                    int semester = Convert.ToInt32(disciplinesSheet.Cells[i, 1].Value);
                    if (!string.IsNullOrEmpty(subjectName) && semester == currentSemester)
                    {
                        subjects.Add(secondLine);
                        subjects2.Add(thirdLine);
                    }
                }



                int startRow = 7;
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    string[] headerLines = dataGridView.Columns[i].HeaderText.Split('\n'); // Разбиваем двухстрочный заголовок на отдельные строки

                    for (int j = 0; j < headerLines.Length; j++)
                    {
                        worksheet.Cells[startRow + j - 1, i + 1] = headerLines[j]; // Записываем каждую строку заголовка в отдельную ячейку
                    }
                }

                for (int i = 0; i < subjects.Count; i++)
                {
                    worksheet.Cells[4, i + 3] = subjects[i]; // Используем индексы строк и столбцов, начиная с 1
                }

                for (int i = 0; i < subjects2.Count; i++)
                {
                    worksheet.Cells[5, i + 3] = subjects2[i]; // Используем индексы строк и столбцов, начиная с 1
                }

                // Запись данных
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        worksheet.Cells[i + startRow + 1, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                    }
                }



                // хз
                worksheet.Cells[6, 2].Value = "Дисциплины";
                worksheet.Cells[5, 2].Value = "Ф.И.О. Преподавателя";
                worksheet.Cells[5, 2].ColumnWidth = 23;
                worksheet.Cells[5, 1].Value = "№";
                worksheet.Cells[6, 1].Value = "";
                // 1 2 3 и тд
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
