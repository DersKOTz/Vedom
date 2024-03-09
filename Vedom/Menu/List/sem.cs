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
using DocumentFormat.OpenXml.Spreadsheet;

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
                if ((Properties.Settings.Default.semsestSave == ""))
                {
                    comboBox1.SelectedIndex = 1;
                }
                else
                {
                    comboBox1.SelectedItem = Properties.Settings.Default.semsestSave;
                }
            }
        }
        private void LoadDataFromExcel()
        {
            dataGridView1.Visible = false;
            label1.Visible = true;
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

                    // Создаем списки для хранения значений
                    List<double> totalValues = new List<double>();
                    List<double> respectfulValues = new List<double>();
                    List<double> disrespectfulValues = new List<double>();

                    // Проходим по каждому месяцу
                    for (int month = 1; month <= 12; month++)
                    {
                        // Получаем значения из соответствующих ячеек
                        object cellValue = sheets[month - 1].Cells[i, 34].Value;
                        object cellValue2 = sheets[month - 1].Cells[i, 35].Value;
                        object cellValue3 = sheets[month - 1].Cells[i, 36].Value;

                        // Проверяем, что значения не равны null и являются double
                        if (cellValue != null && cellValue is double)
                        {
                            totalValues.Add((double)cellValue);
                        }
                        if (cellValue2 != null && cellValue2 is double)
                        {
                            respectfulValues.Add((double)cellValue2);
                        }
                        if (cellValue3 != null && cellValue3 is double)
                        {
                            disrespectfulValues.Add((double)cellValue3);
                        }
                    }

                    // Вычисляем суммы значений
                    double totalSum = totalValues.Sum();
                    double respectfulSum = respectfulValues.Sum();
                    double disrespectfulSum = disrespectfulValues.Sum();

                    // Присваиваем суммы соответствующим столбцам
                    row["Всего"] = totalSum;
                    row["Уваж."] = respectfulSum;
                    row["Неуваж."] = disrespectfulSum;



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
                dataGridView1.Visible = true;
                label1.Visible = false;
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

                int countExams = subjects.Count(s => s == "Экзамен");
                int countExams1 = subjects.Count(s => s == "Зачет");
                int countExams2 = subjects.Count(s => s == "Курсовик");
                int countExams3 = subjects.Count(s => s == "Практика");

                // экзамен
                for (int i = 2; i < countExams + 2; i++)
                {
                    worksheet.Cells[6, i + 1] = dataGridView.Columns[i].HeaderText;
                    // экзам
                    for (int ai = 0; ai < countExams; ai++)
                    {
                        worksheet.Cells[4, ai + 3] = subjects[ai]; // Используем индексы строк и столбцов, начиная с 1
                    }
                    // фио
                    for (int ai = 0; ai < countExams; ai++)
                    {
                        worksheet.Cells[5, ai + 3] = subjects2[ai]; // Используем индексы строк и столбцов, начиная с 1
                    }
                }

                // зачет
                for (int i = 2; i < countExams1 + 2; i++)
                {
                    worksheet.Cells[6, i + 6] = dataGridView.Columns[i + countExams].HeaderText;
                    // экзам
                    for (int ai = 0; ai < countExams1; ai++)
                    {
                        worksheet.Cells[4, ai + 8] = subjects[ai + countExams]; // Используем индексы строк и столбцов, начиная с 1
                    }
                    // фио
                    for (int ai = 0; ai < countExams1; ai++)
                    {
                        worksheet.Cells[5, ai + 8] = subjects2[ai + countExams]; // Используем индексы строк и столбцов, начиная с 1
                    }
                }

                // курсовик
                for (int i = 2; i < countExams2 + 2; i++)
                {
                    worksheet.Cells[6, i + 12] = dataGridView.Columns[i + countExams + countExams1].HeaderText;
                    // экзам
                    for (int ai = 0; ai < countExams2; ai++)
                    {
                        worksheet.Cells[4, ai + 14] = subjects[ai + countExams + countExams1]; // Используем индексы строк и столбцов, начиная с 1
                    }
                    // фио
                    for (int ai = 0; ai < countExams2; ai++)
                    {
                        worksheet.Cells[5, ai + 14] = subjects2[ai + countExams + countExams1]; // Используем индексы строк и столбцов, начиная с 1
                    }
                }

                // практика
                for (int i = 2; i < countExams3 + 2; i++)
                {
                    worksheet.Cells[6, i + 14] = dataGridView.Columns[i + countExams + countExams1 + countExams3].HeaderText;
                    // экзам
                    for (int ai = 0; ai < countExams3; ai++)
                    {
                        worksheet.Cells[4, ai + 16] = subjects[ai + countExams + countExams1 + countExams3]; // Используем индексы строк и столбцов, начиная с 1
                    }
                    // фио
                    for (int ai = 0; ai < countExams3; ai++)
                    {
                        worksheet.Cells[5, ai + 16] = subjects2[ai + countExams + countExams1 + countExams3]; // Используем индексы строк и столбцов, начиная с 1
                    }
                }

                // пропуски
                for (int i = dataGridView.Columns.Count - 3; i < dataGridView.Columns.Count; i++)
                {
                    worksheet.Cells[5, i + 10] = dataGridView.Columns[i].HeaderText;
                }



                // Запись данных
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    // запись часов
                    for (int j = dataGridView.Columns.Count - 3; j < dataGridView.Columns.Count; j++)
                    {
                        if (dataGridView.Rows[i].Cells[j].Value != null)
                        {
                            worksheet.Cells[i + 8, j + 10] = dataGridView.Rows[i].Cells[j].Value.ToString();
                        }
                        else
                        {
                            // Обработка случая, когда значение ячейки равно null
                            worksheet.Cells[i + 8, j + 10] = ""; // Или другое значение по умолчанию
                        }
                    }

                    // 1234 и фио
                    for (int j = 0; j < 2; j++)
                    {
                        worksheet.Cells[i + 8, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                    }

                    // экзамен
                    for (int j = 3; j <= countExams + 2; j++)
                    {
                        worksheet.Cells[i + 8, j] = dataGridView.Rows[i].Cells[j - 1].Value.ToString();
                    }

                    // зачет
                    for (int j = countExams + 3; j <= countExams + countExams1 + 2; j++)
                    {
                        worksheet.Cells[i + 8, 8] = dataGridView.Rows[i].Cells[j - 1].Value.ToString();
                        Console.WriteLine("а: " + j.ToString());
                    }

                    // курсач
                    for (int j = countExams + countExams1 + 3; j <= countExams + countExams1 + countExams2 + 2; j++)
                    {
                        worksheet.Cells[i + 8, 14] = dataGridView.Rows[i].Cells[j - 1].Value.ToString();
                        Console.WriteLine("а: " + j.ToString());
                    }

                    // практика
                    for (int j = countExams + countExams1 + countExams2 + 3; j <= countExams + countExams1 + countExams2 + countExams3 + 2; j++)
                    {
                        worksheet.Cells[i + 8, 16] = dataGridView.Rows[i].Cells[j - 1].Value.ToString();
                        Console.WriteLine("а: " + j.ToString());
                    }
                }



                /*
                // 11 32
                for (int i = 11; i <= 32; i++)
                {
                    Excel.Range cellB4 = worksheet.Cells[i, 2]; // 4 - номер строки, 2 - номер столбца (B)
                    object cellValue = cellB4.Value;

                    // Проверка, является ли ячейка B4 пустой
                    if (cellValue == null || string.IsNullOrEmpty(cellValue.ToString()))
                    {
                        // Ваш код, выполняемый, если ячейка B4 пуста
                        int lastColumn = dataGridView1.Columns.Count;
                        int startColumn = lastColumn - 2;
                        Excel.Range rangeToMerge2 = worksheet.Range[worksheet.Cells[i, startColumn], worksheet.Cells[i, lastColumn]];
                        rangeToMerge2.Value = "";
                        // Далее ваше действие с rangeToMerge2, если ячейка B4 пуста
                    }
                }
                */

                // хз
                worksheet.Cells[6, 2].Value = "Дисциплины";
                worksheet.Cells[5, 2].Value = "Ф.И.О. Преподавателя";
                worksheet.Cells[5, 2].ColumnWidth = 23;
                worksheet.Cells[5, 1].Value = "№";
                worksheet.Cells[6, 1].Value = "";

                // 1 2 3 и тд
                int columnCount = dataGridView.Columns.Count;
                Excel.Range range1 = worksheet.Range[worksheet.Cells[7, 1], worksheet.Cells[7, columnCount]];
                for (int i = 1; i <= 20; i++)
                {
                    range1.Cells[1, i].Value = i;
                }
                range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // всего часов
                Excel.Range rangeToMerge5 = worksheet.Range[worksheet.Cells[33, 1], worksheet.Cells[33, dataGridView.Columns.Count - 3]];
                rangeToMerge5.Merge();
                rangeToMerge5.Value = "Всего пропусков (час)";
                rangeToMerge5.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                // часы
                int v = 0;
                int y = 0;
                int n = 0;
                int lastColumn1 = dataGridView.Columns.Count;
                int startColumn1 = lastColumn1 - 2;
                for (int row = 8; row <= 32; row++)
                {
                    v += Convert.ToInt32(worksheet.Cells[row, startColumn1].Value);
                    y += Convert.ToInt32(worksheet.Cells[row, startColumn1 + 1].Value);
                    n += Convert.ToInt32(worksheet.Cells[row, startColumn1 + 2].Value);
                }
                worksheet.Cells[33, startColumn1].Value = v;
                worksheet.Cells[33, startColumn1 + 1].Value = y;
                worksheet.Cells[33, startColumn1 + 2].Value = n;





                // всего в группе
                Excel.Range rangeToMerge6 = worksheet.Range[worksheet.Cells[35, 1], worksheet.Cells[35, 3]];
                rangeToMerge6.Merge();
                rangeToMerge6.Value = "Всего в группе человек";
                Excel.Worksheet studentsSheet = workbook.Sheets["Студенты"];
                Excel.Range fioColumn = studentsSheet.Range["B2:B26"];

                // Получаем массив значений ячеек в столбце "ФИО"
                object[,] fioValues = fioColumn.Value;
                int fioCount = 0;
                for (int i = 1; i <= fioValues.GetLength(0); i++)
                {
                    if (fioValues[i, 1] != null && fioValues[i, 1] != DBNull.Value && !string.IsNullOrWhiteSpace(fioValues[i, 1].ToString()))
                    {
                        fioCount++;
                    }
                }
                worksheet.Cells[35, 4].Value = fioCount;


                // колво неусп
                Excel.Range rangeToMerge7 = worksheet.Range[worksheet.Cells[36, 1], worksheet.Cells[36, 3]];
                rangeToMerge7.Merge();
                rangeToMerge7.Value = "Количество неуспевающих";
                int[] kolvoArray = new int[25]; // Создаем массив для хранения значений kolvo для каждой строки
                Excel.Range neysp = worksheet.Range[worksheet.Cells[8, 2], worksheet.Cells[32, 2]]; // Используем строки с 5 по 29
                int rowIndex = 0; // Индекс текущей строки в массиве kolvoArray
                foreach (Excel.Range cell in neysp)
                {
                    if (cell.Value != null && cell.Value.ToString() != "")
                    {
                        int kolvo = 0; // Значение kolvo для текущей строки

                        Excel.Range innerRange = worksheet.Range[worksheet.Cells[cell.Row, 3], worksheet.Cells[cell.Row, dataGridView.Columns.Count - 3]];
                        foreach (Excel.Range innerCell in innerRange)
                        {
                            if (innerCell.Value == null || innerCell.Value.ToString() == "2")
                            {
                                kolvo = 1; // Если найдена двойка или ячейка пуста, устанавливаем kolvo в 1
                                break; // Прерываем цикл, так как условие уже выполнено
                            }
                        }
                        kolvoArray[rowIndex] = kolvo; // Сохраняем значение kolvo для текущей строки
                        rowIndex++; // Переходим к следующей строке
                    }
                }
                int totalKolvo = 0;
                foreach (int kolvoValue in kolvoArray)
                {
                    totalKolvo += kolvoValue;
                }
                worksheet.Cells[36, 4].Value = totalKolvo;

                Array.Clear(kolvoArray, 0, kolvoArray.Length);
                rowIndex = 0;


                // колво на 4 и 5
                Excel.Range rangeToMerge8 = worksheet.Range[worksheet.Cells[37, 1], worksheet.Cells[37, 3]];
                rangeToMerge8.Merge();
                rangeToMerge8.Value = "Количество успевающих на 4 и 5";

                int totalKolvo1 = 0; // Переменная для подсчета количества успевающих на оценках 4, 5 и "+"
                foreach (Excel.Range cell in neysp)
                {
                    if (cell.Value != null && cell.Value.ToString() != "")
                    {
                        int kolvo = 0; // Значение kolvo для текущей строки
                        Excel.Range innerRange = worksheet.Range[worksheet.Cells[cell.Row, 3], worksheet.Cells[cell.Row, dataGridView.Columns.Count - 3]];

                        bool containsEmptyValue = false; // Флаг для обнаружения пустых значений в диапазоне

                        // Проверяем, что внутренний диапазон не пустой и не содержит пустых значений
                        foreach (Excel.Range innerCell in innerRange)
                        {
                            if (innerCell.Value == null || string.IsNullOrWhiteSpace(innerCell.Value.ToString()))
                            {
                                containsEmptyValue = true;
                                break;
                            }
                        }

                        // Если в диапазоне есть пустые значения, пропускаем эту строку
                        if (containsEmptyValue)
                            continue;

                        bool containsOnlyFourFiveAndPlus = true;
                        foreach (Excel.Range innerCell in innerRange)
                        {
                            if (innerCell.Value != null && !string.IsNullOrWhiteSpace(innerCell.Value.ToString()))
                            {
                                string valueStr = innerCell.Value.ToString();
                                if (valueStr != "4" && valueStr != "5" && valueStr != "зач")
                                {
                                    containsOnlyFourFiveAndPlus = false;
                                    break;
                                }
                            }
                        }
                        if (containsOnlyFourFiveAndPlus)
                        {
                            kolvo = 1; // Если все значения в строке - только 4, 5 или "+", устанавливаем kolvo в 1
                        }

                        totalKolvo1 += kolvo; // Добавляем кол-во успевающих на оценках 4, 5 и "+" в общий счетчик
                    }
                }
                worksheet.Cells[37, 4].Value = totalKolvo1;


                // абс усп
                Excel.Range rangeToMerge9 = worksheet.Range[worksheet.Cells[38, 1], worksheet.Cells[38, 3]];
                rangeToMerge9.Merge();
                rangeToMerge9.Value = "Абсолютная успеваемость в %";
                object значение_ячейки_32_6 = worksheet.Cells[35, 4].Value;
                object значение_ячейки_33_6 = worksheet.Cells[36, 4].Value;
                if (значение_ячейки_32_6 != null && значение_ячейки_33_6 != null)
                {
                    float знач = Convert.ToSingle(значение_ячейки_32_6) - Convert.ToSingle(значение_ячейки_33_6);
                    worksheet.Cells[38, 4].Value = Math.Round(знач / Convert.ToSingle(значение_ячейки_32_6) * 100, 1);
                }


                // кач усп
                Excel.Range rangeToMerge10 = worksheet.Range[worksheet.Cells[39, 1], worksheet.Cells[39, 3]];
                rangeToMerge10.Merge();
                rangeToMerge10.Value = "Качественная успеваемость в %";

                object значение_ячейки_34_6 = worksheet.Cells[37, 4].Value;
                if (значение_ячейки_32_6 != null && значение_ячейки_34_6 != null)
                {
                    worksheet.Cells[39, 4].Value = Math.Round(Convert.ToSingle(значение_ячейки_34_6) / Convert.ToSingle(значение_ячейки_32_6) * 100, 1);
                }


                // поогулы на 1
                Excel.Range rangeToMerge11 = worksheet.Range[worksheet.Cells[40, 1], worksheet.Cells[40, 3]];
                rangeToMerge11.Merge();
                rangeToMerge11.Value = "Прогулы на 1 человека час";
                worksheet.Cells[40, 4].Value = Math.Round(Convert.ToDouble(worksheet.Cells[33, startColumn1 + 2].Value) / Convert.ToDouble(worksheet.Cells[35, 4].Value), 1);


                // предметы чето там
                Excel.Range range = worksheet.Range[worksheet.Cells[5, 3], worksheet.Cells[6, dataGridView.Columns.Count - 3]];
                range.Orientation = 90;
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.Columns.AutoFit();
                // range.ColumnWidth += 2;

                /*
                // объед пропусков
                Excel.Range range11 = worksheet.Range[worksheet.Cells[5, dataGridView.Columns.Count - 2], worksheet.Cells[6, dataGridView.Columns.Count - 2]];
                Excel.Range range12 = worksheet.Range[worksheet.Cells[5, dataGridView.Columns.Count - 1], worksheet.Cells[6, dataGridView.Columns.Count - 1]];
                Excel.Range range13 = worksheet.Range[worksheet.Cells[5, dataGridView.Columns.Count], worksheet.Cells[6, dataGridView.Columns.Count]];
                Excel.Range range14 = worksheet.Range[worksheet.Cells[4, dataGridView.Columns.Count - 2], worksheet.Cells[4, dataGridView.Columns.Count]];
                range11.Merge();
                range12.Merge();
                range13.Merge();
                range14.Merge();
                range14.Value = "Пропуски (час)";
                range11.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range11.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range12.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range12.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range13.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range13.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range14.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range14.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                */

                // предметы UI
                Excel.Range range15 = worksheet.Range[worksheet.Cells[6, 3], worksheet.Cells[6, dataGridView.Columns.Count - 3]];
                range15.WrapText = true;
                range15.RowHeight = 110;
                range15.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range15.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range15.Columns.AutoFit();
                //range15.ColumnWidth += 2;


                int endColumn = dataGridView.Columns.Count - 3; // Конечная колонка, где заканчивается поиск




                Excel.Range r1 = worksheet.Range[worksheet.Cells[5, 1], worksheet.Cells[6, 2]];
                r1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                r1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                Excel.Range rangeToMerge13 = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[1, 8]];
                rangeToMerge13.Merge();
                rangeToMerge13.Value = "СВОДНАЯ ВЕДОМОСТЬ УСПЕВАЕМОСТИ ОБУЧАЮЩИХСЯ " + Properties.Settings.Default.group;

                Excel.Range rangeToMerge14 = worksheet.Range[worksheet.Cells[2, 2], worksheet.Cells[2, 8]];
                rangeToMerge14.Merge();
                rangeToMerge14.Value = Properties.Settings.Default.kurs + " курс за " + "_" + Properties.Settings.Default.semsestSave + "_" + " семестр" + Properties.Settings.Default.years + " учебного года";

                //рапмки
                Excel.Range rangeRama = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[33, dataGridView.Columns.Count]];
                rangeRama.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeRama.Borders.Weight = Excel.XlBorderWeight.xlThin;

                Excel.Range rangeRama1 = worksheet.Range[worksheet.Cells[35, 1], worksheet.Cells[40, 4]];
                rangeRama1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeRama1.Borders.Weight = Excel.XlBorderWeight.xlThin;

                worksheet.Cells[42, 5].Value = "Заведующий " + Properties.Settings.Default.fak + " _________________";
                worksheet.Cells[43, 5].Value = "Классный руководитель ___________________";
                worksheet.Cells[44, 5].Value = "Староста ________________________________";

                worksheet.Columns[1].AutoFit();

                // ячейки
                Excel.Range Rang1 = worksheet.Range[worksheet.Cells[8, 3], worksheet.Cells[32, dataGridView.Columns.Count]];
                Rang1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // часы
                Excel.Range Rang2 = worksheet.Range[worksheet.Cells[33, dataGridView.Columns.Count - 2], worksheet.Cells[33, dataGridView.Columns.Count]];
                Rang2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // доп
                Excel.Range Rang3 = worksheet.Range[worksheet.Cells[35, 4], worksheet.Cells[40, 4]];
                Rang3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

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

                // Сохраняем изменения в файле
                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                MessageBox.Show("Данные сохранены в Excel файл!");
            }
        }

        private void print_Click(object sender, EventArgs e)
        {
            string fileName = "vedom.xlsx";
            ExportToExcel(dataGridView1, fileName);


            string excelFilePath = "vedom.xlsx";
            // Название листа
            string sheetName = "Ведомость семестр №" + Properties.Settings.Default.semsestSave;
            // Создание объекта приложения Excel
            Excel.Application excelApp = new Excel.Application();
            // Открытие книги Excel
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelFilePath);
            // Получение листа по имени
            Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[sheetName];
            int lastColumnIndex = dataGridView1.Columns.Count;
            // Если количество столбцов меньше 10, установите lastColumnIndex на 10
            if (lastColumnIndex < 10)
            {
                lastColumnIndex = 10;
            }
            // Формирование диапазона от A1 до последнего столбца
            Excel.Range excelRange = excelWorksheet.Range["A1", excelWorksheet.Cells[45, lastColumnIndex]];
            /*
            // Автоматическая подгонка размеров страницы по содержимому
            excelRange.Columns.AutoFit();
            excelRange.Rows.AutoFit();
            */
            // Вписать лист на одну страницу
            excelWorksheet.PageSetup.FitToPagesWide = 1;
            excelWorksheet.PageSetup.FitToPagesTall = 1;

            // Печать всего листа
            excelRange.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Закрытие книги Excel
            excelWorkbook.Close(false);
            // Закрытие приложения Excel
            excelApp.Quit();
        }
    }
}
