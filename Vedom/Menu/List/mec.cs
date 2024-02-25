﻿using System;
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
            if (Properties.Settings.Default.semsestSave != null)
            {
                comboBox1.SelectedItem = Properties.Settings.Default.semsestSave;
            }
            else
            {
                comboBox1.SelectedIndex = 0;
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
                    if (WorksheetExists(workbook, mecSheetName))
                    {
                        mecSheet = workbook.Sheets[mecSheetName];
                    }
                    else
                    {
                        mecSheet = workbook.Sheets.Add();
                        mecSheet.Name = mecSheetName;
                        workbook.Save();
                    }
                }
                else
                {
                    attendanceSheet = workbook.Sheets.Add();
                    attendanceSheet.Name = attendanceSheetName;

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

                if (studentsSheet != null && attendanceSheet != null && mecSheet != null)
                {
                    // Получение номера текущего семестра из настроек                  
                    int currentSemester = Convert.ToInt32(Properties.Settings.Default.semsestSave);

                    DataTable dt = new DataTable();
                    dt.Columns.Add("№");
                    dt.Columns.Add("ФИО");

                    // Получение списка предметов текущего семестра из листа "Дисциплины"
                    List<string> subjects = new List<string>();
                    Excel.Worksheet disciplinesSheet = workbook.Sheets["Дисциплины"];
                    for (int i = 2; i <= disciplinesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        string subjectName = disciplinesSheet.Cells[i, 2].Value?.ToString();
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

                    // Загрузка данных из листа "Ведомость"
                    for (int i = 2; i <= studentsSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        DataRow row = dt.NewRow();
                        row["№"] = studentsSheet.Cells[i, 1].Value;
                        row["ФИО"] = studentsSheet.Cells[i, 2].Value;

                        // Загрузка оценок для каждого предмета из листа "Ведомость"
                        foreach (string subject in subjects)
                        {
                            int columnIndex = FindColumnIndex(mecSheet, subject);
                            if (columnIndex != -1)
                            {
                                row[subject] = mecSheet.Cells[i + 3, columnIndex].Value; // !1
                            }
                        }

                        // Загрузка данных о прогулах из листа "Прогулы"
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

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;


        }

        private int FindColumnIndex(Excel.Worksheet sheet, string columnName)
        {
            if (sheet != null && sheet.Cells != null)
            {
                int lastColumn = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                for (int j = 1; j <= lastColumn; j++)
                {
                    if (sheet.Cells[4, j]?.Value?.ToString() == columnName)
                    {
                        return j;
                    }
                }
            }
            return -1; // Столбец с заданным заголовком не найден
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

                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    worksheet.Cells[4, i + 1] = dataGridView.Columns[i].HeaderText;
                }

                // назв и группа
                Excel.Range rangeNaz = worksheet.Range["A1:L1"];
                rangeNaz.Merge();
                rangeNaz.Value = "Ведомость аттестации и посещаемости студентов по группе NAZV" + " за " + selectedDate.ToString("MMMM yyyy", CultureInfo.CreateSpecificCulture("ru-RU"));

                // Объединяем ячейки предметов
                Excel.Range rangeToMerge = worksheet.Range[worksheet.Cells[3, 3], worksheet.Cells[3, dataGridView.Columns.Count - 3]];
                rangeToMerge.Merge();
                rangeToMerge.Value = "Успеваемость по дисциплинам";
                rangeToMerge.Font.Size = 10;
                rangeToMerge.Columns.AutoFit();
                rangeToMerge.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangeToMerge.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                // предметы чето там
                Excel.Range range = worksheet.Range[worksheet.Cells[4, 3], worksheet.Cells[4, dataGridView.Columns.Count - 3]];
                range.Orientation = 90;
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.Columns.AutoFit();
                range.ColumnWidth += 2;

                // Объед ячейки пропуски
                int lastColumn = dataGridView.Columns.Count;
                int startColumn = lastColumn - 2;
                Excel.Range rangeToMerge2 = worksheet.Range[worksheet.Cells[3, startColumn], worksheet.Cells[3, lastColumn]];
                rangeToMerge2.Merge();
                rangeToMerge2.Value = "Пропуски";
                rangeToMerge2.Columns.AutoFit();
                rangeToMerge2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // пропуски чето там
                Excel.Range rangeToMerge2_1 = worksheet.Range[worksheet.Cells[4, startColumn], worksheet.Cells[4, lastColumn]];
                rangeToMerge2_1.Orientation = 90;
                rangeToMerge2_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangeToMerge2_1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                rangeToMerge2_1.ColumnWidth = 5;
                // №
                Excel.Range rangeToMerge3 = worksheet.Range["A3:A4"];
                rangeToMerge3.Merge();
                rangeToMerge3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangeToMerge3.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                // фио
                Excel.Range rangeToMerge4 = worksheet.Range["B3:B4"];
                rangeToMerge4.Merge();
                rangeToMerge4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangeToMerge4.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                // всего часов
                Excel.Range rangeToMerge5 = worksheet.Range[worksheet.Cells[30, 1], worksheet.Cells[30, dataGridView.Columns.Count - 3]];
                rangeToMerge5.Merge();
                rangeToMerge5.Value = "Всего часов";
                rangeToMerge5.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                // всего в группе
                Excel.Range rangeToMerge6 = worksheet.Range[worksheet.Cells[32, 1], worksheet.Cells[32, 5]];
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
                worksheet.Cells[32, 6].Value = fioCount;

                // клас рус
                Excel.Range rangeToMerge12 = worksheet.Range[worksheet.Cells[40, 1], worksheet.Cells[40, 6]];
                rangeToMerge12.Merge();
                rangeToMerge12.Value = "Классный руководитель _________________";

                // староста
                Excel.Range rangeToMerge13 = worksheet.Range[worksheet.Cells[40, 7], worksheet.Cells[40, 11]];
                rangeToMerge13.Merge();
                rangeToMerge13.Value = "Староста _________________";

                // Запись данных
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        if (dataGridView.Rows[i].Cells[j].Value != null)
                        {
                            worksheet.Cells[i + 5, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                        }
                        else
                        {
                            // Обработка случая, когда значение ячейки равно null
                            worksheet.Cells[i + 5, j + 1] = ""; // Или другое значение по умолчанию
                        }
                    }
                }

                // колво неусп
                Excel.Range rangeToMerge7 = worksheet.Range[worksheet.Cells[33, 1], worksheet.Cells[33, 5]];
                rangeToMerge7.Merge();
                rangeToMerge7.Value = "Количество неуспевающих";
                int[] kolvoArray = new int[25]; // Создаем массив для хранения значений kolvo для каждой строки
                Excel.Range neysp = worksheet.Range[worksheet.Cells[5, 2], worksheet.Cells[29, 2]]; // Используем строки с 5 по 29
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
                worksheet.Cells[33, 6].Value = totalKolvo;

                Array.Clear(kolvoArray, 0, kolvoArray.Length);
                rowIndex = 0;


                // колво на 4 и 5
                Excel.Range rangeToMerge8 = worksheet.Range[worksheet.Cells[34, 1], worksheet.Cells[34, 5]];
                rangeToMerge8.Merge();
                rangeToMerge8.Value = "Количество успевающих на 4 и 5";
                foreach (Excel.Range cell in neysp)
                {
                    if (cell.Value != null && cell.Value.ToString() != "")
                    {
                        int kolvo = 0; // Значение kolvo для текущей строки
                        Excel.Range innerRange = worksheet.Range[worksheet.Cells[cell.Row, 3], worksheet.Cells[cell.Row, dataGridView.Columns.Count - 3]];
                        // Проверяем, что внутренний диапазон не пустой
                        bool innerRangeEmpty = true;
                        foreach (Excel.Range innerCell in innerRange)
                        {
                            if (innerCell.Value != null)
                            {
                                innerRangeEmpty = false;
                                break;
                            }
                        }
                        if (!innerRangeEmpty)
                        {
                            bool containsOnlyFourAndFive = true;
                            foreach (Excel.Range innerCell in innerRange)
                            {
                                if (innerCell.Value == null)
                                {
                                    // Если значение innerCell равно null, переходим к следующей итерации цикла
                                    continue;
                                }

                                if (innerCell.Value != null && innerCell.Value.ToString() != "")
                                {
                                    int value;
                                    if (!int.TryParse(innerCell.Value.ToString(), out value))
                                    {
                                        containsOnlyFourAndFive = false;
                                        break;
                                    }

                                    if (value != 4 && value != 5)
                                    {
                                        containsOnlyFourAndFive = false;
                                        break;
                                    }
                                }
                            }
                            if (containsOnlyFourAndFive)
                            {
                                kolvo = 1; // Если все значения в строке - только 4 или 5, устанавливаем kolvo в 1
                            }
                        }
                        kolvoArray[rowIndex] = kolvo; // Сохраняем значение kolvo для текущей строки
                        rowIndex++; // Переходим к следующей строке
                    }
                }
                // Складываем значения kolvo для каждой строки
                int totalKolvo1 = 0;
                foreach (int kolvoValue in kolvoArray)
                {
                    totalKolvo1 += kolvoValue;
                }
                worksheet.Cells[34, 6].Value = totalKolvo1;


                // абс усп
                Excel.Range rangeToMerge9 = worksheet.Range[worksheet.Cells[35, 1], worksheet.Cells[35, 5]];
                rangeToMerge9.Merge();
                rangeToMerge9.Value = "Абсолютная успеваемость в %";


                // кач усп
                Excel.Range rangeToMerge10 = worksheet.Range[worksheet.Cells[36, 1], worksheet.Cells[36, 5]];
                rangeToMerge10.Merge();
                rangeToMerge10.Value = "Качественная успеваемость в %";


                // поогулы на 1
                Excel.Range rangeToMerge11 = worksheet.Range[worksheet.Cells[37, 1], worksheet.Cells[37, 5]];
                rangeToMerge11.Merge();
                rangeToMerge11.Value = "Прогулы на 1 человека час";








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
