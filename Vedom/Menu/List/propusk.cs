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
using System.Globalization;
using System.Threading;

namespace Vedom.Menu.List
{
    public partial class propusk : Form
    {
        public propusk()
        {
            InitializeComponent();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "MMMM yyyy";
            dateTimePicker1.ShowUpDown = true;
        }

        private void propusk_Load(object sender, EventArgs e)
        {
            DateTime selectedDate = dateTimePicker1.Value;
            string selectedMonthYear = selectedDate.ToString("MMMM yyyy");
            LoadDataFromExcel(selectedMonthYear);
        }

        private void LoadDataFromExcel(string selectedMonthYear)
        {
            dataGridView1.Visible = false;
            label1.Visible = true;
            string fileName = "vedom.xlsx";
            string studentsSheetName = "студенты";
            string attendanceSheetName = "Прогулы " + selectedMonthYear + " " + Properties.Settings.Default.semsestSave;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
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

                    bool selectedMonthYearExists = WorksheetExists(workbook, attendanceSheetName);

                    if (selectedMonthYearExists)
                    {
                        attendanceSheet = workbook.Sheets[attendanceSheetName];
                    }
                    else
                    {
                        attendanceSheet = workbook.Sheets.Add();
                        attendanceSheet.Name = attendanceSheetName;
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

                        for (int i = 2; i <= studentsSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                        {
                            DataRow row = dt.NewRow();
                            row["№"] = studentsSheet.Cells[i, 1].Value;
                            row["ФИО"] = studentsSheet.Cells[i, 2].Value;
                            dt.Rows.Add(row);
                        }

                        for (int i = 2; i <= attendanceSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                        {
                            if (i - 2 >= 0 && i - 2 < dt.Rows.Count)
                            {
                                DataRow row = dt.Rows[i - 2];
                                for (int j = 1; j <= 31; j++)
                                {
                                    if (attendanceSheet.Cells[i, j + 2].Value != null)
                                    {
                                        row[j.ToString()] = attendanceSheet.Cells[i, j + 2].Value;
                                    }
                                }

                                if (i <= attendanceSheet.Cells.Rows.Count && i <= attendanceSheet.Cells.Columns.Count)
                                {
                                    row["Всего"] = attendanceSheet.Cells[i, 34].Value;
                                    row["Уваж."] = attendanceSheet.Cells[i, 35].Value;
                                    row["Неуваж."] = attendanceSheet.Cells[i, 36].Value;
                                }
                            }
                        }

                        dataGridView1.DataSource = dt;
                    }

                }



                dataGridView1.Columns[0].Width = 30;
                for (int j = 2; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    dataGridView1.Columns[j].Width = 60;
                }

                dataGridView1.Columns[0].ReadOnly = true;
                dataGridView1.Columns[1].ReadOnly = true;
                dataGridView1.AllowUserToAddRows = false;
                dataGridView1.Visible = true;
                label1.Visible = false;

                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии файла: {ex.Message}\n попробуйте еще раз!");
                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

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

        int Printer = 0;
        private void save_Click(object sender, EventArgs e)
        {
            string fileName = "vedom.xlsx";
            string studentsSheetName = "прогулы" + " " + Properties.Settings.Default.semsestSave; // объявляем и инициализируем переменную studentsSheetName здесь
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;
            try
            {

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
                    // Здесь идет вызов метода LoadOrCreateAttendanceSheet(), где передаем оба требуемых аргумента               
                    Excel.Worksheet worksheet = null;

                    // Получаем выбранную дату из DateTimePicker
                    DateTime selectedDate = dateTimePicker1.Value;

                    // Формируем название листа по месяцу и году
                    string monthYearSheetName = "Прогулы " + selectedDate.ToString("MMMM yyyy", CultureInfo.CreateSpecificCulture("ru-RU")) + " " + Properties.Settings.Default.semsestSave;

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

                        // Проходимся по каждой строке в таблице dt
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            // Получаем значения "№" и "ФИО" из текущей строки и записываем их в соответствующие ячейки в Excel
                            worksheet.Cells[i + 2, 1] = dt.Rows[i]["№"];
                            worksheet.Cells[i + 2, 2] = dt.Rows[i]["ФИО"];

                            // Проверяем, если значение "ФИО" не пустое, иначе переходим к следующей строке
                            if (!string.IsNullOrEmpty(dt.Rows[i]["ФИО"].ToString()))
                            {
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
                                        worksheet.Cells[i + 2, j + 2] = null; // Или другое значение по умолчанию
                                    }
                                }

                                // Записываем сумму в столбец "Всего"
                                worksheet.Cells[i + 2, 34] = total;

                                // Проверяем, содержит ли ячейка значение DBNull
                                if (dt.Rows[i]["Уваж."] != DBNull.Value)
                                {
                                    worksheet.Cells[i + 2, 35] = dt.Rows[i]["Уваж."]; // Записываем данные для столбца "Уваж."
                                    int uvazhValue = Convert.ToInt32(dt.Rows[i]["Уваж."]); // Приводим значение из столбца "Уваж." к типу int
                                    int neuvazhValue = total - uvazhValue; // Вычисляем значение для столбца "Неуваж."
                                    worksheet.Cells[i + 2, 36] = neuvazhValue; // Записываем значение в столбец "Неуваж."
                                }
                                else
                                {
                                    // Если значение DBNull, записываем 0 в ячейку столбца "Уваж." и "Неуваж."
                                    worksheet.Cells[i + 2, 35] = 0;
                                    worksheet.Cells[i + 2, 36] = 0;
                                }
                            }
                        }

                    }

                    worksheet.Cells[28, 1].Value = $"Ведомость посещаемости студентов по группе {Properties.Settings.Default.group} за {selectedDate.ToString("MMMM yyyy", CultureInfo.CreateSpecificCulture("ru-RU"))}";
                    worksheet.Cells[28, 1].Font.Size = 14;

                    Excel.Range rangeRama = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[26, 36]];
                    rangeRama.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeRama.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    rangeRama.RowHeight = 19;

                    Excel.Range rangeRama1 = worksheet.Range[worksheet.Cells[1, 3], worksheet.Cells[26, 33]];
                    rangeRama1.ColumnWidth = 4;
                    rangeRama1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Cells[1, 1].ColumnWidth = 4;
                    worksheet.Columns[2].AutoFit();
                    rangeRama1.ColumnWidth = 4;

                    Excel.Range rangeRama3 = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 36]];
                    rangeRama3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                    Excel.Range rangeRama2 = worksheet.Range[worksheet.Cells[33, 3], worksheet.Cells[26, 36]];
                    rangeRama2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    // печать и тд
                    Excel.Range rangePrint = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[28, 36]];
                    worksheet.PageSetup.PrintArea = rangePrint.Address;
                    worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                    worksheet.PageSetup.LeftMargin = excelApp.InchesToPoints(0.5);
                    worksheet.PageSetup.RightMargin = excelApp.InchesToPoints(0.5);
                    worksheet.PageSetup.TopMargin = excelApp.InchesToPoints(0.5);
                    worksheet.PageSetup.BottomMargin = excelApp.InchesToPoints(0.5);

                    // Установить масштаб для вписывания листа на одну страницу
                    worksheet.PageSetup.Zoom = false;
                    worksheet.PageSetup.FitToPagesWide = 1;
                    worksheet.PageSetup.FitToPagesTall = 1;


                    if (Printer == 1)
                    {
                        worksheet.PrintOutEx();
                    }


                    workbook.Save();
                    workbook.Close();
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    MessageBox.Show("Данные сохранены в Excel файл!");
                }
            }
            catch (Exception ex)
            {
                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                MessageBox.Show($"Ошибка при сохранение файла: {ex.Message}\n попробуйте еще раз!");
            }
        }

        private void dateTimePicker1_Leave(object sender, EventArgs e)
        {
            DateTime selectedDate = dateTimePicker1.Value;
            string selectedMonthYear = selectedDate.ToString("MMMM yyyy");
            LoadDataFromExcel(selectedMonthYear);
        }

        private void print_Click(object sender, EventArgs e)
        {
            DateTime selectedDate = dateTimePicker1.Value;
            string selectedMonthYear = selectedDate.ToString("MMMM yyyy");
            Printer = 1;
            save_Click(sender, e);
            Thread.Sleep(2000);
            Printer = 0;
        }
    }
}
