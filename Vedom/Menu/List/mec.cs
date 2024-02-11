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
                    // Получение номера текущего семестра из настроек                  
                    int currentSemester = Convert.ToInt32(Properties.Settings.Default.semsestSave);

                    DataTable dt = new DataTable();
                    dt.Columns.Add("№");
                    dt.Columns.Add("ФИО");

                    // Добавление столбцов для каждого предмета с текущим семестром
                    Excel.Worksheet disciplinesSheet = workbook.Sheets["Дисциплины"];
                    for (int i = 2; i <= disciplinesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        string subjectName = disciplinesSheet.Cells[i, 2].Value?.ToString(); // Получаем название предмета из листа "Дисциплины"
                        int semester = Convert.ToInt32(disciplinesSheet.Cells[i, 1].Value); // Получаем номер семестра
                        if (!string.IsNullOrEmpty(subjectName) && semester == currentSemester)
                            dt.Columns.Add(subjectName); // Добавляем столбец с названием предмета в таблицу данных только если семестр равен текущему
                    }

                    dt.Columns.Add("Всего");
                    dt.Columns.Add("Уваж.");
                    dt.Columns.Add("Неуваж.");

                    for (int i = 2; i <= studentsSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; i++)
                    {
                        DataRow row = dt.NewRow();
                        row["№"] = studentsSheet.Cells[i, 1].Value;
                        row["ФИО"] = studentsSheet.Cells[i, 2].Value;

                        // Загрузка данных о посещаемости для каждого предмета с текущим семестром
                        for (int j = 2; j <= disciplinesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; j++)
                        {
                            string subjectName = disciplinesSheet.Cells[j, 2].Value?.ToString();
                            int semester = Convert.ToInt32(disciplinesSheet.Cells[j, 1].Value);
                            if (!string.IsNullOrEmpty(subjectName) && semester == currentSemester)
                            {
                                // Предположим, что информация о посещаемости каждого студента
                                // содержится в столбцах, начиная со столбца с индексом 6
                                int columnIndex = j + 3; // 3 - учитываем первые три столбца (№, ФИО, Всего)
                                row[subjectName] = attendanceSheet.Cells[i, columnIndex].Value;
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime selectedDate = dateTimePicker1.Value;
            string selectedMonthYear = selectedDate.ToString("MMMM yyyy");
            LoadDataFromExcel(selectedMonthYear);
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

       

    }
}
