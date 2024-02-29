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
using System.Collections;
using System.Collections.Specialized;


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

            Properties.Settings.Default.semestr = new System.Collections.Specialized.StringCollection();
            foreach (var item in comboBox1.Items)
            {
                Properties.Settings.Default.semestr.Add(item.ToString());
            }

            Properties.Settings.Default.Save();
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






                if (worksheet != null)
                {
                    // Удаление содержимого
                    worksheet.Cells.ClearContents();
                    // Удаление форматирования
                    worksheet.Cells.ClearFormats();
                    // Получаем данные из DataGridView

                    worksheet.Cells[1, 1].Value = ("Семестр");
                    worksheet.Cells[1, 2].Value = ("Название");
                    worksheet.Cells[1, 3].Value = ("Тип оценивания");
                    worksheet.Cells[1, 4].Value = ("Преподаватель");

                    comboBox1.SelectedItem = "";

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            if (dataGridView1.Rows[i].Cells[j].Value != null)
                            {
                                worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            }
                            else
                            {
                                // Обработка случая, когда значение ячейки равно null
                                worksheet.Cells[i + 2, j + 1] = ""; // Или другое значение по умолчанию
                            }
                        }
                    }

                    Excel.Range usedRange = worksheet.UsedRange;
                    int lastRow = usedRange.Rows.Count;

                    // Временно добавляем колонку с числовым значением для порядка сортировки
                    Excel.Range tempColumn = worksheet.Cells[1, usedRange.Columns.Count + 1]; // Начинаем с последнего столбца + 1
                    tempColumn.Value = "0"; // Значение по умолчанию для всех строк
                    for (int i = 2; i <= lastRow; i++)
                    {
                        string value = worksheet.Cells[i, 3].Value.ToString(); // Значение из третьего столбца
                        switch (value)
                        {
                            case "Экзамен":
                                tempColumn.Cells[i, 1].Value = "1";
                                break;
                            case "Зачет":
                                tempColumn.Cells[i, 1].Value = "2";
                                break;
                            case "Курсовик":
                                tempColumn.Cells[i, 1].Value = "3";
                                break;
                            case "Практика":
                                tempColumn.Cells[i, 1].Value = "4";
                                break;
                        }
                    }
                    // Сортировка по столбцу 1 (A) по возрастанию, начиная со второй строки
                    Excel.Range sortRange = worksheet.Range["A2"].Resize[lastRow - 1]; // Игнорируем первую строку
                    sortRange.Sort(sortRange.Columns[1], Excel.XlSortOrder.xlAscending);
                    // Сортировка по временной колонке, начиная со второй строки
                    sortRange = worksheet.Range["A2", tempColumn.Cells[lastRow, 1]]; // Диапазон от A2 до временной колонки в последней строке
                    sortRange.Sort(sortRange.Columns[usedRange.Columns.Count + 1], Excel.XlSortOrder.xlAscending);
                    // Удаляем временную колонку
                    tempColumn.EntireColumn.Delete();


                }

                // Создаем коллекцию строк
                StringCollection semestr = Properties.Settings.Default.semestr;
                // Преобразуем коллекцию в List<string> для сортировки
                List<string> sortedSemestr = semestr.Cast<string>().ToList();
                // Сортируем список по возрастанию
                sortedSemestr.Sort();
                // Создаем новую коллекцию строк
                StringCollection sortedCollection = new StringCollection();
                // Добавляем отсортированные элементы обратно в StringCollection
                foreach (var item in sortedSemestr)
                {
                    sortedCollection.Add(item);
                }
                // Присваиваем отсортированную коллекцию переменной semestr
                Properties.Settings.Default.semestr = sortedCollection;
                Properties.Settings.Default.Save();


                workbook.Save();
                workbook.Close();
                excelApp.Quit();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            MessageBox.Show("Данные сохранены в Excel файл!");

        }



    }
}
