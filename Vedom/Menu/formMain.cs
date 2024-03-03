using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Vedom.Menu
{
    public partial class formMain : Form
    {
        public formMain()
        {
            InitializeComponent();
        }

        private void formMain_Load(object sender, EventArgs e)
        {
            home_Click(sender, e);
        }

        public void kill()
        {
            try
            {
                Process[] processes = Process.GetProcessesByName("excel");
                foreach (Process process in processes)
                {
                    // Получаем объект приложения Excel для каждого процесса
                    Excel.Application excelApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Excel.Application;
                    // Сохраняем и закрываем каждую книгу в приложении Excel
                    foreach (Excel.Workbook workbook in excelApp.Workbooks)
                    {
                        workbook.Save();
                        workbook.Close();
                    }
                    // Закрываем приложение Excel
                    excelApp.Quit();
                }
            }
            catch 
            {
                Application.Exit();
            }
        }

        private void close_Click(object sender, EventArgs e)
        {
            kill();
            Application.Exit();

        }

        public void restart(object sender, EventArgs e)
        {
            Application.Restart();
            Environment.Exit(0);
        }

        private void menuClose()
        {
            foreach (Control control in content.Controls)
            {
                if (control is Form form)
                {
                    form.Close();
                    break;
                }
            }
        }
        public void OpenForm<T>() where T : Form, new()
        {
            menuClose();
            T newForm = new T();
            newForm.TopLevel = false;
            content.Controls.Add(newForm);
            newForm.Show();
        }

        private void home_Click(object sender, EventArgs e)
        {
            OpenForm<Menu.List.home>();
        }

        private void student_Click(object sender, EventArgs e)
        {
            OpenForm<Menu.List.student>();
        }


        private void predmet_Click(object sender, EventArgs e)
        {
            OpenForm<Menu.List.predmet>();
        }


        private bool isDragging = false;
        private Point offset;
        private void formMain_MouseDown(object sender, MouseEventArgs e)
        {
            isDragging = true;
            offset = new Point(e.X, e.Y);
        }

        private void formMain_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point newLocation = this.Location;
                newLocation.X += e.X - offset.X;
                newLocation.Y += e.Y - offset.Y;
                this.Location = newLocation;
            }
        }

        private void formMain_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }

        private void close_MouseEnter(object sender, EventArgs e)
        {
            close.Image = Vedom.Properties.Resources.free_icon_clear_1632708;
            close.BackColor = Color.FromArgb(0x79, 0xb6, 0xc9);
        }

        private void close_MouseLeave(object sender, EventArgs e)
        {
            close.BackColor = Color.LightBlue;
            close.Image = Vedom.Properties.Resources.free_icon_delete_cross_63694;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string filePath = "vedom.xlsx";
            Process.Start(filePath);
        }

        private void propusk_Click_1(object sender, EventArgs e)
        {
            OpenForm<Menu.List.propusk>();
        }

        private void mec_Click_1(object sender, EventArgs e)
        {
            OpenForm<Menu.List.mec>();
        }

        private void sem_Click_1(object sender, EventArgs e)
        {
            OpenForm<Menu.List.sem>();
        }
    }
}
