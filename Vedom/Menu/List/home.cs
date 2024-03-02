using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Vedom.Menu.List
{
    public partial class home : Form
    {
        public home()
        {
            InitializeComponent();
        }

        private void save_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.group = group.Text;
            Properties.Settings.Default.kurs = kurs.Text;
            Properties.Settings.Default.fak = fak.Text;
            Properties.Settings.Default.years = years.Text;
            Properties.Settings.Default.Save();
        }

        private void home_Load(object sender, EventArgs e)
        {
            group.Text = Properties.Settings.Default.group;
            kurs.Text = Properties.Settings.Default.kurs;
            fak.Text = Properties.Settings.Default.fak;
            years.Text = Properties.Settings.Default.years;
        }
    }
}
