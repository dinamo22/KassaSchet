
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
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Security;

namespace KassaSchet
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        //проверка ввода, разрешены только 0-9, результат в последний столбец
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                try
                {
                    Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                }
                catch (Exception x)
                {
                    MessageBox.Show(x.Message);
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                }
            }

            int result = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    result += Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value);
                }
                for (int j = 3; j < dataGridView1.Rows[i].Cells.Count - 1; j++)
                {
                    result -= Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value);
                }
            }
            dataGridView1.Rows[0].Cells[5].Value = result;
        }
        //для ручной вставки
        private void dataGridView1_CellValueChanged()
        {
            try
            {
                int result = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        result += Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value);
                    }
                    for (int j = 3; j < dataGridView1.Rows[i].Cells.Count - 1; j++)
                    {
                        result -= Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value);
                    }
                }
                dataGridView1.Rows[0].Cells[5].Value = result;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        //выбор excel файла с последующим добавлением 5 элементов из него в новую строку
        private void ExportFromExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var excelApp = new Excel.Application();
            openFileDialog1 = new OpenFileDialog()
            {
                FileName = "Select a xlsx file",
                Filter = "xls files (*.xls*)|*.xls*",
                Title = "Open excel file"
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    excelApp.Workbooks.Open(openFileDialog1.FileName);
                    Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                    //excelApp.ActiveCell.CurrentRegion.Select();

                    int numberosomething = 4;
                    try
                    {
                        while (Convert.ToInt32(workSheet.Cells[numberosomething, "C"].Value) != 0 
                            || Convert.ToInt32(workSheet.Cells[numberosomething, "D"].Value) != 0
                            || Convert.ToInt32(workSheet.Cells[numberosomething, "F"].Value) != 0)
                        {
                            dataGridView1.Rows.Add(Convert.ToInt32(workSheet.Cells[numberosomething, "C"].Value),
                                                   0,
                                                   0,
                                                   Convert.ToInt32(workSheet.Cells[numberosomething, "D"].Value), 
                                                   Convert.ToInt32(workSheet.Cells[numberosomething, "F"].Value));
                            numberosomething++;
                        }
                        MessageBox.Show("Циферки перенесены успешно!");
                    }
                    catch(Exception exc)
                    {
                        MessageBox.Show(exc.Message);
                    }         
                    excelApp.Quit();
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                    try
                    {
                        excelApp.Quit();
                    }
                    catch (Exception xx)
                    {
                        MessageBox.Show(xx.Message);
                    }
                }
            }
            dataGridView1_CellValueChanged();
        }
        //app exit
        private void выходToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
