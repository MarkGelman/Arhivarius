using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Arhivarius
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
                // Создаем объект приложения Excel
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                // Открываем книгу Excel
                Workbook workbook = excelApp.Workbooks.Open(@"Arhivarius\StudentsDB.xls");
                // Выбираем первый лист
                Worksheet worksheet = (Worksheet)workbook.Sheets[1];

                // Получаем используемый диапазон ячеек в листе
                Range range = worksheet.UsedRange;

            // Создаем DataTable для хранения данных
            System.Data.DataTable dt = new System.Data.DataTable();

                // Заполняем DataTable данными из Excel
                for (int i = 1; i <= range.Rows.Count; i++)
                {
                    DataRow row = dt.NewRow();
                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        if (i == 1)
                        {
                            // Если это первая строка, добавляем название столбца
                            dt.Columns.Add(range.Cells[i, j].Value.ToString());
                        }
                        else
                        {
                            // Заполняем данные из ячеек Excel в DataTable
                            row[j - 1] = range.Cells[i, j].Value;
                        }
                    }
                    if (i > 1)
                        dt.Rows.Add(row);
                }

                // Закрываем Excel
                workbook.Close(false);
                excelApp.Quit();

                // Отображаем данные в DataGridView
                dataGridView1.DataSource = dt;
            
        }
    }

}
    

