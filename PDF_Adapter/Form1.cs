using iTextSharp.text;
using iTextSharp.text.pdf;
using PDF_Adapter.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDF_Adapter
{
    public partial class frmHome : Form
    {
        public frmHome()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "PDF File|*.pdf", ValidateNames = true })
            {
                if(sfd.ShowDialog() == DialogResult.OK)
                {
                    // Генерируем документ
                    iTextSharp.text.Document doc = new iTextSharp.text.Document(PageSize.A4.Rotate());
                    
                    // Задаём шрифт
                    BaseFont baseFont = BaseFont.CreateFont("times.ttf", "cp1251", BaseFont.EMBEDDED);
                    iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);

                    try
                    {
                        // Создаём файл
                        PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                        doc.Open();

                        // Вносим текст с заданным шрифтом
                        doc.Add(new iTextSharp.text.Phrase(txtInfo.Text,font));
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        doc.Close();
                    }
                }
            }
        }

        private void btnSaveGrid_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "PDF File|*.pdf", ValidateNames = true })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    // Генерируем документ
                    iTextSharp.text.Document doc = new iTextSharp.text.Document(PageSize.A4.Rotate());
                    try
                    {
                        // Задаём шрифт
                        BaseFont baseFont = BaseFont.CreateFont("times.ttf", "cp1251", BaseFont.EMBEDDED);
                        iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);

                        // Создаём файл
                        PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                        doc.Open();

                        // Создаём таблицу
                        PdfPTable table = new PdfPTable(4);

                        // Создаём стилизованную ячейку (прокаченную :D)
                        PdfPCell cell = new PdfPCell(new Phrase("Любой текст", font));
                        cell.BackgroundColor = BaseColor.ORANGE;
                        cell.Colspan = dataGridView1.ColumnCount; // Объёдиняем указанное кол-во ячеек в одну
                        cell.HorizontalAlignment = 1; // 0-слева;1-по центру;2-справа

                        // Добавляём стилизованную ячейку в таблицу
                        table.AddCell(cell);

                        // Добавляём заголовки
                        for (int k = 0; k < dataGridView1.Columns.Count; k++)
                            table.AddCell(dataGridView1.Columns[k].HeaderText);
                        
                        // Добавляём данные в таблицу
                        for (int j = 0; j < dataGridView1.Rows.Count - 1; j++)
                        {
                            DataGridViewRow row = new DataGridViewRow();
                            row = dataGridView1.Rows[j];
                            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                                table.AddCell((row.Cells[i].Value != null) ? row.Cells[i].Value.ToString() : "");
                        }

                        // Вставляём таблицу в документ
                        doc.Add(table);

                        doc.Add(new iTextSharp.text.Phrase("GUCCI GANG :)", font));

                        // Добавляём изображение
                        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance("bmw.jpg");
                        img.ScalePercent(50.3f);
                        //img.SetAbsolutePosition(200, 50); // Точное положение на странице XY
                        img.Alignment = 1; // 0-слева;1-по центру;2-справа
                        doc.Add(img);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        doc.Close();
                    }
                }
            }
        }
    }
}
