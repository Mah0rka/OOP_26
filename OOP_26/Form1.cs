using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace OOP_26
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            comboBox1.Items.Add(@"C:\Users\koval\source\repos\OOP_26-master\OOP_26\службова1.doc");
            comboBox1.Items.Add(@"C:\Users\koval\source\repos\OOP_26-master\OOP_26\службова2.doc");
        }
        Word.Application word = new Word.Application();
        Word.Document doc;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Object missingObj = System.Reflection.Missing.Value;
                Object templatePathObj = comboBox1.SelectedItem.ToString();

                doc = word.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                doc.Activate();

                foreach (Word.FormField f in doc.FormFields)
                {
                    switch (f.Name)
                    {
                        case "whomPosition":
                            f.Range.Text = textBox1.Text;
                            break;
                        case "institution":
                            f.Range.Text = textBox2.Text;
                            break;
                        case "whomName":
                            f.Range.Text = textBox3.Text;
                            break;
                        case "fromPosition":
                            f.Range.Text = textBox4.Text;
                            break;
                        case "fronName":
                            f.Range.Text = textBox5.Text;
                            break;
                    }
                }
                //Збереження по визначеному шляху
                Object savePath = @"D:\Збережений файл.doc";
                doc.SaveAs2(ref savePath);
                //Пошук 
                string findText = textBox11.Text;
                string replaceWith = textBox12.Text;
                bool found = false;

                foreach (Word.Range range in doc.StoryRanges)
                {
                    Word.Find find = range.Find;
                    find.Text = findText;
                    find.Replacement.Text = replaceWith;//заміна тектсу

                    if (find.Execute(Replace: WdReplace.wdReplaceAll))
                    {
                        found = true;
                    }
                }
                
                if (found)
                {
                    MessageBox.Show($"Текст '{findText}' було знайдено та змінено на '{replaceWith}'", "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Текст '{findText}' не було знайдено", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                word.Visible = true;


            }
            catch(Exception ex) 
            {
                if (doc != null)
                {
                    doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                    doc = null;
                }

                if (word != null)
                {
                    word.Quit();
                    word = null;
                }

                MessageBox.Show("Виникла помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (doc != null)
            {
                doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                doc = null;
            }

            if (word != null)
            {
                word.Quit();
                word = null;
            }

        }
    }
}
