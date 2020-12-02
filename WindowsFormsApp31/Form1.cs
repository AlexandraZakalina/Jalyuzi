using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp31
{
    public partial class Form1 : Form
    {
        double sum = 0;
        int shir = 0;
        int dlin = 0;
        double metr = 0;
        string mater = "";
        private void Repwo(string subToReplace, string text, Word.Document worddoc)
        {
            var range = worddoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: subToReplace, ReplaceWith: text);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            shir = Convert.ToInt32(textBox1.Text);
            dlin = Convert.ToInt32(textBox2.Text);
            mater = comboBox1.Text;
            if(mater == "Пластик")
            {
                metr = 9.90;
            }
            else
            {
                metr = 15.50;
            }
            sum = metr * (shir * dlin/100);

            label3.Text += Convert.ToString(shir) + ".00 cm x " + Convert.ToString(dlin) + ".00 cm";
            label4.Text += comboBox1.Text;
            label5.Text += Convert.ToString(sum) + "p.";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Квитанция сохранена");// вывод сообщения о сохранении
            var WordApp = new Word.Application();
            WordApp.Visible = false;
            //путь к шаблону
            var Worddoc = WordApp.Documents.Open(Application.StartupPath + @"\Бланк-квитанция.docx");
            //заполнение
            Repwo("{date}", DateTime.Now.ToLongDateString(), Worddoc);
            Repwo("{itog}", sum.ToString(), Worddoc);
            Repwo("mater}", mater.ToString(), Worddoc);
            Repwo("shir}", shir.ToString(), Worddoc);
            Repwo("{dlin}", dlin.ToString(), Worddoc);
            //сохранение документа
            Worddoc.SaveAs2(Application.StartupPath + $"\\Квитанция на сумму {sum} от {DateTime.Now.ToLongDateString()}" + ".docx");
            //открываем документ
            WordApp.Visible = true;
        }
    }
}
