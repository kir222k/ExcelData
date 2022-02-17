/* Кирилл Уваров 2022г. 10 февраля. u.k.send@gmail.com. +79062644029
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace ExcelData.Class
{
    public class Prompt : IDisposable
    {
        private Form prompt { get; set; }
        public string Result { get; }

        public Prompt(string text, string caption)
        {
            List<string> listSheets = new List<string>
            {
                "Укажите имя листа книги Excel"
            };
            Result = ShowDialog(text, caption, listSheets);
        }

        public Prompt(string text, string caption, List<string> listSheets)
        {
            Result = ShowDialog(text, caption, listSheets);

        }

        //use a using statement
        private string ShowDialog(string text, string caption, List<string> listSheets)
        {
            prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen,
                TopMost = true
            };


            Label textLabel = new Label() { Left = 50, Top = 20, Text = text, Dock = DockStyle.Top, TextAlign = ContentAlignment.MiddleCenter };
            //TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 300 };
            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };

            //
            ComboBox cbx = new ComboBox()
            {
                Left = 50,
                Top = 30,
                Width=300,  
                Text = listSheets[0],
                DataSource = listSheets
            };

            confirmation.Click += (sender, e) => { prompt.Close(); };
            //prompt.Controls.Add(textBox);
            ///
            prompt.Controls.Add(cbx);

            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? cbx.Text : "";
        }



        public void Dispose()
        {
            //See Marcus comment
            if (prompt != null)
            {
                prompt.Dispose();
            }
        }
    }
}
