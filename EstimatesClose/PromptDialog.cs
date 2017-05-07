using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EstimatesName
{
    public static class PromptDialog
    {
        public static string ShowDialog(string caption)
        {
            var prompt = new Form
            {
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Width = 350,
                Height = 120,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
            var textBox = new TextBox() { Left = 20, Top = 20, Width = 300 };
            var confirmation = new Button() { Text = @"Ok", Left = 220, Width = 100, Top = 50 };
            confirmation.Click += (sender, e) => prompt.Close();
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;
            prompt.ShowDialog();
            return textBox.Text;
        }
    }
}
