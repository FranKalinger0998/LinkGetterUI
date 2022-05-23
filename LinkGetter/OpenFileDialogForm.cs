using System;
using System.Drawing;
using System.IO;
using System.Security;
using System.Windows.Forms;

namespace LinkGetter
{
    internal class OpenFileDialogForm:Form
    {
        [STAThread]
        public static void FileDialogxx()
        {
            Application.SetCompatibleTextRenderingDefault(false);
            Application.EnableVisualStyles();
            Application.Run(new OpenFileDialogForm());
        }

        private Button selectButton;
    private OpenFileDialog openFileDialog1;
    private TextBox textBox1;

    public OpenFileDialogForm()
    {
        openFileDialog1 = new OpenFileDialog();
        selectButton = new Button
        {
            Size = new Size(100, 20),
            Location = new Point(15, 15),
            Text = "Select file"
        };
        selectButton.Click += new EventHandler(SelectButton_Click);
        textBox1 = new TextBox
        {
            Size = new Size(300, 300),
            Location = new Point(15, 40),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };
        ClientSize = new Size(330, 360);
        Controls.Add(selectButton);
        Controls.Add(textBox1);
    }
    private void SetText(string text)
    {
        textBox1.Text = text;
    }
    private void SelectButton_Click(object sender, EventArgs e)
    {
        if (openFileDialog1.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var sr = new StreamReader(openFileDialog1.FileName);
                SetText(sr.ReadToEnd());
            }
            catch (SecurityException ex)
            {
                MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }
        }
    }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // OpenFileDialogForm
            // 
            this.ClientSize = new System.Drawing.Size(278, 244);
            this.Name = "OpenFileDialogForm";
            this.Load += new System.EventHandler(this.OpenFileDialogForm_Load);
            this.ResumeLayout(false);

        }

        private void OpenFileDialogForm_Load(object sender, EventArgs e)
        {

        }
    }
}
