using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;

using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace LinkGetter
{
    public partial class Form1 : Form
    {
        Uri linkxxx;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }                                                                                                                                              

        private void button1_Click(object sender, EventArgs e)//btnSelect
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Select File";
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Filter = "All files (*.*)|*.*|Excel File (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "")
            { PathTextBox.Text = openFileDialog1.FileName;
              
              linkxxx = new Uri(openFileDialog1.FileName, UriKind.Absolute);
            }
            else
            { PathTextBox.Text = "You didn't select the file!"; }
        }

        private void btnOpen_Click(object sender, EventArgs e)//btmExecute
        {
            Excel excel = new Excel(linkxxx, 1);
            
            int counter = 0;

            while (true)
            {
                
               
                if (excel.readCell(counter, 0) == null)
                {
                    
                    break;
                }
                if (excel.readCell(counter, 0) == "")
                {
                    
                    break;
                }
                try
                {
                    string link = "= HYPERLINK(\"" + getLink(excel.readCell(counter, 0)) + "\")";

                    
                    excel.writeToCell(counter, 1, link);
                }
                catch 
                {
                    excel.writeToCell(counter, 1, "xxxxx");
                }

                
                counter++;
                
            }
            excel.save();
            excel.wbclose();
            

        }

        private void BtnForward_Click(object sender, EventArgs e)
        {
            
        }
        public static string getLink(string search)
        {
            IWebDriver driver = new ChromeDriver();

            driver.Navigate().GoToUrl("http://www.google.com");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(driver1 => driver1.FindElement(By.Name("q")));
            IWebElement slazemSe = driver.FindElement(By.Id("L2AGLb"));
            slazemSe.Click();
            
            IWebElement searchInput = driver.FindElement(By.Name("q"));
            searchInput.SendKeys(search);
            searchInput.SendKeys(OpenQA.Selenium.Keys.Enter);
            
            try
            {

                wait.Until(driver1 => driver1.FindElement(By.ClassName("LC20lb")));
                IWebElement FirstLink = driver.FindElement(By.ClassName("LC20lb"));
                FirstLink.Click();

            }
            catch
            {



            }
            string xxurl;
            try
            {

                xxurl = driver.Url;

            }
            catch
            {

                xxurl = "";

            }

            driver.Quit();
            return xxurl;



        }

    }
}








