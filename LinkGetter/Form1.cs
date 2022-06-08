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
            openFileDialog1.InitialDirectory = @"C:\";//--"C:\\";
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
            //excel.writeToCell(0, 0, "radi");
            // excel.writeToCell(0, 1, "radi i ovaj");
            // MessageBox.Show(excel.readCell(0, 0));


            //string[] lista = { "banana", "papiga", "krumpir" };
            //List<string> lista = new List<string>();
            //string[] linksx = { };
            //List<string> linksx = new List<string>();
            //List<string> termsList = new List<string>();
            int counter = 0;

            while (true)
            {
                //MessageBox.Show(PathTextBox.Text);
               
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

                    //excel.writeALinkToCell(counter, 1, getLink(excel.readCell(counter, 0)));
                    excel.writeToCell(counter, 1, link);
                }
                catch (Exception ex)
                {
                    excel.writeToCell(counter, 1, "xxxxx");
                }

                //excel.writeToCell(counter, 1, getLink(excel.readCell(counter, 0)));
                //= HYPERLINK("https://en.wiktionary.org/wiki/naziv")
                // "= HYPERLINK((""https://www.google.ru/?q="" & b1)"
                counter++;
                
            }
            excel.save();
            excel.wbclose();
            /*foreach (String item in lista)
            {
                
                linksx.Add(getLink(item));
            }
            
            for (int i = 0; i < lista.Length; i++)
            {
                excel.writeToCell(i, 0, lista[i]);
                excel.writeToCell(i, 1, linksx[i]);
                Console.WriteLine(lista[i]);
                Console.WriteLine(linksx[i]);
            }*/

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
            //driver.Manage().Window.Maximize();

            //driver.SwitchTo().Window(driver.WindowHandles[1]);
            //IWebElement consent = driver.FindElement(By.XPath("//*[@id=\"L2AGLb\"]/div"));
            //consent.Click();
            IWebElement searchInput = driver.FindElement(By.Name("q"));
            searchInput.SendKeys(search);
            searchInput.SendKeys(OpenQA.Selenium.Keys.Enter);
            //IWebElement all_options = driver.FindElement(By.Name("option"));
            //div[@id='ires']//h3/a[1]/@href
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








