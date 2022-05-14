using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
namespace Assignment
{
    class Program
    {
        static void Main(string[] args)
        {

            ChromeOptions options = new ChromeOptions();
            using (WebDriver driver = new ChromeDriver(options))
            {
                driver.Navigate().GoToUrl("https://srh.bankofchina.com/search/whpj/searchen.jsp");

                // Get elements

                // Get date boxes
                IWebElement start = driver.FindElement(By.Name("erectDate"));
                IWebElement end = driver.FindElement(By.Name("nothing"));
                start.Clear();
                end.Clear();

                // Get currency list

                IWebElement dropDown = driver.FindElement(By.Name("pjname"));
                SelectElement s = new SelectElement(dropDown);
                IList<IWebElement> els = s.Options;
                int count = els.Count;
                List<String> currencies = new List<string>();

                for (int i = 1; i < count; i++)
                    currencies.Add(els.ElementAt(i).Text);

                count--;
                // Get search button

                IWebElement search = driver.FindElement(By.XPath("//input[@value='search']"));

                // Set date
                DateTime dateTime = DateTime.UtcNow.Date;
                String currentDate = dateTime.ToString("yyyy/MM/dd");
                dateTime = DateTime.Now.AddDays(-2);
                String TwoDaysAgo = dateTime.ToString("yyyy/MM/dd");

                start.SendKeys(TwoDaysAgo);
                end.SendKeys(currentDate);

                // Loop over currencies
                
                for(int i = 0; i < count; i++)
                {
                    // Set currency
                    dropDown = driver.FindElement(By.Name("pjname"));
                    dropDown.SendKeys(currencies[i]);

                    // Click on search
                    search = driver.FindElement(By.XPath("//input[@value='search']"));
                    search.Click();

                    // Write to file

                    String c = driver.FindElement(By.XPath("/html/body/table[2]")).Text;

                    string fileName = currencies[i] + "_" + TwoDaysAgo + "_To_" + currentDate + ".txt";
                    using (StreamWriter writer = new StreamWriter(fileName))
                    {
                        writer.WriteLine(c);
                    }


                }
            }

            Console.WriteLine("Scraping finished.");
        }
    }
}
