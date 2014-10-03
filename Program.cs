using ConsoleApplication4;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace UniverParser
{
    static class Program
    {
        static readonly string rootPage = "http://edu.ru";
        static readonly string searchPage = "http://edu.ru/abitur/act.1/index.php";
        static readonly string listPage = "http://edu.ru/abitur/act.2/index.php";
        internal static int rowCounter = 3;

        static void Main(string[] args)
        {
            var towns = GetTowns(searchPage);
            foreach (var town in towns)
            {
                var htmlDocument = GetHtmlDocument(listPage + "?town_=" + 45000000 + "&show_results=300");
                var univerLinkCollection = GetUniverLinkCollection(htmlDocument);
                var univerCollection = new List<Univer>();
                foreach (var univerLink in univerLinkCollection)
                {
                    var univer = GetGeneralInfo(univerLink);
                    univer.Management = GetManagement(univerLink);
                    univerCollection.Add(univer);
                }

                FillExcel(town.Value, univerCollection);
            }
        }

        private static Univer GetGeneralInfo(string univerLink)
        {
            var htmlDocument = GetHtmlDocument(univerLink);

            var tdPart = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/td[2]") != null ? "td[2]" : "td";
            var univer = new Univer();
            if (htmlDocument.DocumentNode.SelectSingleNode("//h1[@class='cart']") != null)
            {
                univer.Name = htmlDocument.DocumentNode.SelectSingleNode("//h1[@class='cart']").InnerHtml.Clean();
            }

            var site = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/a");
            if (site != null && (site.InnerHtml.Contains("www") || site.InnerHtml.Contains("http")))
            {
                univer.Site = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/a").InnerHtml.Clean();
            }

            if (htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/" + tdPart + "/b") != null)
            {
                univer.Form = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/" + tdPart + "/b").InnerHtml.Clean();
            }

            if (htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/" + tdPart + "/text()[6]") != null)
            {
                univer.Address = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/" + tdPart + "/text()[6]").InnerHtml.Clean();
            }

            if (htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/" + tdPart + "/text()[13]") != null)
            {
                univer.Telephone = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/" + tdPart + "/text()[13]").InnerHtml.Clean();
            }

            if (htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/" + tdPart + "/a[2]") != null)
            {
                univer.Email = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table[3]/tr/" + tdPart + "/a[2]").InnerHtml.Clean();
            }

            if (htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/div") != null)
            {
                var div = htmlDocument.DocumentNode.SelectNodes("//td[@class='tdcont']/div").Count > 4 ? "[2]" : string.Empty;
                var fullString = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/div" + div).InnerHtml.Clean();
                DateTime date;
                var culture = CultureInfo.CreateSpecificCulture("ru-RU");
                var styles = DateTimeStyles.None;
                var datePart = fullString.Split(new string[] { ":" }, StringSplitOptions.None)[1];
                if (DateTime.TryParse(datePart, culture, styles, out date))
                {
                    univer.LastModified = date;
                }
            }

            if (htmlDocument.DocumentNode.SelectNodes("//td[@class='tdcont']/a").Where(x => x.Attributes["href"].Value.Contains("/abitur/act.3/")).Count() > 0)
            {
                var link = htmlDocument.DocumentNode.SelectNodes("//td[@class='tdcont']/a").First(x => x.Attributes["href"].Value.Contains("act.3/"));
                univer.Link = new Dictionary<string, string>() { { link.InnerText, rootPage + link.Attributes["href"].Value } };
            }

            return univer;
        }

        private static Management GetManagement(string univerLink)
        {
            var htmlDocument = GetHtmlDocument(univerLink.Replace("ds.1", "ds.2"));
            var management = new Management();
            if (htmlDocument.DocumentNode.SelectSingleNode("//table[@class='t2']/tr[2]/td[1]") != null)
            {
                management.Position = htmlDocument.DocumentNode.SelectSingleNode("//table[@class='t2']/tr[2]/td[1]").InnerHtml.Clean();
            }

            if (htmlDocument.DocumentNode.SelectSingleNode("//table[@class='t2']/tr[2]/td[2]") != null)
            {
                management.FIO = htmlDocument.DocumentNode.SelectSingleNode("//table[@class='t2']/tr[2]/td[2]").InnerHtml.Clean();
            }

            if (htmlDocument.DocumentNode.SelectSingleNode("//table[@class='t2']/tr[2]/td[3]") != null)
            {
                management.Phone = htmlDocument.DocumentNode.SelectSingleNode("//table[@class='t2']/tr[2]/td[3]").InnerHtml.Clean();
            }

            if (htmlDocument.DocumentNode.SelectSingleNode("//table[@class='t2']/tr[2]/td[4]") != null)
            {
                management.Email = htmlDocument.DocumentNode.SelectSingleNode("//table[@class='t2']/tr[2]/td[4]").InnerHtml.Clean();
            }

            return management;
        }

        static void FillExcel(string townName, IList<Univer> univerCollection)
        {
            var excelApp = new Excel.Application { Visible = false };
            var excelBook = excelApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + @"..\..\result.xlsx");
            var excelSheet = (Excel.Worksheet)excelBook.Sheets[1];
            foreach (var univer in univerCollection)
            {
                excelSheet.Cells[rowCounter, 1] = univer.Name;
                excelSheet.Cells[rowCounter, 2] = univer.Site;
                excelSheet.Cells[rowCounter, 3] = univer.Form;
                excelSheet.Cells[rowCounter, 4] = univer.Address;
                excelSheet.Cells[rowCounter, 5] = univer.Telephone;
                excelSheet.Cells[rowCounter, 6] = univer.Email;
                excelSheet.Cells[rowCounter, 7] = univer.Management.Position;
                excelSheet.Cells[rowCounter, 8] = univer.Management.FIO;
                excelSheet.Cells[rowCounter, 9] = univer.Management.Phone;
                excelSheet.Cells[rowCounter, 10] = univer.Management.Email;
                excelSheet.Cells[rowCounter, 11] = townName;
                excelSheet.Cells[rowCounter, 12] = univer.LastModified;
                if (univer.Link != null)
                    excelSheet.Hyperlinks.Add((Excel.Range)excelSheet.Cells[rowCounter, 13], univer.Link.First().Value, Type.Missing, univer.Link.First().Key, univer.Link.First().Key);
                rowCounter++;
            }

            excelBook.Save();
            excelBook.Close();
            excelApp.Quit();
        }

        static ICollection<string> GetUniverLinkCollection(HtmlDocument htmlDocument)
        {
            var univerLinkCollection = new List<string>();
            int i = 1;
            while (htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table/tr[" + i + "]/td[2]/a") != null)
            {
                var link = htmlDocument.DocumentNode.SelectSingleNode("//td[@class='tdcont']/table/tr[" + i + "]/td[2]/a");
                univerLinkCollection.Add(rootPage + link.Attributes["href"].Value);
                i++;
            }

            return univerLinkCollection;
        }

        static IDictionary<string, string> GetTowns(string searchPage)
        {
            var htmlDocument = GetHtmlDocument(searchPage);
            var townRoot = htmlDocument.DocumentNode.SelectSingleNode("//select[@name='town_']");
            var towns = new Dictionary<string, string>();
            foreach (var optionNode in townRoot.ChildNodes)
            {
                if (optionNode.Name != "#text" && optionNode.Attributes["value"].Value != string.Empty && optionNode.StreamPosition > 13183)
                {
                    towns.Add(optionNode.Attributes["value"].Value, optionNode.NextSibling.InnerHtml.Clean());
                }
            }
            return towns;
        }

        static async Task<HtmlDocument> GetHtmlDocumentAsync(string url)
        {
            var document = new HtmlDocument();
            var webReq = (HttpWebRequest)WebRequest.Create(url);
            using (var response = await webReq.GetResponseAsync())
            {
                using (var responseStream = response.GetResponseStream())
                {
                    document.Load(responseStream);
                }
            }

            return document;
        }

        static HtmlDocument GetHtmlDocument(string url)
        {
            var document = new HtmlDocument();
            var stream = WebRequest.Create(url).GetResponse().GetResponseStream();
            var reader = new StreamReader(stream, Encoding.GetEncoding(1251));
            document.Load(reader);
            return document;
        }

        static string Clean(this String str)
        {
            return str.Replace("&quot;", string.Empty).Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("&nbsp;", string.Empty).Trim();
        }
    }
}
