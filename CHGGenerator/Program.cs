using System.Xml.Linq;
using ClosedXML.Excel;

namespace CHGGenerator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<Article> articles = ReadArticlesFromExcel("ArticleFeesNotDeleted.xlsx");
            WriteArticlesToXml(articles, "CHG-ART SDCORP-723746.xml");
        }

        private static List<Article> ReadArticlesFromExcel(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);

            var rows = worksheet.RowsUsed().Skip(1);

            return rows.Select(row => new Article
                {
                    Store = int.Parse(row.Cell("A").Value.ToString()),
                    TaxId = int.Parse(row.Cell("B").Value.ToString()),
                    TaxName = row.Cell("C").Value.ToString(),
                    ArticleNumber = int.Parse(row.Cell("D").Value.ToString()),
                    ArticleName = row.Cell("E").Value.ToString()
                })
                .ToList();
        }

        public static void WriteArticlesToXml(List<Article> articles, string filePath)
        {
            var xmlDoc = new XDocument(
                new XElement("ChangeRequests", new XAttribute("Count", ""), new XAttribute("ID", ""), new XAttribute("Update", ""),
                    new XElement("Header",
                        new XElement("Time", DateTime.Now.ToString("HHmmss")),
                        new XElement("Date", DateTime.Now.ToString("yyyyMMdd"))),
                    new XElement("Requests")));

            foreach (var article in articles)
            {
                xmlDoc.Root.Element("Requests").Add(
                    new XElement("Request", new XAttribute("Type", "Delete"), new XAttribute("Method", "TaxesArticle"),
                        new XElement("ItemID", article.ArticleNumber),
                        new XElement("TaxFeeId", article.TaxId),
                        new XElement("StoreId", article.Store)));
            }

            xmlDoc.Save(filePath);
        }
    }
}