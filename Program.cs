using OfficeOpenXml;
using System.Globalization;
using System.Xml.Linq;

namespace ExcelToXml
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.License.SetNonCommercialPersonal("shubham lavangare");

            string outputPath = @"C:\Users\admin\Documents\25914\1111\document.xml";
            string[] allowedIDs = { "ID07351", "ID07359" };

            var root = new XElement("orders");

            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\admin\Documents\25914\sampledatafoodsales (1).xlsx")))
            {
                var ws = package.Workbook.Worksheets["FoodSales"];
                if (ws == null)
                {
                    Console.WriteLine("Worksheet 'FoodSales' not found.");
                    return;
                }

                for (int row = 2; row <= ws.Dimension.End.Row; row++)
                {
                    var id = ws.Cells[row, 1].Text;
                    if (!allowedIDs.Contains(id))
                        continue;

                    // Read order date from Excel (column 2)
                    var dateText = ws.Cells[row, 2].Text;
                    DateTime orderDateValue = DateTime.TryParse(dateText, out var dt) ? dt : DateTime.MinValue;

                    var city = ws.Cells[row, 3].Text;
                    var category = ws.Cells[row, 4].Text;
                    var product = ws.Cells[row, 5].Text;
                    var qtyText = ws.Cells[row, 6].Text;
                    var contact = ws.Cells[row, 7].Text;
                    var unitPriceText = ws.Cells[row, 8].Text;

                    int quantity = int.TryParse(qtyText, out int q) ? q : 0;
                    decimal price = decimal.TryParse(unitPriceText, NumberStyles.Any, CultureInfo.InvariantCulture, out var p) ? p : 0M;
                    decimal total = quantity * price;

                    string name = contact.Split(',')[0];
                    string address = contact.Contains(",") ? contact.Substring(contact.IndexOf(',') + 1).Trim() : "";
                    string region = address.Length > 9 ? address.Substring(address.Length - 14).Trim() : "N/A";

                    var shipOrder = new XElement("shiporder",
                        new XAttribute("orderid", id),
                        new XAttribute("orderdate", orderDateValue == DateTime.MinValue ? "" : orderDateValue.ToString("dd-MMM")),
                        new XElement("orderperson", "Customer"),
                        new XElement("shipto",
                            new XElement("name", name),
                            new XElement("address", address),
                            new XElement("city", city),
                            new XElement("region", region)
                        ),
                        new XElement("item",
                            new XElement("title", category + " - " + product),
                            new XElement("note", ""),
                            new XElement("quantity", quantity),
                            new XElement("price", price),
                            new XElement("total", total)
                        )
                    );

                    root.Add(shipOrder);
                }
            }

            var doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
            doc.Save(outputPath);

            Console.WriteLine($"XML file created at: {outputPath}");
        }
    }
}
