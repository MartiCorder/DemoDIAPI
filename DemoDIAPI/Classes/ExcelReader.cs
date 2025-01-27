using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace DemoDIAPI.Classes
{
    using ClosedXML.Excel;

    public class ExcelData
    {
        //Document
        public string DocNum { get; set; }
        public DateTime DocDate { get; set; }
        public string Comments { get; set; }
        //Client
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string FederalTaxId { get; set; }
        //Article
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public string ItemGroup { get; set; }
        //Prices
        public decimal Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal Discount { get; set; }

        public override string ToString()
        {
            return $"DocNum: {DocNum}, Date: {DocDate:yyyy-MM-dd}, Customer: {CardName}, " +
                   $"Item: {ItemName}, Quantity: {Quantity}, Price: {Price:N2}, Discount: {Discount:P}";
        }
    }

    public class ExcelReader
    {
        public List<ExcelData> ReadExcelFile(string filePath)
        {
            var orders = new List<ExcelData>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();

                foreach (var row in rows.Skip(1))
                {
                    try
                    {
                        var order = new ExcelData
                        {
                            DocNum = row.Cell(1).GetString(),
                            DocDate = DateTime.Parse(row.Cell(2).GetString()),
                            Comments = row.Cell(3).GetString(),
                            CardCode = row.Cell(4).GetString(),
                            CardName = row.Cell(5).GetString(),
                            FederalTaxId = row.Cell(6).GetString(),
                            ItemCode = row.Cell(7).GetString(),
                            ItemName = row.Cell(8).GetString(),
                            ItemGroup = row.Cell(9).GetString(),
                            Quantity = decimal.Parse(row.Cell(10).GetString().Replace(",", ".")),
                            Price = decimal.Parse(row.Cell(11).GetString().Replace(",", ".")),
                            Discount = decimal.Parse(row.Cell(12).GetString().Replace(",", "."))
                        };
                        orders.Add(order);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error reading row: {row.RowNumber()}. Error: {ex.Message}");
                    }
                }
            }

            return orders;
        }
    }

    public class ExcelReader2
    {
        public List<ExcelData> ReadExcelFile(string filePath)
        {
            var orders = new List<ExcelData>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();

                foreach (var row in rows.Skip(1))
                {
                    try
                    {
                        var order = new ExcelData
                        {
                            CardCode = row.Cell(1).GetString(),
                            ItemCode = row.Cell(2).GetString(),
                            Quantity = decimal.Parse(row.Cell(3).GetString().Replace(",", "."))
                        };
                        orders.Add(order);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error reading row: {row.RowNumber()}. Error: {ex.Message}");
                    }
                }
            }

            return orders;
        }
    }
}