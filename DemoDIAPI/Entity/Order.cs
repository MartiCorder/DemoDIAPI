using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DemoDIAPI.Classes;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace DemoDIAPI.Entity
{
    public class Order
    {
        public string DocNum { get; set; }

        public DateTime DocDate { get; set; }
        public string CardCode { get; set; }

        public string CardName { get; set; }
        public string Description { get; set; }
        public List<ExcelData> orderLines { get; set; }

        public Order()
        {
        }

        public Order(string docNum, DateTime docDate, string cardCode, List<ExcelData> orderLine, string description, string cardName)
        {
            DocNum = string.Empty;
            DocDate = DateTime.Now;
            CardCode = cardCode;
            CardName = cardName;
            Description = description;
            this.orderLines = orderLine;
            CardName = cardName;
        }
        public override string ToString()
        {
            return $"DocNum: {DocNum}, CardCode: {CardCode}, CardName: {CardName}, Description: {Description}, OrderLine: {orderLines.ToString()}";
        }
    }
}