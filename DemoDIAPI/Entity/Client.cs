using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoDIAPI.Classes
{
    public class Client
    {
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string FederalTaxId { get; set; }

        public Client(string cardCode, string cardName, string federalTaxId)
        {
            CardCode = cardCode;
            CardName = cardName;
            FederalTaxId = federalTaxId;
        }

        public override string ToString()
        {
            return $"{CardCode} - {CardName}";
        }
    }
}