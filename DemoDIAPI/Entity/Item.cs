using System;

namespace DemoDIAPI.Classes
{
    public class Item
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public string ItemGroup { get; set; }

        public Item(string itemCode, string itemName,  string itemGroup)
        {
            ItemCode = itemCode;
            ItemName = itemName;
            ItemGroup = itemGroup;
        }

        public override string ToString()
        {
            return $"{ItemCode} - {ItemName}";
        }
    }
}