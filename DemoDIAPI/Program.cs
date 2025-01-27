using DemoDIAPI.Classes;
using DemoDIAPI.Entity;
using DemoDIAPI;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Runtime.InteropServices;

Console.WriteLine("----- Exemple DIAPI -----");
var company = new SAPbobsCOM.Company
{
    Server = "ESONEPC0GW647",
    UserName = "manager",
    Password = "seidor",
    DbServerType = BoDataServerTypes.dst_MSSQL2019,
    CompanyDB = "SBODemoES",
    DbUserName = "sa",
    DbPassword = "SAPB1Admin",
};

Console.WriteLine($"\nProvant de connectar a BBDD: {company.CompanyDB}");
var result = company.Connect();

if (result != 0)
{
    Console.WriteLine(company.GetLastErrorDescription());
    return;
}else
{
    Console.WriteLine("Connexió establerta correctament");
}

Console.WriteLine("\nExecuta: \n1 -> Exercici 1 \n2 -> Exercici 2\n");

var execute = Console.ReadLine();
switch (execute)
{
    case "1":
        Console.WriteLine("\n-----Exercici 1-----\n");
        Exercici1(company);
        break;
    case "2":
        Console.WriteLine("\n-----Exercici 2-----\n");
        Exercici2(company);
        break;
}

static void Exercici1(SAPbobsCOM.Company company)
{
    var reader = new ExcelReader();
    var excelDataList = reader.ReadExcelFile(@"C:\Users\mcorder\source\repos\DemoDIAPI\DemoDIAPI\ExelData\Libro1.xlsx");

    try
    {
        var groupedOrders = excelDataList.GroupBy(x => x.DocNum);
        var orders = new List<Order>();

        foreach (var orderGroup in groupedOrders)
        {
            var firstLineOrder = orderGroup.First();
            var order = new Order
            {
                DocNum = firstLineOrder.DocNum,
                DocDate = firstLineOrder.DocDate,
                CardCode = firstLineOrder.CardCode,
                Description = firstLineOrder.Comments,
                orderLines = orderGroup.ToList()
            };
            orders.Add(order);
        }

        OrderController.ProcessOrders(company, orders);

    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
    }
}

static void Exercici2(SAPbobsCOM.Company company)
{
    var reader = new ExcelReader2();
    var excelDataList = reader.ReadExcelFile(@"C:\Users\mcorder\source\repos\DemoDIAPI\DemoDIAPI\ExelData\Facturas.xlsx");
    
    var groupedLines = excelDataList.GroupBy(l => l.CardCode);

    foreach (var customerGroup in groupedLines)
    {
        try
        {
            Documents oInvoice = (Documents)company.GetBusinessObject(BoObjectTypes.oInvoices);
            var currentCardCode = customerGroup.Key;
            oInvoice.CardCode = currentCardCode;
            oInvoice.DocDate = DateTime.Today;

            bool isFirstLine = true;

            //Creem diccionari per controlar si ja hem usat la linia de la comanda
            var lineTracking = new Dictionary<(int DocEntry, int LineNum), double>();

            foreach (var line in customerGroup)
            {
                string query = $@"
                    SELECT T0.DocEntry, T0.LineNum, T0.OpenQty 
                    FROM RDR1 T0 
                    INNER JOIN ORDR T1 ON T0.DocEntry = T1.DocEntry 
                    WHERE T1.CardCode = '{line.CardCode}' 
                    AND T0.ItemCode = '{line.ItemCode}'
                    AND T0.OpenQty >= {line.Quantity}
                    AND T1.DocStatus = 'O'";

                var recordSet = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordSet.DoQuery(query);

                try
                {
                    if (!isFirstLine)//Si no es la primera linia de la factura n'afegim una
                        oInvoice.Lines.Add();

                    if (!recordSet.EoF)
                    {
                        int docEntry = Convert.ToInt32(recordSet.Fields.Item("DocEntry").Value);
                        int lineNum = Convert.ToInt32(recordSet.Fields.Item("LineNum").Value);
                        double openQty = Convert.ToDouble(recordSet.Fields.Item("OpenQty").Value);

                        if (lineTracking.TryGetValue((docEntry, lineNum), out var usedQty))
                        {
                            if (usedQty + line.Quantity <= openQty) //Sumem quantitatUsada + quantitatNova per veure si hi ha prou openQty
                                lineTracking[(docEntry, lineNum)] += line.Quantity;
                            else
                                continue;
                        }
                        else
                        {
                            if (line.Quantity <= openQty)
                                lineTracking[(docEntry, lineNum)] = line.Quantity;
                            else
                                continue;
                        }

                        oInvoice.Lines.BaseEntry = docEntry;
                        oInvoice.Lines.BaseLine = lineNum;
                        oInvoice.Lines.BaseType = 17;
                    }
                    else
                    {
                        Console.WriteLine($"No open entry found for item {line.ItemCode}. Adding line directly.");
                        oInvoice.Lines.ItemCode = line.ItemCode;
                    }

                    oInvoice.Lines.Quantity = line.Quantity;
                    oInvoice.Lines.UoMEntry = 1;
                    isFirstLine = false;
                }
                finally
                {
                    Utilities.Release(recordSet);
                }
            }

            if (oInvoice.Add() == 0)
                Console.WriteLine($"Invoice successfully created for client {currentCardCode}");
            else
                Console.WriteLine($"Error creating invoice for client {currentCardCode}: {company.GetLastErrorDescription()}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing client {customerGroup.Key}: {ex.Message}");
        }
    }
}

if (company.Connected)
    company.Disconnect();
Utilities.Release(company);