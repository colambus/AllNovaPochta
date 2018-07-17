using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;

namespace AllNovaPochta
{
    public class Reception
    {
        public string Monday { get; set; }
        public string Tuesday { get; set; }
        public string Wednesday { get; set; }
        public string Thursday { get; set; }
        public string Friday { get; set; }
        public string Saturday { get; set; }
        public string Sunday { get; set; }
    }

    public class Delivery
    {
        public string Monday { get; set; }
        public string Tuesday { get; set; }
        public string Wednesday { get; set; }
        public string Thursday { get; set; }
        public string Friday { get; set; }
        public string Saturday { get; set; }
        public string Sunday { get; set; }
    }

    public class Schedule
    {
        public string Monday { get; set; }
        public string Tuesday { get; set; }
        public string Wednesday { get; set; }
        public string Thursday { get; set; }
        public string Friday { get; set; }
        public string Saturday { get; set; }
        public string Sunday { get; set; }
    }

    //public class Datum
    //{
    //    public string SiteKey { get; set; }
    //    public string Description { get; set; }
    //    public string DescriptionRu { get; set; }
    //    public string ShortAddress { get; set; }
    //    public string ShortAddressRu { get; set; }
    //    public string Phone { get; set; }
    //    public string TypeOfWarehouse { get; set; }
    //    public string Ref { get; set; }
    //    public string Number { get; set; }
    //    public string CityRef { get; set; }
    //    public string CityDescription { get; set; }
    //    public string CityDescriptionRu { get; set; }
    //    public string Longitude { get; set; }
    //    public string Latitude { get; set; }
    //    public string PostFinance { get; set; }
    //    public string BicycleParking { get; set; }
    //    public string PaymentAccess { get; set; }
    //    public string POSTerminal { get; set; }
    //    public string InternationalShipping { get; set; }
    //    public int TotalMaxWeightAllowed { get; set; }
    //    public int PlaceMaxWeightAllowed { get; set; }
    //    public Reception Reception { get; set; }
    //    public Delivery Delivery { get; set; }
    //    public Schedule Schedule { get; set; }
    //    public string DistrictCode { get; set; }
    //    public string WarehouseStatus { get; set; }
    //}

    public class Datum
    {
        public string Ref { get; set; }
        public string SettlementType { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string Description { get; set; }
        public string DescriptionRu { get; set; }
        public string SettlementTypeDescription { get; set; }
        public string SettlementTypeDescriptionRu { get; set; }
        public string Region { get; set; }
        public string RegionsDescription { get; set; }
        public string RegionsDescriptionRu { get; set; }
        public string Area { get; set; }
        public string AreaDescription { get; set; }
        public string AreaDescriptionRu { get; set; }
        public string Index1 { get; set; }
        public string Index2 { get; set; }
        public string IndexCOATSU1 { get; set; }
        public string Delivery1 { get; set; }
        public string Delivery2 { get; set; }
        public string Delivery3 { get; set; }
        public string Delivery4 { get; set; }
        public string Delivery5 { get; set; }
        public string Delivery6 { get; set; }
        public string Delivery7 { get; set; }
        public string Warehouse { get; set; }
        public List<string> Conglomerates { get; set; }
    }

    public class RootObject
    {
        public bool success { get; set; }
        public List<Datum> data { get; set; }
    }

    public class NovaPochta
    {
        public string calledMethod { get; set; }
        public string modelName { get; set; }
        public string apiKey { get; set; }
        public MethodProperties methodProperties { get; set; }
    }

    public class MethodProperties
    {
        public int Page { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            HttpClient client = new HttpClient();
            ExcelPackage package = new ExcelPackage(new System.IO.FileInfo("Cities.xlsx"));
            int i = 1, page = 1;
            int errors = 0;


            while (errors<10)
            {
                var nava = new NovaPochta
                {
                    apiKey = "[ВАШ КЛЮЧ]",
                    calledMethod = "getSettlements",
                    modelName = "AddressGeneral",
                    methodProperties = new MethodProperties
                    {
                        Page = page
                    }
                };
                var jNova = JsonConvert.SerializeObject(nava);

                var content = new StringContent(
                        jNova,
                        Encoding.UTF8,
                        "application/json");

                client.DefaultRequestHeaders
                    .Accept
                    .Add(new MediaTypeWithQualityHeaderValue("text/xml"));
                var response = client.PostAsync("http://testapi.novaposhta.ua/v2.0/json/AddressGeneral/getWarehouses", content).Result;

                var result = response.Content.ReadAsStringAsync().Result;
                RootObject s = JsonConvert.DeserializeObject<RootObject>(result);

                ExcelWorksheet worksheet;
                if (package.Workbook.Worksheets.Count == 0)
                    worksheet = package.Workbook.Worksheets.Add("First");
                else
                    worksheet = package.Workbook.Worksheets.First();


                //foreach (var el in s.data)
                //{
                //    worksheet.Cells[i, 1].Value = el.Description;
                //    worksheet.Cells[i, 2].Value = el.Number;
                //    worksheet.Cells[i, 3].Value = el.CityDescription;
                //    i++;
                //}

                if (s.data == null)
                {
                    if (errors == 1)
                    {
                        Thread.Sleep(30000);
                        errors = 0;
                        continue;
                    }
                    else
                    {                       
                        errors++;
                        continue;
                    }
                }
                else {
                    if (s.data.Count() == 0)
                        errors++;
                    else
                    {
                        errors = 0;
                    }
                }

                foreach (var el in s.data)
                {
                    worksheet.Cells[i, 1].Value = el.Description;
                    worksheet.Cells[i, 2].Value = el.AreaDescription;
                    i++;
                }

                page++;
            }
            package.Save();

            Console.ReadKey();
        }
    }
}
