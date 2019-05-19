using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
namespace BTC_parser
{
    class Program
    {
        static void Main()
        {
            List<BTC> btc = new List<BTC>();
            BTC.BTC_full_info[] btc_full = new BTC.BTC_full_info[2];
            Console.WriteLine("Start");
            byte counter = 0;


            HtmlWeb webDoc = new HtmlWeb();
            HtmlDocument html = webDoc.Load("https://btc.com/stats/diff");
            HtmlNodeCollection nodes = html.DocumentNode.SelectNodes("//td");
            if (nodes != null)
            {

                BTC.BTC_full_info btc_full_obj = new BTC.BTC_full_info();
                foreach (var tag in nodes)
                {
                    if (counter == 7 || counter == 14)
                    {

                        if (counter == 14)
                        {
                            btc_full[1] = new BTC.BTC_full_info().copyARGS(btc_full_obj);
                            break;
                        }
                        else
                            btc_full[0] = new BTC.BTC_full_info().copyARGS(btc_full_obj);

                    }
                    switch (counter)
                    {
                        case 0: case 7: btc_full_obj.height = double.Parse(tag.InnerText.Replace(" ", "").Replace("\n", "")); break;
                        case 2: case 9: btc_full_obj.lvl = tag.InnerText; break;
                        case 5: case 12: btc_full_obj.time_diging = tag.InnerText; break;
                        case 6: case 13: btc_full_obj.power = tag.InnerText; break;
                    }
                    counter++;



                    //Console.WriteLine(tag.InnerText); 

                }
            }

            DateTime date = DateTime.Now;
            for (byte i = 0; i < 3; i++)
            {
                webDoc = new HtmlWeb();
                html = webDoc.Load("https://btc.com/block?date=" + date.Year + "-" + (date.Month < 10 ? "0" + date.Month : date.Month + "") + "-" + date.Day);
                nodes = html.DocumentNode.SelectNodes("//td");
                double lvl_of_hardness = btc_full[0].height;
                if (nodes != null)
                {
                    counter = 0;
                    BTC btc_obj = new BTC();
                    foreach (var tag in nodes)
                    {

                        if (counter > 9)
                        {
                            counter = 0;
                            if (btc_obj.height >= lvl_of_hardness)
                                btc_obj.Full_Info = btc_full[0];
                            else
                                btc_obj.Full_Info = btc_full[1];
                            btc.Add(new BTC().copyARGS(btc_obj));

                        }

                        switch (counter)
                        {
                            case 0: btc_obj.height = double.Parse(tag.InnerText.Replace(" ", "").Replace("\n", "")); break;
                            case 1: btc_obj.ovner = tag.InnerText; break;
                            case 8: btc_obj.date = tag.InnerText.Replace(" ", "").Replace("\n", ""); break;
                        }
                        counter++;


                        Console.WriteLine(tag.InnerText);

                    }
                }
                date = date.AddDays(-1);
            }
            Console.WriteLine("end");

            foreach (var tag in btc)
            {
                //не выводит полное инфо 
                Console.WriteLine(tag.ovner + " " + tag.date + " " + tag.height + " " + tag.Full_Info.time_diging + " " + tag.Full_Info.power + " " + tag.Full_Info.lvl);
            }

            makeExcel(btc);
            Console.ReadLine();

        }
        static void makeExcel(List<BTC> btc)
        {
            //Объявляем приложение
            Excel.Application ex = new Excel.Application();

            //Отобразить Excel
            ex.Visible = true;

            //Количество листов в рабочей книге
            ex.SheetsInNewWorkbook = 1;

            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;

            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            var date = DateTime.Now;
            //Название листа (вкладки снизу)
            sheet.Name = "BTC";

            sheet.Cells[1, 1] = "Высота";
            sheet.Cells[1, 2] = "Время нахождения";
            sheet.Cells[1, 3] = "Потраченное время";
            sheet.Cells[1, 4] = "Владелец";
            sheet.Cells[1, 5] = "Сложность";
            sheet.Cells[1, 6] = "Мощность сети";
            sheet.Cells[1, 7] = "Время создания документа";
            sheet.Cells[1, 8] = date.ToString();
            //Пример заполнения ячеек
            for (int i = 1; i <= btc.Count; i++)
            {

                sheet.Cells[i + 1, 1] = btc[i - 1].height;
                sheet.Cells[i + 1, 2] = btc[i - 1].date;
                sheet.Cells[i + 1, 3] = btc[i - 1].Full_Info.time_diging;
                sheet.Cells[i + 1, 4] = btc[i - 1].ovner;
                sheet.Cells[i + 1, 5] = btc[i - 1].Full_Info.lvl;
                sheet.Cells[i + 1, 6] = btc[i - 1].Full_Info.power;

            }

            


        }
        class BTC
        {
            internal double height { get; set; }
            internal string date { get; set; }
            internal string ovner { get; set; }


            internal BTC_full_info Full_Info;

            internal string[] get()
            {
                return new string[] { height + "", date, Full_Info.time_diging, ovner, Full_Info.lvl, Full_Info.power };
            }
            internal BTC copyARGS(BTC parent)
            {
                this.height = parent.height;
                this.date = parent.date;
                this.ovner = parent.ovner;
                this.Full_Info = new BTC_full_info().copyARGS(parent.Full_Info);
                return this;
            }

            internal struct BTC_full_info
            {
                internal double height { get; set; }
                internal string time_diging { get; set; }
                internal string lvl { get; set; }
                internal string power { get; set; }
                internal BTC_full_info copyARGS(BTC_full_info parent)
                {
                    this.height = parent.height;
                    this.time_diging = parent.time_diging;
                    this.lvl = parent.lvl;
                    this.power = parent.power;
                    return this;
                }
            }
        }

    }
}
