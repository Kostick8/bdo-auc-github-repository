using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace GetItemInfo3
{
    internal class Program
    {
        private static readonly List<string> ListItemInfo = new List<string> { };
        private static string StartCell = "A2";
        private static string StartCell2 = "A2";
        private static string StartCellDate = "A2";
        private static int StepColumn = 10;
        private static int StepRow = 100;
        private static string[] Options;
        private static SheetsService sheetsService = null;
        private static string spreadsheetId;
        private static string Project_id;
        private static string Region = "ru";
        private static int TotalData = 90;
        private static int PauseAfterError = 45000;
        private static int Attempts = 5;
        private static int TotalDate = 10;
        private static readonly string UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 YaBrowser/24.4.0.0 Safari/537.36";
        private static bool LogError = false;



        static async Task Main()
        {           
            for (int i = 0; i < 3; i++) Console.Beep();
            ReadItems();
            ReadOptions();
            _ = await GetDate(StartCellDate);
            _ = await GetInfo();
            for (int i = 0; i < 3; i++) Console.Beep();
        }


        private static async Task<bool> GetDate(string Cell)
        {           
            IList<IList<object>> ListResult3 = GoogleTable.GetSheet(sheetsService, spreadsheetId, Cell);          
            if (ListResult3 != null)
            {
                int ListResult3CountRow = ListResult3.Count;
                int ListResult3CountColumn = ListResult3[0].Count;
                int TotalDateList = ListResult3[1].Count - 2;

                Match matchDay = null;
                if (TotalDateList != 0)
                    matchDay = Regex.Match((string)ListResult3[0][TotalDateList + 1], @"(\d+)\.\d+\.\d+");

                if (TotalDateList == 0 || matchDay.Success)
                {                   
                    int Day;
                    if (TotalDateList == 0) Day = -1;
                    else Day = int.Parse(matchDay.Groups[1].Value);



                    if (TotalDateList == 0 || DateTime.Now.Day != Day)
                    {
                        int ListItemInfoCount = ListItemInfo.Count;
                        int IDsLenght = ListItemInfoCount / 16;
                        if (ListItemInfoCount % 16 != 0) IDsLenght += 1;

                        string[] ID = new string[IDsLenght];
                        for (int k = 0; k < ListItemInfoCount; k += 16)
                        {
                            string TmpID = "";
                            for (int i = 0; i < 16; i++)
                            {

                                if (i + k >= ListItemInfoCount) break;

                                TmpID += ListItemInfo[i + k] + ",";
                            }
                            ID[k / 16] = TmpID.Substring(0, TmpID.Length - 1);
                        }

                        List<IList<object>> ListResult4 = new List<IList<object>> { };

                        string Result = "";
                        int IDLenght = ID.Length;

                        for (int i = 0; i < IDLenght; i++)
                        {
                            int TotalAttempts = 0;
                            while (TotalAttempts < Attempts)
                            {
                                using (var client = new HttpClient())
                                {
                                    client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
                                    using (var request = new HttpRequestMessage(HttpMethod.Post, "https://api.arsha.io/v2/" + Region + "/GetWorldMarketSubList?lang=ru"))
                                    {
                                        using (var content = new StringContent("[" + ID[i] + "]", null, "application/json"))
                                        {
                                            request.Content = content;
                                            var response = await client.SendAsync(request);
                                            try
                                            {
                                                response.EnsureSuccessStatusCode();
                                                Result += await response.Content.ReadAsStringAsync();
                                                break;
                                            }
                                            catch (Exception ex)
                                            {
                                                LogErrorWrite("GetWorldMarketSubListData\r\nGetWorldMarketSubList\r\n" + ex.Message);
                                                Thread.Sleep(PauseAfterError);
                                            }
                                            TotalAttempts++;
                                        }
                                    }
                                }
                            }

                            if (TotalAttempts == Attempts)
                            {
                                LogErrorWrite("GetDate, не удалось получить данные.\r\nПрограмма будет закрыта.");
                                Environment.Exit(0);
                            }
                        }

                        Regex NameItem = new Regex("\"name\":\"(.+?)\"");
                        Regex IDItem = new Regex("\"id\":(\\d+)");
                        Regex totalTradesItem = new Regex("\"totalTrades\":(\\d+)");

                        MatchCollection MatchName = NameItem.Matches(Result);
                        MatchCollection IDName = IDItem.Matches(Result);
                        MatchCollection MatchstotalTrades = totalTradesItem.Matches(Result);

                        int NameItemCount = MatchName.Count;
                        int IDItemCount = IDName.Count;
                        int MatchstotalTradesCount = MatchstotalTrades.Count;
                        int ListResult4Column = TotalDate + 2;

                        int StepColumn = 0;
                        if (TotalDate == TotalDateList) StepColumn = 1;
                        else if (TotalDate < TotalDateList)
                            StepColumn = TotalDateList - TotalDate + 1;

                        var ItemDate = new List<object>
                        {
                            ListResult3[0][0],  // предмет
                            ListResult3[0][1]  // id
                        };

                        for (int j = 0; j < TotalDateList - StepColumn; j++)
                            ItemDate.Add(ListResult3[0][StepColumn + j + 2]);  // добавить даты



                        ItemDate.Add(DateTime.Now.ToString("dd.MM.yy"));  // добавить сегодняшнюю дату

                        int ItemDateCount = ItemDate.Count;
                        for (int j = 0; j < ListResult4Column - ItemDateCount; j++)
                            ItemDate.Add(DateTime.Now.AddDays(j + 1).ToString("dd.MM.yy"));  // добавить следующие даты

                        ListResult4.Add(ItemDate);

                        // для всех строк с предметами
                        for (int i = 0; i < ListResult3CountRow - 1; i++)
                        {
                            var ItemInformation = new List<object> { };

                            if (i < NameItemCount) ItemInformation.Add(MatchName[i].Groups[1].Value);  // предмет
                            if (i < IDItemCount) ItemInformation.Add(IDName[i].Groups[1].Value);  // id





                            // добавить даты и totalTrades, которые есть в таблице
                            for (int j = 0; j < TotalDateList - StepColumn; j++)
                                ItemInformation.Add(ListResult3[i + 1][StepColumn + j + 2]);

                            // добавить сегодняшние данные
                            if (i < MatchstotalTradesCount) ItemInformation.Add(MatchstotalTrades[i].Groups[1].Value);

                            int Count;
                            if (ListResult3CountColumn > ListResult4Column) Count = ListResult3CountColumn - ItemInformation.Count;
                            else Count = ListResult4Column - ItemInformation.Count;

                            // добавить пустые строки
                            for (int j = 0; j < Count; j++) ItemInformation.Add("");

                            ListResult4.Add(ItemInformation);
                        }
                        GoogleTable.ClearSheet(sheetsService, spreadsheetId, Cell);
                        GoogleTable.SetValue(sheetsService, spreadsheetId, ListResult4, Cell, "ROWS");
                        return true;
                    }
                }             
            }
            else LogErrorWrite("Таблица пустая");
            return false;
        }

        private static async Task<bool> GetInfo()
        {
            int ListItemInfoCount = ListItemInfo.Count;
            int IDsLenght = ListItemInfoCount / 16;
            if (ListItemInfoCount % 16 != 0) IDsLenght += 1;

            string[] ID = new string[IDsLenght];
            for (int k = 0; k < ListItemInfoCount; k += 16)
            {
                string TmpID = "";
                for (int i = 0; i < 16; i++)
                {
                    if (i + k >= ListItemInfoCount) break;
                    TmpID += ListItemInfo[i + k] + ",";
                }
                ID[k / 16] = TmpID.Substring(0, TmpID.Length - 1);
            }

            List<IList<object>> ListResult = await GetWorldMarketSubList(ID, ListItemInfoCount);          
            if (ListResult.Count > 0)
            {
                try
                {
                    GoogleTable.SetValue(sheetsService, spreadsheetId, ListResult, StartCell, "ROWS");
                }
                catch (Exception ex)
                {
                    LogErrorWrite(ex.Message);
                }
            }
 
            string[] IDs = new string[IDsLenght];
            string[] sids = new string[IDsLenght];
            for (int k = 0; k < ListItemInfoCount; k += 16)
            {
                string TmpID = "";
                string Tmpsid = "";
                for (int i = 0; i < 16; i++)
                {
                    if (i + k >= ListItemInfoCount) break;
                    for (int j = 0; j < 6; j++)
                    {
                        TmpID += ListItemInfo[i + k] + ",";
                        Tmpsid += j.ToString() + ",";
                    }

                }
                IDs[k / 16] = TmpID.Substring(0, TmpID.Length - 1);
                sids[k / 16] = Tmpsid.Substring(0, Tmpsid.Length - 1);
            }

            ListResult = await GetListInfo(IDs, sids, ListItemInfoCount);
            if (ListResult.Count > 0)
            {
                try
                {
                    GoogleTable.SetValue(sheetsService, spreadsheetId, ListResult, StartCell2, "COLUMNS");
                }
                catch (Exception ex)
                {
                    LogErrorWrite(ex.Message);
                }
            }
            return true;
        }

        async public static Task<List<IList<object>>> GetListInfo(string[] IDs, string[] sids, int TotalItems)
        {
            List<IList<object>> ListResult = new List<IList<object>> { };
            string BiddingInfoList = "";
            string MarketPriceInfo = "";
            int IDsLenght = IDs.Length;

            for (int i = 0; i < IDsLenght; i++)
            {
                int TotalAttempts = 0;
                while (TotalAttempts < Attempts)
                {
                    using (var client = new HttpClient())
                    {
                        client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
                        using (var request = new HttpRequestMessage(HttpMethod.Get, "https://api.arsha.io/v2/" + Region + "/GetBiddingInfoList?id=" + IDs[i] + "&sid=" + sids[i] + "&lang=ru"))
                        {

                            var response = await client.SendAsync(request);
                            try
                            {
                                response.EnsureSuccessStatusCode();
                                BiddingInfoList += await response.Content.ReadAsStringAsync();
                                break;
                            }
                            catch (Exception ex)
                            {
                                LogErrorWrite("GetListInfo\r\nGetBiddingInfoList\r\n"  + ex.Message);
                                Thread.Sleep(PauseAfterError);
                            }
                            TotalAttempts++;
                        }
                    }
                }
                if (TotalAttempts == Attempts)
                {
                    LogErrorWrite("GetListInfo, не удалось получить данные.\r\nПрограмма будет закрыта.");
                    Environment.Exit(0);
                }

                TotalAttempts = 0;
                while (TotalAttempts < Attempts)
                {
                    using (var client = new HttpClient())
                    {
                        client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
                        using (var request = new HttpRequestMessage(HttpMethod.Get, "https://api.arsha.io/v2/" + Region + "/GetMarketPriceInfo?id=" + IDs[i] + "&sid=" + sids[i] + "&lang=ru"))
                        {
                            var response = await client.SendAsync(request);
                            try
                            {
                                response.EnsureSuccessStatusCode();
                                MarketPriceInfo += await response.Content.ReadAsStringAsync();
                                break;
                            }
                            catch (Exception ex)
                            {
                                LogErrorWrite("GetListInfo\r\nGetMarketPriceInfo\r\n" + ex.Message);
                                Thread.Sleep(PauseAfterError);
                            }
                            TotalAttempts++;
                        }
                    }
                }
                if (TotalAttempts == Attempts)
                {
                    LogErrorWrite("GetListInfo, не удалось получить данные.\r\nПрограмма будет закрыта.");
                    Environment.Exit(0);
                }
            }         

            //File.WriteAllText(@"BiddingInfoList.txt", BiddingInfoList);
            //File.WriteAllText(@"MarketPriceInfo.txt", MarketPriceInfo);

            Regex regexItemInfo = new Regex("name.+?(?=name|$)");
            MatchCollection matchItemInfo = regexItemInfo.Matches(BiddingInfoList);

            int matchItemInfoCount = matchItemInfo.Count;
            MatchCollection[] matchCollectionsName = new MatchCollection[TotalItems * 6];
            MatchCollection[] matchCollectionsID = new MatchCollection[TotalItems * 6];
            MatchCollection[] matchCollectionsBuyers = new MatchCollection[TotalItems * 6];
            MatchCollection[] matchCollectionsPrice = new MatchCollection[TotalItems * 6];
            MatchCollection[] matchCollectionsSellers = new MatchCollection[TotalItems * 6];

            Regex NameItem = new Regex("name\":\"(.+?)\"");
            Regex IDItem = new Regex("\"id\":(\\d+)");
            Regex priceItem = new Regex("price\":(\\d+)");
            Regex buyersItem = new Regex("buyers\":(\\d+)");
            Regex sellersItem = new Regex("sellers\":(\\d+)");
            Regex regexMarketPriceInfo = new Regex("\"(\\d+)\":(\\d+)");

            for (int i = 0; i < matchItemInfoCount; i++)
            {
                matchCollectionsName[i] = NameItem.Matches(matchItemInfo[i].Value);
                matchCollectionsID[i] = IDItem.Matches(matchItemInfo[i].Value);
                matchCollectionsBuyers[i] = buyersItem.Matches(matchItemInfo[i].Value);
                matchCollectionsPrice[i] = priceItem.Matches(matchItemInfo[i].Value);
                matchCollectionsSellers[i] = sellersItem.Matches(matchItemInfo[i].Value);
            }

            MatchCollection matchMarketPriceInfo = regexMarketPriceInfo.Matches(MarketPriceInfo);

            int TotalColumn = StepColumn * 6;

            for (int i = 0; i < TotalColumn; i++)  // заполнить все столбцы
            {
                var ListObject = new List<object>() { };

                if (i % StepColumn == 0)  // имя предмета
                {
                    for (int j = 0; j < matchItemInfoCount; j += 6)
                    {
                        int matchCollectionsPriceCount = matchCollectionsPrice[j + i / StepColumn].Count;
                        for (int k = 0; k < matchCollectionsPriceCount; k++)
                            ListObject.Add(matchCollectionsName[j][0].Groups[1].Value);

                        int StepRow2 = StepRow;
                        if (j == 0) StepRow2 -= 2;
                        for (int k = matchCollectionsPriceCount; k < StepRow2; k++) ListObject.Add("");
                    }
                    ListResult.Add(ListObject);
                    continue;
                }

                if (i % StepColumn == 1)  // id
                {
                    for (int j = 0; j < matchItemInfoCount; j += 6)
                    {
                        int matchCollectionsPriceCount = matchCollectionsPrice[j + i / StepColumn].Count;
                        for (int k = 0; k < matchCollectionsPriceCount; k++)
                            ListObject.Add(matchCollectionsID[j][0].Groups[1].Value);

                        int StepRow2 = StepRow;
                        if (j == 0) StepRow2 -= 2;
                        for (int k = matchCollectionsPriceCount; k < StepRow2; k++) ListObject.Add("");
                    }
                    ListResult.Add(ListObject);
                    continue;
                }

                if (i % StepColumn == 2)  // price
                {

                    for (int j = 0; j < matchItemInfoCount; j += 6)
                    {
                        int matchCollectionsPriceCount = matchCollectionsPrice[j + i / StepColumn].Count;
                        List<int> ListIndex = SortPrice(matchCollectionsPrice[j + i / StepColumn]);

                        for (int k = 0; k < matchCollectionsPriceCount; k++)
                            ListObject.Add(matchCollectionsPrice[j + i / StepColumn][ListIndex[k]].Groups[1].Value);

                        int StepRow2 = StepRow;
                        if (j == 0) StepRow2 -= 2;
                        for (int k = matchCollectionsPriceCount; k < StepRow2; k++) ListObject.Add("");
                    }
                    ListResult.Add(ListObject);
                    continue;
                }

                if (i % StepColumn == 3)
                {
                    for (int j = 0; j < matchItemInfoCount; j += 6)
                    {
                        int matchCollectionsPriceCount = matchCollectionsPrice[j + i / StepColumn].Count;
                        List<int> ListIndex = SortPrice(matchCollectionsPrice[j + i / StepColumn]);

                        for (int k = 0; k < matchCollectionsPriceCount; k++)
                            ListObject.Add(matchCollectionsBuyers[j + i / StepColumn][ListIndex[k]].Groups[1].Value);

                        int StepRow2 = StepRow;
                        if (j == 0) StepRow2 -= 2;
                        for (int k = matchCollectionsPriceCount; k < StepRow2; k++) ListObject.Add("");
                    }
                    ListResult.Add(ListObject);
                    continue;
                }

                if (i % StepColumn == 4)
                {
                    for (int j = 0; j < matchItemInfoCount; j += 6)
                    {
                        int matchCollectionsPriceCount = matchCollectionsPrice[j + i / StepColumn].Count;
                        List<int> ListIndex = SortPrice(matchCollectionsPrice[j + i / StepColumn]);

                        for (int k = 0; k < matchCollectionsPriceCount; k++)
                            ListObject.Add(matchCollectionsSellers[j + i / StepColumn][ListIndex[k]].Groups[1].Value);

                        int StepRow2 = StepRow;
                        if (j == 0) StepRow2 -= 2;
                        for (int k = matchCollectionsPriceCount; k < StepRow2; k++) ListObject.Add("");
                    }
                    ListResult.Add(ListObject);
                    continue;
                }

                if (i % StepColumn == 5)  // дата
                {
                    for (int j = 0; j < matchItemInfoCount; j += 6)
                    {
                        for (int k = 90 - TotalData; k < 90; k++)
                        {
                            long unixTimestamp = long.Parse(matchMarketPriceInfo[k + j * 90 + i / StepColumn * 90].Groups[1].Value) / 1000;
                            DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                            DateTime dateTime = epoch.AddSeconds(unixTimestamp);
                            ListObject.Add(dateTime.ToString("dd.MM.yy"));
                        }

                        int EmptyStringCount = StepRow - TotalData;
                        if (j == 0) EmptyStringCount -= 2;
                        for (int k = 0; k < EmptyStringCount; k++) ListObject.Add("");
                    }
                    ListResult.Add(ListObject);
                    continue;
                }

                if (i % StepColumn == 6)  // цена
                {
                    for (int j = 0; j < matchItemInfoCount; j += 6)
                    {
                        for (int k = 90 - TotalData; k < 90; k++)
                            ListObject.Add(matchMarketPriceInfo[k + j * 90 + i / StepColumn * 90].Groups[2].Value);

                        int EmptyStringCount = StepRow - TotalData;
                        if (j == 0) EmptyStringCount -= 2;
                        for (int k = 0; k < EmptyStringCount; k++) ListObject.Add("");
                    }

                    ListResult.Add(ListObject);
                    continue;
                }

                // заполнить пустой столбец
                for (int j = 0; j < matchItemInfoCount / 6 * StepRow - 2; j++) ListObject.Add("");

                ListResult.Add(ListObject);
            }
            return ListResult;
        }
        

        private static List<int> SortPrice(MatchCollection matchCollectionPrice)
        {
            List<int> ListIndex = new List<int>();
            int matchCollectionPriceCount = matchCollectionPrice.Count;
            while (ListIndex.Count < matchCollectionPriceCount)
            {
                int Index = 0;
                long MaxValue = -1;
                for (int j = 0; j < matchCollectionPriceCount; j++)
                {
                    long Value = long.Parse(matchCollectionPrice[j].Groups[1].Value);
                    if (Value > MaxValue && !ListIndex.Contains(j))
                    {
                        MaxValue = Value;
                        Index = j;
                    }
                }
                ListIndex.Add(Index);
            }
            return ListIndex;
        }

        private static void ReadItems()
        {
            string[] ListItems = File.ReadAllLines("Items.txt");
            for (int i = 0; i < ListItems.Length; i++)
            {
                if (ListItems[i] == "") continue;
                Match match = Regex.Match(ListItems[i], @"\d+", RegexOptions.RightToLeft);
                if (match.Success)
                {
                    if (!ListItemInfo.Contains(match.Value.Trim()))
                        ListItemInfo.Add(match.Value.Trim());
                }
                else LogErrorWrite(ListItems[i] + "\r\nID не найден.");
            }
        }

        private static void ReadOptions()
        {
            Options = File.ReadAllLines("Options.txt");

            Match matchsProject_id = Regex.Match(Options[2], @"(?<=\=).+");
            if (matchsProject_id.Success) Project_id = matchsProject_id.Value.Trim();
            else
            {
                LogErrorWrite("Project_id не найден");
                Environment.Exit(0);
            }

            sheetsService = GoogleTable.AuthorizeGoogleApp(Project_id);

            Match matchStartCell = Regex.Match(Options[0], @"(?<=\=).+");
            if (matchStartCell.Success) StartCell = matchStartCell.Value.Trim();

            Match matchspreadsheetId = Regex.Match(Options[1], @"(?<=\=).+");
            if (matchspreadsheetId.Success) spreadsheetId = matchspreadsheetId.Value.Trim();
            else
            {
                LogErrorWrite("ID таблицы не найден");
                Environment.Exit(0);
            }

            Match matchStartCell2 = Regex.Match(Options[3], @"(?<=\=).+");
            if (matchStartCell2.Success) StartCell2 = matchStartCell2.Value.Trim();

            Match matchStepColumn = Regex.Match(Options[4], @"\d+");
            if (matchStepColumn.Success) StepColumn = int.Parse(matchStepColumn.Value);

            Match matchStepRow = Regex.Match(Options[5], @"\d+");
            if (matchStepRow.Success) StepRow = int.Parse(matchStepRow.Value);

            Match matchRegion = Regex.Match(Options[6], @"(?<=\=).+");
            if (matchRegion.Success)
            {
                if (Regex.IsMatch(matchRegion.Value.Trim(), "^(na|eu|sea|mena|kr|ru|jp|th|tw|sa)$")) Region = matchRegion.Value.Trim();
                else
                {
                    LogErrorWrite("Регион имеет недопустимый формат.\r\nДоступные форматы: na, eu, sea, mena, kr, ru, jp, th, tw, sa.\r\nПрограмма будет закрыта.");
                    Environment.Exit(0);
                }
            }

            Match matchTotalData = Regex.Match(Options[7], @"\d+");
            if (matchTotalData.Success) TotalData = int.Parse(matchTotalData.Value);

            if (StepRow < TotalData)
            {
                LogErrorWrite("StepRow не может быть меньше TotalData.\r\nПрограмма будет закрыта.");
                Environment.Exit(0);
            }

            Match matchStartCellDate = Regex.Match(Options[8], @"(?<=\=).+");
            if (matchStartCellDate.Success) StartCellDate = matchStartCellDate.Value.Trim();


            Match matchPauseAfterError = Regex.Match(Options[9], @"\d+");
            if (matchPauseAfterError.Success) PauseAfterError = int.Parse(matchPauseAfterError.Value);

            Match matchAttempts = Regex.Match(Options[10], @"\d+");
            if (matchAttempts.Success) Attempts = int.Parse(matchAttempts.Value);

            Match matchTotalDate = Regex.Match(Options[11], @"\d+");
            if (matchTotalDate.Success) TotalDate = int.Parse(matchTotalDate.Value);

            Match matchLogError = Regex.Match(Options[12], @"(?<=\=).+");
            if (matchLogError.Success) LogError = bool.Parse(matchLogError.Value.Trim());
        }

        async public static Task<List<IList<object>>> GetWorldMarketSubList(string[] ID, int ItemTotal)
        {
            List<IList<object>> ListResult = new List<IList<object>> { };

            string Result = "";
            int IDLenght = ID.Length;

            for (int i = 0; i < IDLenght; i++)
            {
                int TotalAttempts = 0;
                while (TotalAttempts < Attempts)
                {
                    using (var client = new HttpClient())
                    {
                        client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
                        using (var request = new HttpRequestMessage(HttpMethod.Post, "https://api.arsha.io/v2/" + Region + "/GetWorldMarketSubList?lang=ru"))
                        {
                            var content = new StringContent("[" + ID[i] + "]", null, "application/json");
                            request.Content = content;
                            var response = await client.SendAsync(request);
                            try
                            {
                                response.EnsureSuccessStatusCode();
                                Result += await response.Content.ReadAsStringAsync();
                                break;
                            }
                            catch (Exception ex)
                            {
                                LogErrorWrite("GetWorldMarketSubList\r\nGetWorldMarketSubList\r\n" + ex.Message);
                                Thread.Sleep(PauseAfterError);
                            }
                            TotalAttempts++;
                        }
                    }
                }
                if (TotalAttempts == Attempts)
                {
                    LogErrorWrite("GetWorldMarketSubList, не удалось получить данные.\r\nПрограмма будет закрыта.");
                    Environment.Exit(0);
                }
            }
            
            //File.WriteAllText(@"GetWorldMarketSubList.txt", Result);

            Regex NameItem = new Regex("\"name\":\"(.+?)\"");
            Regex IDItem = new Regex("\"id\":(\\d+)");
            Regex PriceItem = new Regex("\"basePrice\":(\\d+)");
            Regex CurrentStockItem = new Regex("\"currentStock\":(\\d+)");
            Regex totalTradesItem = new Regex("\"totalTrades\":(\\d+)");

            MatchCollection MatchName = NameItem.Matches(Result);
            MatchCollection IDName = IDItem.Matches(Result);
            MatchCollection MatchPrice = PriceItem.Matches(Result);
            MatchCollection MatchStock = CurrentStockItem.Matches(Result);
            MatchCollection MatchstotalTrades = totalTradesItem.Matches(Result);


            int MatchCount = MatchName.Count;
            for (int i = 0; i < MatchCount; i++)
            {
                var ItemInformation = new List<object>()
                {
                    MatchName[i].Groups[1].Value,
                    IDName[i].Groups[1].Value,
                    MatchPrice[i].Groups[1].Value,
                    MatchStock[i].Groups[1].Value,
                    MatchstotalTrades[i].Groups[1].Value
                };
                ListResult.Add(ItemInformation);
            }

            int EmptyStringCount = StepRow * ItemTotal;

            for (int k = 0; k < EmptyStringCount; k++)
                ListResult.Add(new List<object>() { "", "", "", "", "" });
            return ListResult;
        }

        async public static Task<List<IList<object>>> GetWorldMarketSubListData(string[] ID, int ItemTotal, string Cell)
        {
            List<IList<object>> ListResult = new List<IList<object>> { };

            string Result = "";
            int IDLenght = ID.Length;


            for (int i = 0; i < IDLenght; i++)
            {
                int TotalAttempts = 0;
                while (TotalAttempts < Attempts)
                {
                    using (var client = new HttpClient())
                    {
                        client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
                        using (var request = new HttpRequestMessage(HttpMethod.Post, "https://api.arsha.io/v2/" + Region + "/GetWorldMarketSubList?lang=ru"))
                        {
                            using (var content = new StringContent("[" + ID[i] + "]", null, "application/json"))
                            {
                                request.Content = content;
                                var response = await client.SendAsync(request);
                                try
                                {
                                    response.EnsureSuccessStatusCode();
                                    Result += await response.Content.ReadAsStringAsync();
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    LogErrorWrite ("GetWorldMarketSubListData\r\nGetWorldMarketSubList\r\n" + ex.Message);
                                    Thread.Sleep(PauseAfterError);
                                }
                                TotalAttempts++;
                            }
                        }
                    }
                }
            }

            Regex NameItem = new Regex("\"name\":\"(.+?)\"");
            Regex IDItem = new Regex("\"id\":(\\d+)");
            Regex totalTradesItem = new Regex("\"totalTrades\":(\\d+)");

            MatchCollection MatchName = NameItem.Matches(Result);
            MatchCollection IDName = IDItem.Matches(Result);
            MatchCollection MatchstotalTrades = totalTradesItem.Matches(Result);


            // из столбца D1 перенести данные в столбец C1
            Cell += "!D:D";

            IList<IList<object>> ListResult2 = GoogleTable.GetSheet(sheetsService, spreadsheetId, Cell);
            int ListResult2Count = 0;
            if (ListResult2 != null)
            {
                if (ListResult2.Count > 1) ListResult2Count = ListResult2.Count; // если данные были записаны
            }

            int MatchCount = MatchName.Count;
            for (int i = 0; i < MatchCount; i++)
            {
                string totalTradesC1;
                if (ListResult2 != null && i + 1 < ListResult2Count) totalTradesC1 = (string)ListResult2[i + 1][0];
                else totalTradesC1 = MatchstotalTrades[i].Groups[1].Value;

                var ItemInformation = new List<object>()
                {
                    MatchName[i].Groups[1].Value,
                    IDName[i].Groups[1].Value,
                    totalTradesC1,
                    MatchstotalTrades[i].Groups[1].Value
                };
                ListResult.Add(ItemInformation);
            }

            int EmptyStringCount = StepRow * ItemTotal;

            for (int k = 0; k < EmptyStringCount; k++)
                ListResult.Add(new List<object>() { "", "", "", "" });
            return ListResult;
        }

        private static void LogErrorWrite(string Text)
        {
            Console.Beep();
            if (LogError)
            {
                using (StreamWriter sw = File.AppendText(@"LogError.txt")) sw.WriteLine(DateTime.Now.ToString() + "\r\n" + Text + "\r\n");
            }
            else MessageBox.Show(Text);
        }
    }
}
