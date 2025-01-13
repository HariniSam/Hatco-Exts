using Mongoose.IDO;
using Mongoose.IDO.Protocol;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;

namespace FTD_MOSplit_FeltHat
{
    public class FeltHat_EvenSplit : IDOExtensionClass
    {
        private string errorMsg = "";
        DataTable filterResults;
        DataTable newMOPs;
        DataTable fullResults;
        private bool updateTasks = true;

        public override void SetContext(IIDOExtensionClassContext context)
        {
            base.SetContext(context);
            Context.IDO.PostLoadCollection += new IDOEventHandler(IDO_PostLoadCollection);
        }

        void IDO_PostLoadCollection(object sender, IDOEventArgs args)
        {
            LoadCollectionResponseData responseData = args.ResponsePayload as LoadCollectionResponseData;
        }

        [IDOMethod(Flags = MethodFlags.None)]
        public void StartProcess(string facility, string sJson, string company, string division, string user, ref string returnMessage)
        {

            string JobID = CreateJobID();
            string email = GetWorkstationID(ref returnMessage);//GetUserData(JobID, user, ref returnMessage);  
            List<string> origSchNos = new List<string>();

            try
            {
                List<int> BGTasks = new List<int>();
                List<int> completeItems = new List<int>();

                CreateLog(JobID, "StartProcess", "", "", "", "FeltHat");
                CreateLog(JobID, sJson, "", "", "", "FeltHat");

                sJson = sJson.Replace("\"", "'").Replace("@#!", ",");
                var dataLst = JsonConvert.DeserializeObject<List<OperationsInfo>>(sJson);

                foreach (var dataItem in dataLst)
                {
                    string styleNumber = dataItem.StyleNumber;
                    string option = dataItem.Option;
                    string plannedQty = dataItem.PlannedQty;
                    string scheduleNo = dataItem.ScheduleNo;
                    string responsible = dataItem.Responsible;
                    string productNo = dataItem.ProducNo;

                    origSchNos.Add(scheduleNo);

                    if (string.IsNullOrEmpty(scheduleNo.Trim()))
                        scheduleNo = "0";

                    string paraList = JobID + " , " + facility + " , " + styleNumber + " , " + option + " , " + plannedQty + " , " + company + " , " + division + " , " + user + " , " + scheduleNo + " , " + responsible + " , " + productNo + " , " + returnMessage;

                    CreateLog(JobID, paraList, "", "", "", "FeltHat");

                    int TaskID = CreateBGTask("FTD_MOSplit_FeltHat_RunProcess", paraList);

                    BGTasks.Add(TaskID);
                }

                while (BGTasks.Count > 0)
                {
                    foreach (int TaskID in BGTasks)
                    {
                        if (BGTaskCompletedAndSuccess(TaskID))
                        {
                            completeItems.Add(TaskID);
                        }
                        else
                        {
                            Thread.Sleep(5000);
                        }
                    }

                    if (completeItems.Count > 0)
                    {
                        foreach (var TaskID in completeItems)
                        {
                            BGTasks.Remove(TaskID);

                        }
                    }

                }

                if (updateTasks)
                {
                    returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(GetAllSchedulNOsByJob(JobID)) ? "None" : GetAllSchedulNOsByJob(JobID));
                }
                else
                {
                    returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(GetAllSchedulNOsByJob(JobID)) ? "None" : GetAllSchedulNOsByJob(JobID)) + " with error, please check audit log for further references";
                }

                ProcessAppEvent(returnMessage, "", email, string.Join(";", origSchNos));
            }
            catch (Exception ex)
            {
                returnMessage = "Process completed with error, please check audit log for further references";
                CreateLog(JobID, ex.Message, "Exception", "", "", "FeltHat");
                ProcessAppEvent(returnMessage, ex.Message, email, string.Join(";", origSchNos));
            }
        }

        [IDOMethod(Flags = MethodFlags.None)]
        public void RunProcessForEachLine(string JobID, string facility, string styleNumber, string option, string plannedQty, string company, string division, string user, string scheduleNo, string responsible, string productNo, out string returnMessage)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            stopwatch.Stop();
            TimeSpan ts = stopwatch.Elapsed;
            string sTimeTaken = (ts.Hours > 0 ? ts.Hours + ":" : "") + (ts.Minutes > 0 ? ts.Minutes + ":" : "") + ts.Seconds + "." + ts.Milliseconds;
            if (productNo.Trim() != "NaN")
                styleNumber = GetSytle(productNo, JobID);


            //if (string.IsNullOrEmpty(scheduleNo.Trim()))
            //    scheduleNo = "0";

            CreateLog(JobID, "RunProcessForEachLine", facility + ", " + styleNumber + ", " + option + ", " + plannedQty + ", " + company + ", " + division + ", " + user + ", " + scheduleNo + ", " + responsible + ", " + productNo, "", sTimeTaken, "FeltHat");

            returnMessage = "";
            //string email = GetUserData(user, ref returnMessage);
            try
            {
                LstMMOPLP02(JobID, facility, styleNumber, option, plannedQty, company, division, scheduleNo, responsible, out returnMessage);
                returnMessage = returnMessage.TrimStart(',');
                //ProcessAppEvent(returnMessage, "", email);
                returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage);

                //CreateLog(JobID, "New Schedules created", returnMessage.Replace("New Schedules created: ",""), "", "", "FeltHat");
            }
            catch (AggregateException ex)
            {
                string errormessage = ex.Message + " Please contact the administrator.";
                returnMessage = returnMessage.TrimStart(',');
                //ProcessAppEvent(returnMessage, errormessage, email);
                returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage) +
                     "<br>Errors: " + errormessage;
                //CreateLog(JobID, "New Schedules created", returnMessage.Replace("New Schedules created: ", ""), "", "", "FeltHat");

                CreateLog(JobID, "Exception - RunProcessForEachLine", ex.StackTrace, "", "", "FeltHat");

            }
            catch (Exception ex)
            {
                returnMessage = returnMessage.TrimStart(',');
                //ProcessAppEvent(returnMessage, ex.Message, email);
                returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage) +
                     "<br>Errors: " + ex.Message;
                CreateLog(JobID, "New Schedules created", returnMessage.Replace("New Schedules created: ", ""), "", "", "FeltHat");
                CreateLog(JobID, "Exception - RunProcessForEachLine", ex.StackTrace, "", "", "FeltHat");
            }
        }

        public void LstMMOPLP02(string JobID, string Facility, string StyleNumber, string Option, string PlannedQty, string company, string division, string scheduleNo, string responsible, out string returnMessage)
        {
            returnMessage = "";
            int TotalQty = 0;
            try
            {
                CreateLog(JobID, "LstMMOPLP02", Facility + ", " + StyleNumber + ", " + Option + ", " + PlannedQty + ", " + company + ", " + division + ", " + scheduleNo + ", " + responsible, "", "", "FeltHat");

                string Infobar = "";
                bool MOSplit = true;
                int maxQty = 0;

                List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
                lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMMOPLP02"));
                lstparms.Add(new Tuple<string, string>("ROFACI", Facility));
                lstparms.Add(new Tuple<string, string>("ROHDPR", StyleNumber));
                lstparms.Add(new Tuple<string, string>("ROOPTY", Option));
                lstparms.Add(new Tuple<string, string>("F_SCHN", scheduleNo));
                lstparms.Add(new Tuple<string, string>("T_SCHN", scheduleNo));
                lstparms.Add(new Tuple<string, string>("F_RESP", responsible));
                lstparms.Add(new Tuple<string, string>("T_RESP", responsible));

                string sMethodList = BuildMethodParms(lstparms);
                string result = ProcessIONAPI(JobID, sMethodList, out Infobar);

                if (string.IsNullOrEmpty(Infobar))
                {
                    JObject rss = JObject.Parse(result);

                    int count = rss["results"][0]["records"].Count();

                    if (count > 0)
                    {
                        fullResults = new DataTable("SKU_Details");
                        fullResults.Columns.Add("ROPLPN");
                        fullResults.Columns.Add("ROPLPS");
                        fullResults.Columns.Add("ROHDPR");
                        fullResults.Columns.Add("ROPRNO");
                        fullResults.Columns.Add("ROORQA", typeof(int));
                        fullResults.Columns.Add("ROOPTX");
                        fullResults.Columns.Add("ROOPTY");
                        fullResults.Columns.Add("ROOPTZ");
                        fullResults.Columns.Add("ROPLDT");
                        fullResults.Columns.Add("ROFACI");
                        fullResults.Columns.Add("RORESP");
                        fullResults.Columns.Add("ROSTDT");
                        fullResults.Columns.Add("ROFIDT");
                        fullResults.Columns.Add("RORORC");
                        fullResults.Columns.Add("RORORN");
                        fullResults.Columns.Add("RORORL");
                        fullResults.Columns.Add("ROWHLO");
                        fullResults.Columns.Add("ROPSTS");
                        fullResults.Columns.Add("RORORX");
                        fullResults.Columns.Add("ROORQT", typeof(int));
                        fullResults.Columns.Add("ROORTY");

                        var oDataList = new List<Model>();

                        for (int i = 0; i < count; i++)
                        {
                            if (i == 0)
                            {
                                string sMaxQty = rss["results"][0]["records"][i]["MMCFI2"].ToString();
                                if (string.IsNullOrEmpty(sMaxQty))
                                {
                                    returnMessage = "Max quantity is not defined for the style";
                                    throw new Exception("Max quantity is not defined for the style");
                                }
                                else if (rss["results"][0]["records"][i]["ROHDPR"].ToString().Substring(1, 2) == "S")
                                {
                                    returnMessage = "Only felt hats are allowed";
                                    throw new Exception("Only felt hats are allowed");
                                }

                                if (!string.IsNullOrEmpty(PlannedQty))
                                {
                                    int iPlannedQty = int.Parse(PlannedQty);
                                    if (!string.IsNullOrEmpty(sMaxQty))
                                    {
                                        maxQty = int.Parse(Math.Floor(Convert.ToDouble(sMaxQty)).ToString());



                                        if (maxQty >= iPlannedQty)
                                        {
                                            MOSplit = false;
                                        }
                                    }
                                }
                            }

                            DataRow newRow = fullResults.NewRow();

                            newRow["ROPLPN"] = rss["results"][0]["records"][i]["ROPLPN"].ToString();
                            newRow["ROPLPS"] = rss["results"][0]["records"][i]["ROPLPS"].ToString();
                            newRow["ROHDPR"] = rss["results"][0]["records"][i]["ROHDPR"].ToString();
                            newRow["ROPRNO"] = rss["results"][0]["records"][i]["ROPRNO"].ToString();
                            if (rss["results"][0]["records"][i]["ROORQA"] != null)
                                newRow["ROORQA"] = int.Parse(Math.Floor(Convert.ToDouble(rss["results"][0]["records"][i]["ROORQA"].ToString())).ToString());
                            else
                                newRow["ROORQA"] = 0;
                            newRow["ROORQA"] = rss["results"][0]["records"][i]["ROORQA"].ToString();
                            newRow["ROOPTX"] = rss["results"][0]["records"][i]["ROOPTX"].ToString();
                            newRow["ROOPTY"] = rss["results"][0]["records"][i]["ROOPTY"].ToString();
                            if (rss["results"][0]["records"][i]["ROOPTZ"] != null)
                                newRow["ROOPTZ"] = rss["results"][0]["records"][i]["ROOPTZ"].ToString();
                            else
                                newRow["ROOPTZ"] = "";
                            newRow["ROPLDT"] = rss["results"][0]["records"][i]["ROPLDT"].ToString();
                            newRow["ROFACI"] = rss["results"][0]["records"][i]["ROFACI"].ToString();
                            newRow["RORESP"] = rss["results"][0]["records"][i]["RORESP"].ToString();
                            newRow["ROSTDT"] = rss["results"][0]["records"][i]["ROSTDT"].ToString();
                            newRow["ROFIDT"] = rss["results"][0]["records"][i]["ROFIDT"].ToString();
                            newRow["RORORC"] = rss["results"][0]["records"][i]["RORORC"].ToString();
                            newRow["RORORN"] = rss["results"][0]["records"][i]["RORORN"].ToString();
                            newRow["RORORL"] = rss["results"][0]["records"][i]["RORORL"].ToString();
                            newRow["ROWHLO"] = rss["results"][0]["records"][i]["ROWHLO"].ToString();
                            newRow["ROPSTS"] = rss["results"][0]["records"][i]["ROPSTS"].ToString();
                            newRow["RORORX"] = rss["results"][0]["records"][i]["RORORX"].ToString();
                            if (rss["results"][0]["records"][i]["ROORQT"] != null)
                                newRow["ROORQT"] = int.Parse(Math.Floor(Convert.ToDouble(rss["results"][0]["records"][i]["ROORQT"].ToString())).ToString());
                            else
                                newRow["ROORQT"] = 0;

                            fullResults.Rows.Add(newRow);



                            oDataList.Add(new Model()
                            {

                                ROPLPN = rss["results"][0]["records"][i]["ROPLPN"] == null ? "" : rss["results"][0]["records"][i]["ROPLPN"].ToString(),
                                ROPLPS = rss["results"][0]["records"][i]["ROPLPS"] == null ? "" : rss["results"][0]["records"][i]["ROPLPS"].ToString(),
                                ROHDPR = rss["results"][0]["records"][i]["ROHDPR"] == null ? "" : rss["results"][0]["records"][i]["ROHDPR"].ToString(),
                                ROPRNO = rss["results"][0]["records"][i]["ROPRNO"] == null ? "" : rss["results"][0]["records"][i]["ROPRNO"].ToString(),
                                ROORQA = int.Parse(rss["results"][0]["records"][i]["ROORQA"] == null ? "0" : rss["results"][0]["records"][i]["ROORQA"].ToString()),
                                ROOPTX = rss["results"][0]["records"][i]["ROOPTX"] == null ? "" : rss["results"][0]["records"][i]["ROOPTX"].ToString(),
                                ROOPTY = rss["results"][0]["records"][i]["ROOPTY"] == null ? "" : rss["results"][0]["records"][i]["ROOPTY"].ToString(),
                                ROOPTZ = rss["results"][0]["records"][i]["ROOPTZ"] == null ? "" : rss["results"][0]["records"][i]["ROOPTZ"].ToString(),
                                ROPLDT = rss["results"][0]["records"][i]["ROPLDT"] == null ? "" : rss["results"][0]["records"][i]["ROPLDT"].ToString(),
                                ROFACI = rss["results"][0]["records"][i]["ROFACI"] == null ? "" : rss["results"][0]["records"][i]["ROFACI"].ToString(),
                                RORESP = rss["results"][0]["records"][i]["RORESP"] == null ? "" : rss["results"][0]["records"][i]["RORESP"].ToString(),
                                ROSTDT = rss["results"][0]["records"][i]["ROSTDT"] == null ? "" : rss["results"][0]["records"][i]["ROSTDT"].ToString(),
                                ROFIDT = rss["results"][0]["records"][i]["ROFIDT"] == null ? "" : rss["results"][0]["records"][i]["ROFIDT"].ToString(),
                                RORORC = rss["results"][0]["records"][i]["RORORC"] == null ? "" : rss["results"][0]["records"][i]["RORORC"].ToString(),
                                RORORN = rss["results"][0]["records"][i]["RORORN"] == null ? "" : rss["results"][0]["records"][i]["RORORN"].ToString(),
                                RORORL = rss["results"][0]["records"][i]["RORORL"] == null ? "" : rss["results"][0]["records"][i]["RORORL"].ToString(),
                                ROWHLO = rss["results"][0]["records"][i]["ROWHLO"] == null ? "" : rss["results"][0]["records"][i]["ROWHLO"].ToString(),
                                ROPSTS = rss["results"][0]["records"][i]["ROPSTS"] == null ? "" : rss["results"][0]["records"][i]["ROPSTS"].ToString(),
                                RORORX = rss["results"][0]["records"][i]["RORORX"] == null ? "" : rss["results"][0]["records"][i]["RORORX"].ToString(),
                                ROORQT = int.Parse(rss["results"][0]["records"][i]["ROORQT"] == null ? "0" : rss["results"][0]["records"][i]["ROORQT"].ToString()),
                                ROORTY = rss["results"][0]["records"][i]["ROORTY"] == null ? "" : rss["results"][0]["records"][i]["ROORTY"].ToString(),

                            });
                        }

                        TotalQty = fullResults.AsEnumerable().Sum(row => row.Field<int>("ROORQA"));
                        //CreateLog(JobID, "TotalQty", TotalQty.ToString(), "", "", "FeltHat");
                        CreateLog(JobID, "Filled model data", "", "", "", "FeltHat");

                        List<Model> testList1 = oDataList.OrderBy(a => a.ROPRNO).GroupBy(row => new { row.ROPRNO })
               .SelectMany(group =>
               {
                   // Calculate sum of ROORQA within the group
                   int sumROORQA = group.Sum(item => item.ROORQA ?? 0);

                   // If the sum is not a multiple of 12 (except for 12 itself), proceed with grouping
                   if (sumROORQA != maxQty && sumROORQA % maxQty != 0)
                   {
                       return group.Select(item => new Model()
                       {
                           ROPLPN = RemoveDuplicates(String.Join(",", group.Select(a => a.ROPLPN))),
                           ROPLPS = group.Select(a => a.ROPLPS).FirstOrDefault(),
                           ROHDPR = group.Select(a => a.ROHDPR).FirstOrDefault(),
                           ROPRNO = group.Select(a => a.ROPRNO).FirstOrDefault(),
                           ROORQA = group.Sum(a => a.ROORQA),
                           ROOPTX = group.Select(a => a.ROOPTX).FirstOrDefault(),
                           ROOPTY = group.Select(a => a.ROOPTY).FirstOrDefault(),
                           ROOPTZ = group.Select(a => a.ROOPTZ).FirstOrDefault(),
                           ROPLDT = group.Select(a => a.ROPLDT).FirstOrDefault(),
                           ROFACI = group.Select(a => a.ROFACI).FirstOrDefault(),
                           RORESP = group.Select(a => a.RORESP).FirstOrDefault(),
                           ROSTDT = group.Select(a => a.ROSTDT).FirstOrDefault(),
                           ROFIDT = group.Select(a => a.ROFIDT).FirstOrDefault(),
                           RORORC = group.Select(a => a.RORORC).FirstOrDefault(),
                           RORORN = group.Select(a => a.RORORN).FirstOrDefault(),
                           RORORL = group.Select(a => a.RORORL).FirstOrDefault(),
                           ROWHLO = group.Select(a => a.ROWHLO).FirstOrDefault(),
                           ROPSTS = group.Select(a => a.ROPSTS).FirstOrDefault(),
                           RORORX = group.Select(a => a.RORORX).FirstOrDefault(),
                           ROORQT = group.Sum(a => a.ROORQT),
                           ROORTY = group.Select(a => a.ROORTY).FirstOrDefault(),
                       });
                   }
                   else // Otherwise, return items without grouping
                   {
                       return group.Select(item => new Model()
                       {
                           ROPLPN = item.ROPLPN,
                           ROPLPS = item.ROPLPS,
                           ROHDPR = item.ROHDPR,
                           ROPRNO = item.ROPRNO,
                           ROORQA = item.ROORQA,
                           ROOPTX = item.ROOPTX,
                           ROOPTY = item.ROOPTY,
                           ROOPTZ = item.ROOPTZ,
                           ROPLDT = item.ROPLDT,
                           ROFACI = item.ROFACI,
                           RORESP = item.RORESP,
                           ROSTDT = item.ROSTDT,
                           ROFIDT = item.ROFIDT,
                           RORORC = item.RORORC,
                           RORORN = item.RORORN,
                           RORORL = item.RORORL,
                           ROWHLO = item.ROWHLO,
                           ROPSTS = item.ROPSTS,
                           RORORX = item.RORORX,
                           ROORQT = item.ROORQT,
                           ROORTY = item.ROORTY,
                       });
                   }
               })
    .ToList();
                        //CreateLog(JobID, "aggregated model data", "", "", "", "FeltHat");
                        CreateLog(JobID, "testList1", testList1.Count().ToString(), "", "", "FeltHat");

                        var groupedByROPRNO = testList1.GroupBy(item => item.ROPRNO);
                        bool hasSameROPRNO = groupedByROPRNO.Any(groupByROPRNO => groupByROPRNO.Count() > 1);

                        fullResults = ToDataTable(testList1);

                        if (MOSplit)
                        {
                            filterResults = new DataTable("SKU_filter_Details");
                            filterResults.Columns.Add("ROPLPN");
                            filterResults.Columns.Add("ROPLPS");
                            filterResults.Columns.Add("ROHDPR");
                            filterResults.Columns.Add("ROPRNO");
                            filterResults.Columns.Add("ROORQA", typeof(int));
                            filterResults.Columns.Add("ROOPTX");
                            filterResults.Columns.Add("ROOPTY");
                            filterResults.Columns.Add("ROOPTZ");
                            filterResults.Columns.Add("ROPLDT");
                            filterResults.Columns.Add("ROFACI");
                            filterResults.Columns.Add("RORESP");
                            filterResults.Columns.Add("ROSTDT");
                            filterResults.Columns.Add("ROFIDT");
                            filterResults.Columns.Add("RORORC");
                            filterResults.Columns.Add("RORORN");
                            filterResults.Columns.Add("RORORL");
                            filterResults.Columns.Add("ROWHLO");
                            filterResults.Columns.Add("ROPSTS");
                            filterResults.Columns.Add("RORORX");
                            filterResults.Columns.Add("ROORQT", typeof(int));
                            filterResults.Columns.Add("ROORTY");

                            //CreateLog(JobID, "Filled model data to filter results - Start", "", "", "", "FeltHat");
                            DataView dv = new DataView(fullResults);
                            dv.Sort = "ROPRNO, ROOPTX";

                            filterResults = dv.ToTable();
                            //CreateLog(JobID, "Filled model data to filter results - End", "", "", "", "FeltHat");

                            newMOPs = new DataTable("MOP_Details");
                            newMOPs.Columns.Add("CONO");
                            newMOPs.Columns.Add("FACI");
                            newMOPs.Columns.Add("WHLO");
                            newMOPs.Columns.Add("PRNO");
                            newMOPs.Columns.Add("STRT");
                            newMOPs.Columns.Add("PPQT");
                            newMOPs.Columns.Add("PLDT");
                            newMOPs.Columns.Add("PLHM");
                            newMOPs.Columns.Add("ORTY");
                            newMOPs.Columns.Add("RORC");
                            newMOPs.Columns.Add("RORN");
                            newMOPs.Columns.Add("RORL");
                            newMOPs.Columns.Add("RORX");
                            newMOPs.Columns.Add("SCHN");
                            newMOPs.Columns.Add("SIMD");
                            newMOPs.Columns.Add("PSTS");
                            newMOPs.Columns.Add("CURR");//HariniS


                            int i = 0;

                            if (maxQty != 0)
                            {

                                List<RequestTransaction> transactions = new List<RequestTransaction>();

                                List<string> selectedColumns = new List<string>();
                                selectedColumns.Add("SCHN");

                                RequestRecord oRequestRecord = new RequestRecord();
                                oRequestRecord.TX40 = StyleNumber;

                                for (i = 0; i < TotalQty - maxQty; i = i + maxQty)
                                {
                                    RequestTransaction oRequestTransaction = new RequestTransaction();
                                    oRequestTransaction.transaction = "AddScheduleNo";
                                    oRequestTransaction.record = oRequestRecord;
                                    oRequestTransaction.selectedColumns = selectedColumns;

                                    transactions.Add(oRequestTransaction);

                                }

                                //CreateLog(JobID, "Added schdule no", "", "", "", "FeltHat");

                                RequestTransaction oRequestTransactionExt = new RequestTransaction();
                                oRequestTransactionExt.transaction = "AddScheduleNo";
                                oRequestTransactionExt.record = oRequestRecord;
                                oRequestTransactionExt.selectedColumns = selectedColumns;
                                transactions.Add(oRequestTransactionExt);

                                RequestRoot oRequestRoot = new RequestRoot();
                                oRequestRoot.program = "PMS270MI";
                                oRequestRoot.cono = int.Parse(company);
                                oRequestRoot.divi = division;
                                oRequestRoot.excludeEmptyValues = false;
                                oRequestRoot.rightTrim = true;
                                oRequestRoot.maxReturnedRecords = 0;
                                oRequestRoot.transactions = transactions;

                                string jsonString = JsonConvert.SerializeObject(oRequestRoot);

                                //CreateLog(JobID, jsonString, "", "", "", "FeltHat");

                                ResponseRoot oResponseRoot = null;

                                string results = CallServicePost(jsonString, JobID);
                                if (results != string.Empty)
                                {
                                    oResponseRoot = JsonConvert.DeserializeObject<ResponseRoot>(results);
                                }

                                if (oResponseRoot != null)
                                {
                                    int listIndex = 0;
                                    List<ResponseResult> resultsList = oResponseRoot.results;
                                    for (i = 0; i < TotalQty - maxQty; i = i + maxQty)
                                    {
                                        //string ScheduleNumber = AddScheduleNo(StyleNumber);
                                        //returnMessage = returnMessage + ", " + ScheduleNumber;
                                        //fill(maxQty, ScheduleNumber, Facility, company);

                                        ResponseResult oResponseResult = resultsList[listIndex];
                                        string ScheduleNumber = oResponseResult.records[0].SCHN;
                                        //string ScheduleNumber = AddScheduleNo(StyleNumber);
                                        returnMessage = returnMessage + ", " + ScheduleNumber;
                                        fill(JobID, maxQty, ScheduleNumber, Facility, company, hasSameROPRNO);
                                        listIndex++;

                                    }
                                    //CreateLog(JobID, "TotalQty:" + TotalQty.ToString(), "i:" + i.ToString(), "", "", "FeltHat");
                                    int diff = TotalQty - i;
                                    //CreateLog(JobID, "diff", diff.ToString(), "", "", "FeltHat");
                                    if (diff > 0)
                                    {
                                        //string ScheduleNumber = AddScheduleNo(StyleNumber);
                                        //returnMessage = returnMessage + ", " + ScheduleNumber;
                                        //fill(diff, ScheduleNumber, Facility, company);
                                        ResponseResult oResponseResult = resultsList[listIndex];
                                        string ScheduleNumber = oResponseResult.records[0].SCHN;
                                        //string ScheduleNumber = AddScheduleNo(StyleNumber);
                                        returnMessage = returnMessage + ", " + ScheduleNumber;
                                        fill(JobID, diff, ScheduleNumber, Facility, company, hasSameROPRNO);
                                    }

                                    if (newMOPs.Rows.Count > 0)
                                    {
                                        CrtMOPRoot oCrtMOPRoot = new CrtMOPRoot();
                                        oCrtMOPRoot.program = "PMS170MI";
                                        oCrtMOPRoot.cono = int.Parse(company);
                                        oCrtMOPRoot.divi = division;
                                        oCrtMOPRoot.excludeEmptyValues = false;
                                        oCrtMOPRoot.rightTrim = true;
                                        oCrtMOPRoot.maxReturnedRecords = 0;

                                        List<CrtMOPTransaction> oListCrtMOPTransaction = new List<CrtMOPTransaction>();

                                        foreach (DataRow dr in newMOPs.Rows)
                                        {
                                            CrtMOPTransaction oCrtMOPTransaction = new CrtMOPTransaction();
                                            oCrtMOPTransaction.record = CreatePlannedMO(dr);
                                            oCrtMOPTransaction.transaction = "CrtPlannedMO";
                                            List<string> listSelectedColumns = new List<string>();
                                            listSelectedColumns.Add("PLPN");
                                            oCrtMOPTransaction.selectedColumns = listSelectedColumns;
                                            oListCrtMOPTransaction.Add(oCrtMOPTransaction);
                                            //string sPLPN = CreatePlannedMO(dr);
                                            //UpdateScheduleNo();
                                        }

                                        oCrtMOPRoot.transactions = oListCrtMOPTransaction;

                                        string crtMOPjsonString = JsonConvert.SerializeObject(oCrtMOPRoot);

                                        string crtMOPresults = CallServicePost(crtMOPjsonString, JobID);
                                        if (crtMOPresults != string.Empty)
                                        {
                                            //oResponseRoot = JsonConvert.DeserializeObject<ResponseRoot>(crtMOPresults);
                                        }

                                        if (oResponseRoot != null)
                                        {

                                        }
                                        //foreach (DataRow dr in newMOPs.Rows)
                                        //{
                                        //    string sPLPN = CreatePlannedMO(dr);
                                        //    //UpdateScheduleNo();
                                        //}

                                        RemoveOrigSchFromMOs(company, division, JobID);
                                        DelPlannedMOs(company, division, JobID);

                                    }
                                }
                                else
                                {
                                    returnMessage = "PMS270MI/AddScheduleNo returned empty results";
                                    throw new Exception("PMS270MI/AddScheduleNo returned empty results");

                                }
                            }
                            else
                            {
                                returnMessage = "Max Qty (MMCFI2) is 0";
                                throw new Exception("Max Qty (MMCFI2) is 0");
                            }


                        }
                        else
                        {
                            CreateLog(JobID, "Max Qty == 0", StyleNumber, "", "", "FeltHat");

                            string schduleNumber = AddScheduleNo(StyleNumber, JobID);
                            returnMessage = returnMessage + ", " + schduleNumber;
                            for (int i = 0; i < fullResults.Rows.Count; i++)
                            {

                                string planOrderNumber = fullResults.Rows[i]["ROPLPN"].ToString();
                                UpdateScheduleNo(schduleNumber, planOrderNumber, JobID);
                                //returnMessage = returnMessage + ", " + schduleNumber;
                            }
                        }

                    }


                }
            }
            catch (AggregateException ex)
            {
                //string errormessage = ex.Message + " Please contact the administrator.";
                //returnMessage = returnMessage.TrimStart(',');
                //ProcessAppEvent(returnMessage, errormessage, email);
                //returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage) +
                //     "<br>Errors: " + errormessage;
                CreateLog(JobID, "New Schedules created", returnMessage.Replace("New Schedules created: ", ""), "", "", "FeltHat");

                CreateLog(JobID, "Exception - LstMMOPLP02", ex.Message, "", "", "FeltHat");

            }
            catch (Exception ex)
            {
                //returnMessage = returnMessage.TrimStart(',');
                ////ProcessAppEvent(returnMessage, ex.Message, email);
                //returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage) +
                //     "<br>Errors: " + ex.Message;
                CreateLog(JobID, "Exception - LstMMOPLP02", ex.Message, "", "", "FeltHat");
                CreateLog(JobID, "New Schedules created", returnMessage.Replace("New Schedules created: ", ""), "", "", "FeltHat");
            }

        }



        public DataTable ToDataTable(List<Model> items)
        {
            DataTable oResultDataTable = new DataTable(typeof(Model).Name);
            string sResultSetColumns = "ROPLPN,	ROPLPS,	ROHDPR,	ROPRNO,	ROORQA,	ROOPTX,	ROOPTY,	ROOPTZ,	ROPLDT,	ROFACI,	RORESP,	ROSTDT,	ROFIDT,	RORORC,	RORORN,	RORORL,	ROWHLO,	ROPSTS,	RORORX,	ROORQT,	ROORTY";

            string[] sResultSetColumnsList = sResultSetColumns.Split(',');

            System.Reflection.PropertyInfo[] properties =
                typeof(Model)
                .GetProperties();

            foreach (System.Reflection.PropertyInfo prop in properties)
            {
                oResultDataTable.Columns.Add(prop.Name);
            }

            foreach (Model item in items)
            {

                var values = new object[properties.Length];

                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(item, null);
                }
                oResultDataTable.Rows.Add(values);
            }

            return oResultDataTable;
        }

        public void fill(string JobID, int qty, string ScheduleNumber, string facility, string company, bool hasSameROPRNO)
        {
            CreateLog(JobID, "New Schedules created", ScheduleNumber, "", "", "FeltHat");
            //CreateLog(JobID, "fill", "qty: " + qty.ToString(), "", "", "FeltHat");
            int total = 0; int curIndex = 0;
            while (total < qty)
            {
                //CreateLog(JobID, "fill", "filterResults.Rows.Count: " + filterResults.Rows.Count.ToString(), "", "", "FeltHat");

                for (int i = 0; i < filterResults.Rows.Count; i++)
                {
                    if (total == qty)
                    {
                        break;
                    }
                    int rowqty = int.Parse(filterResults.Rows[i]["ROORQA"].ToString());
                    //CreateLog(JobID, "fill", "rowqty: " + rowqty.ToString(), "", "", "FeltHat");
                    if (rowqty > 0)
                    {
                        filterResults.Rows[i]["ROORQA"] = rowqty - 1;
                        total = total + 1;
                        int curQty = 0;
                        string sPRNO = filterResults.Rows[i]["ROPRNO"].ToString();
                        string sWHLO = filterResults.Rows[i]["ROWHLO"].ToString();
                        string sPSTS = filterResults.Rows[i]["ROPSTS"].ToString();
                        string sFIDT = filterResults.Rows[i]["ROFIDT"].ToString();
                        string sRORC = filterResults.Rows[i]["RORORC"].ToString();
                        string sRORN = filterResults.Rows[i]["RORORN"].ToString();
                        string sRORL = filterResults.Rows[i]["RORORL"].ToString();
                        string sRORX = filterResults.Rows[i]["RORORX"].ToString();
                        string sORTY = filterResults.Rows[i]["ROORQT"].ToString();



                        DataRow[] dr = newMOPs.Select(string.Format("PRNO ='{0}' AND SCHN = '{1}' AND CURR = '{2}'", sPRNO, ScheduleNumber, curIndex));
                        //CreateLog(JobID, "fill", "dr count: " + dr.Count().ToString(), "", "", "FeltHat");


                        if (dr.Count() > 0)
                        {
                            if (dr[0]["PPQT"] != null)
                            {
                                curQty = int.Parse(dr[0]["PPQT"].ToString());
                            }
                            dr[0]["PPQT"] = (curQty + 1).ToString();
                            dr[0]["CURR"] = (curIndex + 1).ToString();

                            CreateLog(JobID, "fill", "SCHN: " + ScheduleNumber + " PRNO: " + sPRNO + " PPQT 1: " + (curQty + 1).ToString() + " CURR: " + curIndex.ToString(), "", "", "FeltHat");
                        }

                        else
                        {
                            DataRow newRow = newMOPs.NewRow();

                            newRow["CONO"] = company;
                            newRow["FACI"] = facility;
                            newRow["WHLO"] = sWHLO;
                            newRow["PRNO"] = sPRNO;
                            newRow["STRT"] = "001";
                            newRow["PPQT"] = (curQty + 1).ToString();
                            newRow["PLDT"] = sFIDT;
                            newRow["PLHM"] = "0000";
                            newRow["ORTY"] = "A01";
                            newRow["RORC"] = sRORC;
                            newRow["RORN"] = sRORN;
                            newRow["RORL"] = sRORL;
                            newRow["RORX"] = sRORX;
                            newRow["SCHN"] = ScheduleNumber;
                            newRow["SIMD"] = "0";
                            newRow["PSTS"] = sPSTS;
                            newRow["CURR"] = (curIndex + 1).ToString();

                            newMOPs.Rows.Add(newRow);

                            CreateLog(JobID, "fill", "SCHN: " + ScheduleNumber + " PRNO: " + sPRNO + " PPQT 2: " + (curQty + 1).ToString() + " CURR: " + curIndex.ToString(), "", "", "FeltHat");
                        }
                    }
                }
                curIndex = curIndex + 1;

            }

            //newMOPs group by PRNO
        }

        bool AreAllColumnsEmpty(DataRow dr)
        {
            if (dr == null)
            {
                return true;
            }
            else
            {
                foreach (var value in dr.ItemArray)
                {
                    if (value != null)
                    {
                        return false;
                    }
                }
                return true;
            }
        }

        public string AddScheduleNo(string StyleNumber, string JobID)
        {
            string Infobar = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS270MI/AddScheduleNo"));
            lstparms.Add(new Tuple<string, string>("SCHN", ""));
            lstparms.Add(new Tuple<string, string>("TX40", StyleNumber));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(JobID, sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                int count = rss["results"][0]["records"].Count();

                if (count > 0)
                {
                    string sSCHN = rss["results"][0]["records"][0]["SCHN"].ToString();
                    CreateLog(JobID, "New Schedules created", sSCHN, "", "", "FeltHat");
                    return sSCHN;
                }
            }

            return "";
        }

        public static string RemoveDuplicates(string input)
        {
            string output = string.Empty;
            string[] parts = input.Split(',');
            List<string> list = new List<string>();
            foreach (string part in parts)
            {
                if (!list.Contains(part))
                {
                    list.Add(part);
                }
            }
            output = string.Join(",", list);
            return output;
        }

        //public string CreatePlannedMO(DataRow dr)
        //{
        //    string Infobar = "";
        //    List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
        //    lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS170MI/CrtPlannedMO"));
        //    lstparms.Add(new Tuple<string, string>("CONO", dr["CONO"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("FACI", dr["FACI"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("WHLO", dr["WHLO"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("PRNO", dr["PRNO"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("STRT", dr["STRT"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("PSTS", dr["PSTS"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("PPQT", dr["PPQT"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("PLDT", dr["PLDT"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("PLHM", dr["PLHM"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("ORTY", dr["ORTY"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("RORC", dr["RORC"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("RORN", dr["RORN"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("RORL", dr["RORL"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("RORX", dr["RORX"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("SIMD", dr["SIMD"].ToString()));
        //    lstparms.Add(new Tuple<string, string>("SCHN", dr["SCHN"].ToString()));

        //    string sMethodList = BuildMethodParms(lstparms);
        //    string result = ProcessIONAPI(sMethodList, out Infobar);

        //    if (string.IsNullOrEmpty(Infobar))
        //    {
        //        JObject rss = JObject.Parse(result);

        //        int count = rss["results"][0]["records"].Count();

        //        if (count > 0)
        //        {
        //            string sPLPN = rss["results"][0]["records"][0]["PLPN"].ToString();
        //            return sPLPN;
        //        }
        //    }

        //    return "";
        //}

        public CrtMOPRecord CreatePlannedMO(DataRow dr)
        {
            CrtMOPRecord oCrtMOPRecord = new CrtMOPRecord();
            oCrtMOPRecord.CONO = int.Parse(dr["CONO"].ToString());
            oCrtMOPRecord.FACI = dr["FACI"].ToString();
            oCrtMOPRecord.WHLO = dr["WHLO"].ToString();
            oCrtMOPRecord.PRNO = dr["PRNO"].ToString();
            oCrtMOPRecord.STRT = dr["STRT"].ToString();
            oCrtMOPRecord.PSTS = dr["PSTS"].ToString();
            oCrtMOPRecord.PPQT = dr["PPQT"].ToString();
            oCrtMOPRecord.PLDT = dr["PLDT"].ToString();
            oCrtMOPRecord.PLHM = dr["PLHM"].ToString();
            oCrtMOPRecord.ORTY = dr["ORTY"].ToString();
            oCrtMOPRecord.RORC = dr["RORC"].ToString();
            oCrtMOPRecord.RORN = dr["RORN"].ToString();
            oCrtMOPRecord.RORL = dr["RORL"].ToString();
            oCrtMOPRecord.RORX = dr["RORX"].ToString();
            oCrtMOPRecord.SIMD = dr["SIMD"].ToString();
            oCrtMOPRecord.SCHN = dr["SCHN"].ToString();

            return oCrtMOPRecord;
        }

        public string GetUserData(string JobID, string user, ref string sMessage)
        {
            string Infobar = "", email = "";
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/MNS150MI/GetUserData"));
            lstparms.Add(new Tuple<string, string>("USID", user));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(JobID, sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                email = rss["results"][0]["records"][0]["NAME"].ToString();
            }
            else
            {
                throw new Exception(Infobar);
            }

            return email;


        }

        public string GetWorkstationID(ref string sMessage)
        {
            string Infobar = "", login = "";
            string filter = "Username = '" + IDORuntime.Context.UserName + "'";
            LoadCollectionResponseData loadresponse = this.Context.Commands.LoadCollection("MGCore.UserNames", "WorkstationLogin", filter, "", 1);

            if (loadresponse.Items.Count > 0)
            {
                login = loadresponse[0, "WorkstationLogin"].Value;
            }

            return login;


        }

        public void UpdateScheduleNo(string schduleNumber, string planOrderNumber, string JobID)
        {
            string Infobar = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS170MI/Updat"));
            lstparms.Add(new Tuple<string, string>("PLPN", planOrderNumber));
            lstparms.Add(new Tuple<string, string>("SCHN", schduleNumber));
            lstparms.Add(new Tuple<string, string>("IGWA", "1"));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(JobID, sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                int count = rss["results"][0]["records"].Count();
            }
        }

        public string GetSytle(string productNo, string JobID)
        {
            string sSytle = "";
            string Infobar = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/MMS200MI/Get"));
            lstparms.Add(new Tuple<string, string>("ITNO", productNo));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(JobID, sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                sSytle = rss["results"][0]["records"][0]["HDPR"].ToString();
                CreateLog(JobID, "GetSytle", sSytle, "", "", "FeltHat");
            }

            return sSytle;
        }

        public virtual string BuildMethodParms(List<Tuple<string, string>> lstInput)
        {
            string sFilter = "";

            for (int i = 0; i < lstInput.Count; i++)
            {
                if (String.IsNullOrEmpty(lstInput[i].Item2))
                    continue;

                if (i == 0)
                {
                    sFilter += lstInput[i].Item2 + "?";
                }
                else if (i == (lstInput.Count - 1))
                {
                    sFilter += lstInput[i].Item1 + "=" + lstInput[i].Item2;
                }
                else
                {
                    sFilter += lstInput[i].Item1 + "=" + lstInput[i].Item2 + "&";
                }
            }

            return sFilter;
        }

        public string ProcessIONAPI(string JobID, string sMethod, out string errMsg)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            string sso = "0";
            string serverId = "0";
            string suiteContext = "M3";
            string httpMethod = "GET";
            string methodName = sMethod;
            string parametes = "";
            string contentType = "application/xml";
            string timeout = "10000";
            string result = "";
            errMsg = "";

            InvokeRequestData IDORequest = new InvokeRequestData();
            InvokeResponseData response;
            IDORequest.IDOName = "IONAPIMethods";
            IDORequest.MethodName = "InvokeIONAPIMethod";
            IDORequest.Parameters.Add(sso);
            IDORequest.Parameters.Add(serverId);
            IDORequest.Parameters.Add(new InvokeParameter(suiteContext));
            IDORequest.Parameters.Add(new InvokeParameter(httpMethod));
            IDORequest.Parameters.Add(new InvokeParameter(methodName));
            IDORequest.Parameters.Add(new InvokeParameter(parametes));
            IDORequest.Parameters.Add(new InvokeParameter(contentType));
            IDORequest.Parameters.Add(new InvokeParameter(timeout));
            IDORequest.Parameters.Add(IDONull.Value);
            IDORequest.Parameters.Add(IDONull.Value);
            IDORequest.Parameters.Add(IDONull.Value);
            IDORequest.Parameters.Add(IDONull.Value);

            response = this.Context.Commands.Invoke(IDORequest);
            if (response.IsReturnValueStdError())
            {
                errMsg = response.Parameters[8].Value;
            }
            else
            {
                result = response.Parameters[9].Value;
                JObject rss = JObject.Parse(result);
                int err = int.Parse(rss["nrOfFailedTransactions"].ToString());

                if (err == 1)
                {
                    errMsg = rss["results"][0]["errorMessage"].ToString();
                }

            }
            if (errMsg.Trim() == "0")
            {
                errMsg = response.Parameters[11].Value;
            }

            stopwatch.Stop();
            TimeSpan ts = stopwatch.Elapsed;
            string sTimeTaken = (ts.Hours > 0 ? ts.Hours + ":" : "") + (ts.Minutes > 0 ? ts.Minutes + ":" : "") + ts.Seconds + "." + ts.Milliseconds;

            CreateLog(JobID, sMethod, result, errMsg, sTimeTaken, "FeltHat");


            return result;
        }

        public void ProcessAppEvent(string returnMessage, string errorMessage, string user, string scheduleNo)
        {
            returnMessage = string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage;
            string process = "Felt Hat";

            ApplicationEventParameter[] ParmList = new ApplicationEventParameter[5];
            string result;
            ParmList[0] = new ApplicationEventParameter();
            ParmList[0].Name = "varProcess";
            ParmList[0].Value = process;
            ParmList[1] = new ApplicationEventParameter();
            ParmList[1].Name = "varReturnMsg";
            ParmList[1].Value = returnMessage;
            ParmList[2] = new ApplicationEventParameter();
            ParmList[2].Name = "varErrorMsg";
            ParmList[2].Value = errorMessage;
            ParmList[3] = new ApplicationEventParameter();
            ParmList[3].Name = "varUser";
            ParmList[3].Value = user;
            ParmList[4] = new ApplicationEventParameter();
            ParmList[4].Name = "varOrigSchNo";
            ParmList[4].Value = scheduleNo;

            FireApplicationEvent("FTD_MOPSplit", true, true, out result, ref ParmList);
        }

        private void CreateLog(string JobID, string sRequest, string sResponse, string sErrorMessage, string sTimeTaken, string sModuleName)
        {
            sRequest = new string(sRequest.Take(4000).ToArray());
            sResponse = new string(sResponse.Take(4000).ToArray());
            sErrorMessage = new string(sErrorMessage.Take(4000).ToArray());

            UpdateCollectionRequestData updateRequest = new UpdateCollectionRequestData("FTD_M3ConnectionLogs");
            IDOUpdateItem updateItem = new IDOUpdateItem(UpdateAction.Insert);
            updateItem.Properties.Add("JobID", JobID);
            updateItem.Properties.Add("ModuleName", sModuleName);
            updateItem.Properties.Add("Request", sRequest);
            updateItem.Properties.Add("Response", sResponse);
            updateItem.Properties.Add("ErrorMessage", sErrorMessage);
            updateItem.Properties.Add("TimeTaken", sTimeTaken);
            updateRequest.Items.Add(updateItem);
            this.Context.Commands.UpdateCollection(updateRequest);
        }

        protected string GetBearerToken(string serverID, string SSO, ref string infobar)
        {
            string value;
            try
            {
                InvokeRequestData invokeRequestDatum = new InvokeRequestData();
                invokeRequestDatum.IDOName = "IONAPIMethods";
                invokeRequestDatum.MethodName = "GetBearerToken";
                invokeRequestDatum.Parameters.Add(new InvokeParameter(serverID));
                invokeRequestDatum.Parameters.Add(new InvokeParameter(SSO));
                InvokeParameter invokeParameter = new InvokeParameter();
                invokeParameter.ByRef = true;
                InvokeParameter invokeParameter1 = new InvokeParameter();
                invokeParameter1.ByRef = true;
                invokeRequestDatum.Parameters.Add(invokeParameter);
                invokeRequestDatum.Parameters.Add(invokeParameter1);
                InvokeResponseData invokeResponseDatum = this.Context.Commands.Invoke(invokeRequestDatum);
                infobar = invokeResponseDatum.Parameters[3].Value;
                value = invokeResponseDatum.Parameters[2].Value;
            }
            catch (Exception exception)
            {
                infobar = exception.ToString();
                value = null;
            }
            return value;
        }

        public string CallServicePost(string json, string JobID)
        {
            string empty = string.Empty;
            string token = GetBearerToken("0", "0", ref empty);
            string IONAPIBaseUrl = "https://mingle-ionapi.inforcloudsuite.com";
            var client = new HttpClient
            {
                BaseAddress = new Uri(IONAPIBaseUrl)
            };

            client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));

            var content = new StringContent(json, Encoding.UTF8, "application/json");



            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = client.PostAsync("/HATCO_PRD/M3/m3api-rest/v2/execute", content).Result;

            CreateLog(JobID, json, response.ToString(), response.StatusCode.ToString(), "", "FeltHat");

            if (response.IsSuccessStatusCode)
            {
                string result = response.Content.ReadAsStringAsync().Result;
                return result;
            }
            else
            {
                return string.Empty;
            }


        }

        public void RemoveOrigSchFromMOs(string company, string division, string JobID)
        {
            UpdateMORoot oUpdateMORoot = new UpdateMORoot();
            oUpdateMORoot.program = "PMS170MI";
            oUpdateMORoot.cono = int.Parse(company);
            oUpdateMORoot.divi = division;
            oUpdateMORoot.excludeEmptyValues = false;
            oUpdateMORoot.rightTrim = true;
            oUpdateMORoot.maxReturnedRecords = 0;

            List<UpdateMOTransaction> oListUpdateMOTransaction = new List<UpdateMOTransaction>();

            for (int i = 0; i < fullResults.Rows.Count; i++)
            {
                string sMOList = fullResults.Rows[i]["ROPLPN"].ToString();
                List<string> MOs = sMOList.Split(',').ToList();

                foreach (string sPLPN in MOs)
                {
                    UpdateMOTransaction oUpdateMOTransaction = new UpdateMOTransaction();
                    oUpdateMOTransaction.transaction = "Updat";
                    List<string> listSelectedColumns = new List<string>();
                    oUpdateMOTransaction.selectedColumns = listSelectedColumns;
                    UpdateMO oUpdateMO = new UpdateMO();
                    oUpdateMO.CONO = int.Parse(company);
                    oUpdateMO.PLPN = sPLPN;
                    oUpdateMO.RORC = " ";
                    oUpdateMO.RORN = " ";
                    oUpdateMO.RORL = " ";
                    oUpdateMO.DSP1 = "1";
                    oUpdateMO.DSP3 = "1";
                    oUpdateMO.DSP4 = "1";
                    oUpdateMO.DSP6 = "1";
                    oUpdateMO.DSP7 = "1";
                    oUpdateMO.DSP8 = "1";

                    oUpdateMOTransaction.record = oUpdateMO;
                    oListUpdateMOTransaction.Add(oUpdateMOTransaction);
                }

            }

            oUpdateMORoot.transactions = oListUpdateMOTransaction;

            string updjsonString = JsonConvert.SerializeObject(oUpdateMORoot);

            string updPresults = CallServicePost(updjsonString, JobID);

        }


        public void DelPlannedMOs(string company, string division, string JobID)
        {
            DltRoot oDltRoot = new DltRoot();
            oDltRoot.program = "PMS170MI";
            oDltRoot.cono = int.Parse(company);
            oDltRoot.divi = division;
            oDltRoot.excludeEmptyValues = false;
            oDltRoot.rightTrim = true;
            oDltRoot.maxReturnedRecords = 0;

            List<DltTransaction> oListDltTransaction = new List<DltTransaction>();

            for (int i = 0; i < fullResults.Rows.Count; i++)
            {
                string sMOList = fullResults.Rows[i]["ROPLPN"].ToString();
                List<string> MOs = sMOList.Split(',').ToList();

                foreach (string sPLPN in MOs)
                {
                    DltTransaction oDltTransaction = new DltTransaction();
                    oDltTransaction.transaction = "DelPlannedMO";
                    List<string> listSelectedColumns = new List<string>();
                    oDltTransaction.selectedColumns = listSelectedColumns;
                    DltRecord oDltRecord = new DltRecord();
                    oDltRecord.CONO = int.Parse(company);
                    oDltRecord.PLPN = sPLPN;
                    oDltTransaction.record = oDltRecord;
                    oListDltTransaction.Add(oDltTransaction);
                }

            }

            oDltRoot.transactions = oListDltTransaction;

            string dltjsonString = JsonConvert.SerializeObject(oDltRoot);

            string dltPresults = CallServicePost(dltjsonString, JobID);

        }

        public int CreateBGTask(string BGTask, string BGParms)
        {
            int newTaskID = 0;
            string empty = string.Empty;
            string TaskName = BGTask;
            string TaskParms1 = BGParms;
            string TaskParms2 = "";
            string Infobar = null;
            string TaskID = null;
            string TaskStatusCode = "READY";
            string StringTable = null;
            string RequestingUser = null;
            string PrintPreview = "0";
            string TaskHistoryRowPointer = null;
            string PreviewInterval = null;
            string SchedStartDateTime = null;
            string SchedEndDateTime = null;
            string SchedFreqType = "1";
            string SchedFreqInterval = "1";
            string SchedFreqSubDayType = "1";
            string SchedFreqSubDayInterval = "1";
            string SchedFreqRelativeInterval = "1";
            string SchedFreqRecurrenceFactor = "1";
            string SchedIsEnabled = "0";
            InvokeRequestData invokeRequestDatum = new InvokeRequestData();
            InvokeResponseData invokeResponseDatum = new InvokeResponseData();
            invokeRequestDatum.IDOName = "BGTaskDefinitions";
            invokeRequestDatum.MethodName = "BGTaskSubmit";
            invokeRequestDatum.Parameters.Add(TaskName);
            invokeRequestDatum.Parameters.Add(TaskParms1);
            invokeRequestDatum.Parameters.Add(TaskParms2);
            invokeRequestDatum.Parameters.Add(Infobar);
            invokeRequestDatum.Parameters.Add(TaskID);
            invokeRequestDatum.Parameters.Add(TaskStatusCode);
            invokeRequestDatum.Parameters.Add(StringTable);
            invokeRequestDatum.Parameters.Add(RequestingUser);
            invokeRequestDatum.Parameters.Add(PrintPreview);
            invokeRequestDatum.Parameters.Add(TaskHistoryRowPointer);
            invokeRequestDatum.Parameters.Add(PreviewInterval);
            invokeRequestDatum.Parameters.Add(SchedStartDateTime);
            invokeRequestDatum.Parameters.Add(SchedEndDateTime);
            invokeRequestDatum.Parameters.Add(SchedFreqType);
            invokeRequestDatum.Parameters.Add(SchedFreqInterval);
            invokeRequestDatum.Parameters.Add(SchedFreqSubDayType);
            invokeRequestDatum.Parameters.Add(SchedFreqSubDayInterval);
            invokeRequestDatum.Parameters.Add(SchedFreqRelativeInterval);
            invokeRequestDatum.Parameters.Add(SchedFreqRecurrenceFactor);
            invokeRequestDatum.Parameters.Add(SchedIsEnabled);
            try
            {
                invokeResponseDatum = this.Context.Commands.Invoke(invokeRequestDatum);
                invokeResponseDatum.Parameters[0].ToString();
                invokeResponseDatum.Parameters[1].ToString();
                invokeResponseDatum.Parameters[3].ToString();
                if (!string.IsNullOrWhiteSpace(invokeResponseDatum.Parameters[4].ToString()))
                {
                    string sTaskID = invokeResponseDatum.Parameters[4].ToString();
                    newTaskID = int.Parse(sTaskID);
                }
                empty = invokeResponseDatum.Parameters[3].ToString();
            }
            catch (Exception exception1)
            {
                Exception exception = exception1;
                empty = exception.InnerException.ToString();
                throw exception;
            }

            return newTaskID;
        }

        private bool BGTaskCompletedAndSuccess(int TaskID)
        {
            if (!CheckActiveBGTask(TaskID))
            {
                if (CheckBGTaskHistory(TaskID))
                    return true;
            }
            return false;
        }

        private bool CheckActiveBGTask(int TaskID)
        {
            //CreateLog(JobID, "CheckActiveBGTask", "", "", "", "FeltHat");
            LoadCollectionResponseData oResponse = this.Context.Commands.LoadCollection("ActiveBGTasks", "TaskStatusCode, TaskNumber", "TaskNumber = " + TaskID, "", 1);
            if (oResponse.Items.Count > 0)
            {
                string sStatus = oResponse[0, "TaskStatusCode"].Value.ToString();
                //CreateLog(JobID, "CheckActiveBGTask "+ sStatus, "", "", "", "FeltHat");
                if (sStatus == "WAITING" || sStatus == "READY" || sStatus == "RUNNING")
                    return true;
            }

            return false;
        }

        private bool CheckBGTaskHistory(int TaskID)
        {
            //CreateLog(JobID, "CheckBGTaskHistory", "", "", "", "FeltHat");
            LoadCollectionResponseData oResponse = this.Context.Commands.LoadCollection("BGTaskHistories", "CompletionStatus, TaskNumber, TaskErrorMsg", "TaskNumber = " + TaskID, "", 1);
            if (oResponse.Items.Count > 0)
            {
                string sStatus = oResponse[0, "CompletionStatus"].Value.ToString().Trim();
                //CreateLog(JobID, "CheckBGTaskHistory "+ sStatus, "", "", "", "FeltHat");
                if (sStatus == "")
                    return false;
                else if (sStatus == "0")
                {
                    return true;
                }
                else
                {
                    updateTasks = false;

                    errorMsg = oResponse[0, "TaskErrorMsg"].Value.ToString().Trim();
                    //CreateLog(JobID, "CheckBGTaskHistory " + errorMsg, "", "", "", "FeltHat");
                    return true;
                }
            }
            else
            {
                //CreateLog(JobID, "CheckBGTaskHistory", "Unable to find a Background task for :" + TaskID, "", "", "FeltHat");
                return true;
            }


        }

        private string CreateJobID()
        {
            LoadCollectionResponseData loadresponse;
            loadresponse = this.Context.Commands.LoadCollection("FTD_M3ConnectionLogs", "JobID", "", "JobID DESC", 1);
            int lastJobID = 0;
            if (loadresponse.Items.Count > 0)
            {
                lastJobID = int.Parse(loadresponse[0, "JobID"].Value);

            }
            string JobID = (lastJobID + 1).ToString();
            CreateLog(JobID, "GetJobID", "Get Job ID successfully - " + JobID, "", "", "FeltHat");
            return JobID;



        }

        private string GetAllSchedulNOsByJob(string JobID)
        {
            string sSchedulNOs = "", response = "";
            LoadCollectionResponseData loadresponse;
            loadresponse = this.Context.Commands.LoadCollection("FTD_M3ConnectionLogs", "JobID, Request, Response", "JobID = " + JobID + " AND Request = 'New Schedules created'", "JobID DESC", 0);
            if (loadresponse.Items.Count > 0)
            {
                for (int i = 0; i < loadresponse.Items.Count; i++)
                {
                    response = loadresponse[i, "Response"].Value.ToString().Trim();
                    if (response != "None")
                    {
                        sSchedulNOs = sSchedulNOs + "," + response;
                    }


                }

            }
            sSchedulNOs = String.Join(",", sSchedulNOs.TrimStart(',').Split(',').Select(x => int.Parse(x)).OrderBy(x => x));
            return sSchedulNOs;
        }
    }

    //public class RequestRecord
    //{
    //    public string TX40 { get; set; }
    //}

    //public class RequestRoot
    //{
    //    public string program { get; set; }
    //    public int cono { get; set; }
    //    public string divi { get; set; }
    //    public bool excludeEmptyValues { get; set; }
    //    public bool rightTrim { get; set; }
    //    public int maxReturnedRecords { get; set; }
    //    public List<RequestTransaction> transactions { get; set; }
    //}

    //public class RequestTransaction
    //{
    //    public string transaction { get; set; }
    //    public RequestRecord record { get; set; }
    //    public List<string> selectedColumns { get; set; }
    //}

    //public class ResponseRecord
    //{
    //    public string SCHN { get; set; }
    //}

    //public class ResponseResult
    //{
    //    public string transaction { get; set; }
    //    public List<ResponseRecord> records { get; set; }
    //}

    //public class ResponseRoot
    //{
    //    public List<ResponseResult> results { get; set; }
    //    public bool wasTerminated { get; set; }
    //    public int nrOfSuccessfullTransactions { get; set; }
    //    public int nrOfFailedTransactions { get; set; }
    //}

    //public class CrtMOPRecord
    //{
    //    public int CONO { get; set; }
    //    public string FACI { get; set; }
    //    public string WHLO { get; set; }
    //    public string PRNO { get; set; }
    //    public string STRT { get; set; }
    //    public string PSTS { get; set; }
    //    public string PPQT { get; set; }
    //    public string PLDT { get; set; }
    //    public string PLHM { get; set; }
    //    public string ORTY { get; set; }
    //    public string RORC { get; set; }
    //    public string RORN { get; set; }
    //    public string RORL { get; set; }
    //    public string RORX { get; set; }
    //    public string SIMD { get; set; }
    //    public string SCHN { get; set; }
    //}

    //public class CrtMOPRoot
    //{
    //    public string program { get; set; }
    //    public int cono { get; set; }
    //    public string divi { get; set; }
    //    public bool excludeEmptyValues { get; set; }
    //    public bool rightTrim { get; set; }
    //    public int maxReturnedRecords { get; set; }
    //    public List<CrtMOPTransaction> transactions { get; set; }
    //}

    //public class CrtMOPTransaction
    //{
    //    public string transaction { get; set; }
    //    public CrtMOPRecord record { get; set; }
    //    public List<string> selectedColumns { get; set; }
    //}


    //public class DltRecord
    //{
    //    public int CONO { get; set; }
    //    public string PLPN { get; set; }
    //}

    //public class DltRoot
    //{
    //    public string program { get; set; }
    //    public int cono { get; set; }
    //    public string divi { get; set; }
    //    public bool excludeEmptyValues { get; set; }
    //    public bool rightTrim { get; set; }
    //    public int maxReturnedRecords { get; set; }
    //    public List<DltTransaction> transactions { get; set; }
    //}

    //public class UpdateMOTransaction
    //{
    //    public string transaction { get; set; }
    //    public UpdateMO record { get; set; }
    //    public List<string> selectedColumns { get; set; }
    //}

    //public class UpdateMO
    //{
    //    public int CONO { get; set; }
    //    public string PLPN { get; set; }
    //    public string RORC { get; set; }
    //    public string RORN { get; set; }
    //    public string RORL { get; set; }
    //    public string DSP1 { get; set; }
    //    public string DSP3 { get; set; }
    //    public string DSP4 { get; set; }
    //    public string DSP6 { get; set; }
    //    public string DSP7 { get; set; }
    //    public string DSP8 { get; set; }
    //}

    //public class UpdateMORoot
    //{
    //    public string program { get; set; }
    //    public int cono { get; set; }
    //    public string divi { get; set; }
    //    public bool excludeEmptyValues { get; set; }
    //    public bool rightTrim { get; set; }
    //    public int maxReturnedRecords { get; set; }
    //    public List<UpdateMOTransaction> transactions { get; set; }
    //}

    //public class DltTransaction
    //{
    //    public string transaction { get; set; }
    //    public DltRecord record { get; set; }
    //    public List<string> selectedColumns { get; set; }
    //}

    //public class OperationsInfo
    //{
    //    public string Responsible { get; set; }
    //    public string StyleNumber { get; set; }
    //    public string Option { get; set; }
    //    public string PlannedQty { get; set; }
    //    public string ScheduleNo { get; set; }
    //    public string ProducNo { get; set; }
    //}

    //public class Model
    //{
    //    public string ROPLPN { get; set; }
    //    public string ROPLPS { get; set; }
    //    public string ROHDPR { get; set; }
    //    public string ROPRNO { get; set; }
    //    public int? ROORQA { get; set; }
    //    public string ROOPTX { get; set; }
    //    public string ROOPTY { get; set; }
    //    public string ROOPTZ { get; set; }
    //    public string ROPLDT { get; set; }
    //    public string ROFACI { get; set; }
    //    public string RORESP { get; set; }
    //    public string ROSTDT { get; set; }
    //    public string ROFIDT { get; set; }
    //    public string RORORC { get; set; }
    //    public string RORORN { get; set; }
    //    public string RORORL { get; set; }
    //    public string ROWHLO { get; set; }
    //    public string ROPSTS { get; set; }
    //    public string RORORX { get; set; }
    //    public int? ROORQT { get; set; }
    //    public string ROORTY { get; set; }
    //}
}
