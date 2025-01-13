using Mongoose.IDO;
using Mongoose.IDO.Protocol;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;

namespace FTD_MOSplit_StrawHat
{
    public class StrawHat : IDOExtensionClass
    {
        List<string> opLst = new List<string>();
        string cono = "", divi = "";
        string serverURL = "", tenantID = "", bearerToken = "";

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
        public void ProcessAggregatedData(string scheduleNo, string workCenter, string facility, string company, string division, string styleNumber, string user, ref string returnMessage)
        {
            string sMessage = "";
            returnMessage = "";
            string email = GetWorkstationID(ref returnMessage);//GetUserData(user, ref sMessage);
            cono = company; divi = division;

            try
            {
                //Extract Data
                List<OperationsInfo> operationsLst = LstMWOOPE04(scheduleNo, workCenter, workCenter, facility, "", true);

                if (string.IsNullOrEmpty(sMessage))
                {
                    performSplitting(operationsLst, scheduleNo, styleNumber, company, division, facility, workCenter, ref returnMessage);
                }
                else
                    throw new Exception(sMessage);

                ProcessAppEvent(returnMessage, "", email, scheduleNo);
                returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage);
            }

            catch (Exception ex)
            {
                ProcessAppEvent(returnMessage, ex.Message, email, scheduleNo);
                returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage) +
                     "<br>Errors: " + ex.Message;
            }


        }

        [IDOMethod(Flags = MethodFlags.None)]
        public void ProcessData(string workCenter, string parmData, string facility, string company, string division, string user, ref string returnMessage)
        {
            string sMessage = "", scheduleNo = "", styleNumber = "", operation = ""; int currentOperation = 0, completeStatus = 40; returnMessage = "";
            string email = GetWorkstationID(ref returnMessage);//GetUserData(user, ref sMessage);
            cono = company; divi = division;
            try
            {
                parmData = parmData.Replace("@#!", ",");
                var dataLst = JsonConvert.DeserializeObject<List<OperationsInfo>>(parmData);
                List<string> moLst = dataLst.Select(x => x.MONumber).ToList();
                currentOperation = int.Parse(dataLst.First().OPNumber);
                scheduleNo = dataLst.First().ScheduleNo;
                styleNumber = dataLst.First().StyleNumber;

                List<OperationsInfo> origOperationsLst = LstMWOOPE04(scheduleNo, "", "", facility, "", true);
                if (origOperationsLst.Count > 0)
                {
                    if (string.IsNullOrEmpty(sMessage))
                    {
                        //Validate max qty
                        var maxQtyOperationsLst = origOperationsLst.Where(x => x.MaxQty == 0);
                        if (maxQtyOperationsLst.Count() > 0)
                        {
                            throw new Exception("Max quantity is not available in " + maxQtyOperationsLst.First().MONumber);

                        }
                        else
                        {
                            //Check previous operations which are not completed
                            var workCenterStatus = origOperationsLst.Where(x => x.WorkCenter == workCenter).First().Status;

                            if (int.Parse(workCenterStatus) != completeStatus)
                            {
                                throw new Exception("All previous operations are not completed");

                            }
                            else
                            {
                                operation = origOperationsLst.First().OPNumber;
                                var operationsLst = origOperationsLst.Where(x => x.OPNumber.Equals(operation) && x.ScheduleNo.Equals(scheduleNo)
                                                                                    && moLst.Contains(x.MONumber));
                                performSplitting(operationsLst.ToList(), scheduleNo, styleNumber, company, division, facility, workCenter, ref returnMessage);
                            }

                        }


                    }
                    else
                        throw new Exception(sMessage);
                }

                ProcessAppEvent(returnMessage, "", email, scheduleNo);
                returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage);

            }
            catch (Exception ex)
            {
                ProcessAppEvent(returnMessage, ex.Message, email, scheduleNo);
                returnMessage = "New Schedules created: " + (string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage) +
                     "<br>Errors: " + ex.Message;
            }


        }

        private void performSplitting(List<OperationsInfo> operationsLst, string scheduleNo, string styleNumber, string company, string division, string facility, string workCenter, ref string returnMessage)
        {
            string sMessage = "", item = "", moNumber = "", newMONumber = "", opNumber = "", newScheduleNo = "", startDate = "", jobNumber = ""; returnMessage = "";
            int maxQty = 0, newQty = 0, qty = 0;

            List<UpdateSchInfo> updateOperationsLst = new List<UpdateSchInfo>();

            //Group by MO, Item and OP number
            List<OperationsInfo> splitList = operationsLst
                                                .GroupBy(g => new { g.MONumber, g.Item, g.OPNumber })
                                                .Select(s => new OperationsInfo
                                                {
                                                    MONumber = s.First().MONumber,
                                                    OPNumber = s.First().OPNumber,
                                                    ScheduleNo = s.First().ScheduleNo,
                                                    Item = s.First().Item,
                                                    Status = s.First().Status,
                                                    StartDate = s.First().StartDate,
                                                    StyleNumber = s.First().StyleNumber,
                                                    MaxQty = s.First().MaxQty,
                                                    Qty = s.Sum(q => q.Qty)
                                                })
                                                .OrderBy(i => i.Item)
                                                .ThenBy(o => o.OPNumber)
                                                .ThenBy((q => q.Qty))
                                                .ToList();


            while (splitList.Count > 0)
            {
                jobNumber = "";

                //Only keep MOs that need to be splitted
                splitList.RemoveAll(q => q.Qty == 0);

                for (int i = 0; i < splitList.Count; i++)
                {
                    qty = splitList[i].Qty;
                    item = splitList[i].Item;
                    moNumber = splitList[i].MONumber;
                    opNumber = splitList[i].OPNumber;
                    startDate = splitList[i].StartDate;
                    newMONumber = ""; newScheduleNo = "";
                    maxQty = splitList[i].MaxQty;


                    if (qty > maxQty)
                    {
                        newQty = qty - maxQty;

                        //Split MO
                        if (newQty > 0)
                        {
                            if (string.IsNullOrEmpty(jobNumber))
                                jobNumber = SplitMO(facility, item, moNumber, opNumber, startDate, newQty, jobNumber, ref newMONumber);
                            else
                                SplitMO(facility, item, moNumber, opNumber, startDate, newQty, jobNumber, ref newMONumber);
                        }

                        //Update new qty
                        splitList[i].Qty = newQty;

                        //Save new MO to the update list
                        UpdateSchInfo updOpObj = new UpdateSchInfo()
                        {
                            Item = item,
                            OPNumber = opNumber,
                            MaxQty = maxQty,
                            Qty = maxQty,
                            MONumber = newMONumber,
                            OrigMONumber = moNumber
                        };

                        updateOperationsLst.Add(updOpObj);
                    }
                    else
                    {

                        //Save MO to the update list
                        UpdateSchInfo updOpObj = new UpdateSchInfo()
                        {
                            Item = item,
                            OPNumber = opNumber,
                            MaxQty = maxQty,
                            Qty = qty,
                            MONumber = moNumber,
                            OrigMONumber = ""
                        };

                        updateOperationsLst.Add(updOpObj);

                        //MO is not spliited
                        splitList[i].Qty = 0;

                    }
                }

                //Update new MOs 
                for (int i = 0; i < splitList.Count; i++)
                {
                    //Trigger Job for the batch
                    if (i == 0 && !string.IsNullOrEmpty(jobNumber))
                    {
                        TriggerJob(jobNumber);
                    }

                    //Wait till splitting is done for each MO (Wait till the last digit of the MO Status is '0')
                    int wait = 0;
                    while (Get(facility, splitList[i].Item, splitList[i].MONumber) > 0 && wait < 15)
                    {
                        Thread.Sleep(5000);
                        wait++;
                    }
                }
            }

            //Now the splitting is completed, update MOs with new schedule numbers
            AssignScheduleNumbers(updateOperationsLst, styleNumber, company, division, facility, workCenter, scheduleNo, ref returnMessage);


        }

        public void AssignScheduleNumbers(List<UpdateSchInfo> updateOperationsLst, string styleNumber, string company, string division, string facility, string workCenter, string scheduleNo, ref string returnMessage)
        {
            string newScheduleNo = "", prevItem = "", prevOP = "", item = "", op = "";
            int sum = 0, qty = 0;

            List<UpdateSchInfo> orderedList = updateOperationsLst.OrderBy(i => i.Item)
                                                                .ThenBy(o => o.OPNumber)
                                                                .ThenBy((q => q.Qty))
                                                                .ToList();
            updateOperationsLst = null;

            //Group by Item-OP Number and assign Schedule Numbers
            for (int i = 0; i < orderedList.Count; i++)
            {
                item = orderedList[i].Item;
                op = orderedList[i].OPNumber;
                qty = orderedList[i].Qty;

                if (prevItem.Equals(item) && prevOP.Equals(op))
                {
                    sum = sum + qty;

                    if (sum > orderedList[i].MaxQty)
                    {
                        //Create new schedule
                        newScheduleNo = AddScheduleNo(styleNumber);
                        returnMessage = string.IsNullOrEmpty(returnMessage) ? newScheduleNo : (returnMessage + "," + newScheduleNo);

                        sum = qty;

                    }
                }
                else
                {
                    sum = 0;

                    prevItem = item;
                    prevOP = op;

                    //Create new schedule
                    newScheduleNo = AddScheduleNo(styleNumber);
                    returnMessage = string.IsNullOrEmpty(returnMessage) ? newScheduleNo : (returnMessage + "," + newScheduleNo);

                    sum = sum + qty;
                }

                //Update new schedule number in all new MO operations &&
                //Update original schedule number in all new MO operations
                UpdScheduleNum(company, division, facility, workCenter, orderedList[i].MONumber, item, scheduleNo, newScheduleNo);


            }

        }

        public List<OperationsInfo> LstMWOOPE04(string scheduleNo, string fromWorkCenter, string toWorkCenter, string facility, string moNumber, bool checkStatus)
        {
            List<OperationsInfo> operationsLst = new List<OperationsInfo>();
            string masterTicket = "", status = "", Infobar = "";
            int maxrecs = 5000;

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOOPE04"));
            lstparms.Add(new Tuple<string, string>("VOFACI", facility));
            lstparms.Add(new Tuple<string, string>("VOSCHN", scheduleNo));
            lstparms.Add(new Tuple<string, string>("F_PLGR", fromWorkCenter));
            lstparms.Add(new Tuple<string, string>("T_PLGR", toWorkCenter));
            lstparms.Add(new Tuple<string, string>("F_MFNO", moNumber));
            lstparms.Add(new Tuple<string, string>("T_MFNO", moNumber));
            lstparms.Add(new Tuple<string, string>("cono", cono));
            lstparms.Add(new Tuple<string, string>("divi", divi));
            lstparms.Add(new Tuple<string, string>("maxrecs", maxrecs.ToString()));
            string sMethodList = BuildMethodParms(lstparms);

            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                for (int i = 0; i < rss["results"][0]["records"].Count(); i++)
                {
                    masterTicket = rss["results"][0]["records"][i]["VOTXT1"].ToString();
                    status = rss["results"][0]["records"][i]["VOWOST"].ToString();

                    if (string.IsNullOrEmpty(masterTicket))
                    {
                        //for MO split - skip status 90, else - include status 90
                        if (!checkStatus || !status.Equals("90"))
                        {
                            OperationsInfo op = new OperationsInfo()
                            {
                                Facility = rss["results"][0]["records"][i]["VOFACI"].ToString(),
                                WorkCenter = rss["results"][0]["records"][i]["VOPLGR"].ToString(),
                                Qty = int.Parse(rss["results"][0]["records"][i]["VOORQT"].ToString()),
                                Item = rss["results"][0]["records"][i]["VOPRNO"].ToString(),
                                MONumber = rss["results"][0]["records"][i]["VOMFNO"].ToString(),
                                OPNumber = rss["results"][0]["records"][i]["VOOPNO"].ToString(),
                                ScheduleNo = rss["results"][0]["records"][i]["VOSCHN"].ToString(),
                                Status = rss["results"][0]["records"][i]["VOWOST"].ToString(),
                                StartDate = rss["results"][0]["records"][i]["VOSTDT"].ToString(),
                                MaxQty = Convert.ToInt32(Math.Round(Convert.ToDouble(rss["results"][0]["records"][i]["MMCFI2"].ToString())))
                            };
                            operationsLst.Add(op);
                        }

                    }

                }

            }
            else
            {
                throw new Exception(Infobar);
            }

            return operationsLst;

        }

        public string AddScheduleNo(string styleNumber)
        {
            string Infobar = "";
            string scheduleNo = "";
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS270MI/AddScheduleNo"));
            lstparms.Add(new Tuple<string, string>("SCHN", ""));
            lstparms.Add(new Tuple<string, string>("TX40", styleNumber));
            lstparms.Add(new Tuple<string, string>("cono", cono));
            lstparms.Add(new Tuple<string, string>("divi", divi));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                scheduleNo = rss["results"][0]["records"][0]["SCHN"].ToString();

            }
            else
            {
                throw new Exception(Infobar);
            }

            return scheduleNo;
        }

        public string SplitMO(string facility, string item, string moNumber, string opNumber, string origOrderStartdate, int splitQty, string jobNumber, ref string newMONumber)
        {
            string Infobar = "", newJobNumber = "";
            newMONumber = "";
            string newOrderStartDate = DateTime.Now.ToString("yyyyMMdd");
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS100MI/SplitMO"));
            lstparms.Add(new Tuple<string, string>("FACI", facility));
            lstparms.Add(new Tuple<string, string>("PRNO", item));
            lstparms.Add(new Tuple<string, string>("MFNO", moNumber));
            lstparms.Add(new Tuple<string, string>("OPNO", opNumber));
            lstparms.Add(new Tuple<string, string>("OSTD", origOrderStartdate));
            lstparms.Add(new Tuple<string, string>("NSTD", newOrderStartDate));
            lstparms.Add(new Tuple<string, string>("KORQ", splitQty.ToString()));
            lstparms.Add(new Tuple<string, string>("BJNO", jobNumber));
            lstparms.Add(new Tuple<string, string>("cono", cono));
            lstparms.Add(new Tuple<string, string>("divi", divi));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                newJobNumber = rss["results"][0]["records"][0]["BJNO"].ToString();
                newMONumber = rss["results"][0]["records"][0]["MFNO"].ToString();
            }
            else
            {
                Infobar = "MO " + moNumber + "  : " + Infobar;
                throw new Exception(Infobar);
            }

            return newJobNumber;


        }

        public void TriggerJob(string jobNumber)
        {
            string Infobar = "";
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS100MI/TriggerJob"));
            lstparms.Add(new Tuple<string, string>("TRXN", "SplitMO"));
            lstparms.Add(new Tuple<string, string>("BJNO", jobNumber));
            lstparms.Add(new Tuple<string, string>("cono", cono));
            lstparms.Add(new Tuple<string, string>("divi", divi));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (!string.IsNullOrEmpty(Infobar))
            {
                throw new Exception(Infobar);
            }

        }

        public int Get(string facility, string item, string moNumber)
        {
            string Infobar = "";
            string status = ""; int lastDigit = 1;
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS100MI/Get"));
            lstparms.Add(new Tuple<string, string>("FACI", facility));
            lstparms.Add(new Tuple<string, string>("PRNO", item));
            lstparms.Add(new Tuple<string, string>("MFNO", moNumber));
            lstparms.Add(new Tuple<string, string>("cono", cono));
            lstparms.Add(new Tuple<string, string>("divi", divi));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                status = rss["results"][0]["records"][0]["WHST"].ToString();
                status = !string.IsNullOrEmpty(status) ? status[status.Length - 1].ToString() : "1";
                int.TryParse(status, out lastDigit);

            }
            else
            {
                Infobar = "MO " + moNumber + "  : " + Infobar;
                throw new Exception(Infobar);
            }

            return lastDigit;


        }

        public void UpdScheduleNum(string company, string division, string facility, string workCenter, string moNumber, string productNo, string scheduleNo, string newScheduleNo)
        {
            //Get list of operation numbers
            if (opLst.Count == 0)
            {
                List<OperationsInfo> operationsUpdLst = LstMWOOPE04(scheduleNo, "", "", facility, moNumber, false);
                opLst = operationsUpdLst.Select(x => x.OPNumber).Distinct().OrderBy(i => i).ToList();
            }

            //Update MO header, operations and materials with new schedule No
            string Infobar = "";
            UpdateMO(company, facility, moNumber, productNo, newScheduleNo, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {

                //Update MO with original schedule number
                UpdateOperation(company, division, facility, moNumber, productNo, scheduleNo);

            }
            else
            {
                Infobar = "MO " + moNumber + "  : " + Infobar;
                throw new Exception(Infobar);
            }

        }

        public void UpdateMO(string company, string facility, string moNumber, string productNo, string newScheduleNo, out string Infobar)
        {
            //Update MO header, materials and operations
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/EXT007MI/UpdSchedule"));
            lstparms.Add(new Tuple<string, string>("CONO", company));
            lstparms.Add(new Tuple<string, string>("FACI", facility));
            lstparms.Add(new Tuple<string, string>("MFNO", moNumber));
            lstparms.Add(new Tuple<string, string>("PRNO", productNo));
            lstparms.Add(new Tuple<string, string>("SCHN", newScheduleNo));
            lstparms.Add(new Tuple<string, string>("cono", cono));
            lstparms.Add(new Tuple<string, string>("divi", divi));


            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (!string.IsNullOrEmpty(Infobar))
            {
                Infobar = "MO " + moNumber + "  : " + Infobar;
                throw new Exception(Infobar);

            }

        }

        public void UpdateOperation(string company, string division, string facility, string moNumber, string productNo, string scheduleNo)
        {
            UpdOperationRoot oUpdOperationRoot = new UpdOperationRoot();
            oUpdOperationRoot.program = "PMS100MI";
            oUpdOperationRoot.cono = int.Parse(company);
            oUpdOperationRoot.divi = division;
            oUpdOperationRoot.excludeEmptyValues = false;
            oUpdOperationRoot.rightTrim = true;
            oUpdOperationRoot.maxReturnedRecords = 0;

            List<UpdOperationTransaction> oListUpdOperationTransaction = new List<UpdOperationTransaction>();

            for (int i = 0; i < opLst.Count; i++)
            {
                UpdOperationTransaction oUpdOperationTransaction = new UpdOperationTransaction();
                oUpdOperationTransaction.transaction = "UpdOperation";

                //output fields
                List<string> listSelectedColumns = new List<string>();
                oUpdOperationTransaction.selectedColumns = listSelectedColumns;

                //input fields
                UpdOperationRecord oUpdOperationRecord = new UpdOperationRecord();
                oUpdOperationRecord.CONO = int.Parse(company);
                oUpdOperationRecord.FACI = facility;
                oUpdOperationRecord.MFNO = moNumber;
                oUpdOperationRecord.PRNO = productNo;
                oUpdOperationRecord.OPNO = opLst[i];
                oUpdOperationRecord.TXT1 = scheduleNo;

                oUpdOperationTransaction.record = oUpdOperationRecord;
                oListUpdOperationTransaction.Add(oUpdOperationTransaction);
            }

            oUpdOperationRoot.transactions = oListUpdOperationTransaction;
            string UpdOperationjsonString = JsonConvert.SerializeObject(oUpdOperationRoot);
            string UpdOperationresults = CallServicePost(UpdOperationjsonString);
        }


        public string GetUserData(string user, ref string sMessage)
        {
            string Infobar = "", email = "";
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/MNS150MI/GetUserData"));
            lstparms.Add(new Tuple<string, string>("USID", user));
            lstparms.Add(new Tuple<string, string>("cono", cono));
            lstparms.Add(new Tuple<string, string>("divi", divi));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

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





        #region Common

        public string CallServicePost(string json)
        {
            string empty = string.Empty;

            this.bearerToken = string.IsNullOrEmpty(this.bearerToken) ? GetBearerToken("0", "0", ref empty) : this.bearerToken;
            this.serverURL = string.IsNullOrEmpty(this.serverURL) ? GetIONAPIInfo("0", ref empty) : this.serverURL;

            var client = new HttpClient
            {
                BaseAddress = new Uri(this.serverURL)
            };

            client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));

            var content = new StringContent(json, Encoding.UTF8, "application/json");

            string url = "/" + this.tenantID + "/M3/m3api-rest/v2/execute";

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", this.bearerToken);
            var response = client.PostAsync(url, content).Result;

            CreateLog(json, response.ToString(), response.StatusCode.ToString(), "", "StrawHat");


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

        protected string GetIONAPIInfo(string serverID, ref string infobar)
        {
            string value;
            try
            {
                InvokeRequestData invokeRequestDatum = new InvokeRequestData();
                invokeRequestDatum.IDOName = "IONAPIMethods";
                invokeRequestDatum.MethodName = "GetIONAPIInfo";
                invokeRequestDatum.Parameters.Add(new InvokeParameter(serverID)); //ServerID
                invokeRequestDatum.Parameters.Add(IDONull.Value); //Response Server URL             
                invokeRequestDatum.Parameters.Add(IDONull.Value); //Response Tenant ID
                invokeRequestDatum.Parameters.Add(IDONull.Value); //Reponse Infobar
                InvokeResponseData invokeResponseDatum = this.Context.Commands.Invoke(invokeRequestDatum);

                infobar = invokeResponseDatum.Parameters[3].Value;
                this.tenantID = invokeResponseDatum.Parameters[2].Value;

                value = invokeResponseDatum.Parameters[1].Value;


            }
            catch (Exception exception)
            {
                infobar = exception.ToString();
                value = null;
            }
            return value;
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


        public string ProcessIONAPI(string sMethod, out string errMsg)
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
            string sTimeTaken = (ts.Hours > 0 ? ts.Hours + ":" : "") + (ts.Minutes > 0 ? ts.Minutes + ":" : "") + ts.Seconds + (ts.Seconds > 0 ? "." : "") + ts.Milliseconds;

            CreateLog(sMethod, result, errMsg, sTimeTaken, "StrawHat");


            return result;
        }

        public void ProcessAppEvent(string returnMessage, string errorMessage, string user, string scheduleNo)
        {
            returnMessage = string.IsNullOrEmpty(returnMessage) ? "None" : returnMessage;
            string process = "Straw Hat";

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


        private void CreateLog(string sRequest, string sResponse, string sErrorMessage, string sTimeTaken, string sModuleName)
        {
            sRequest = new string(sRequest.Take(4000).ToArray());
            sResponse = new string(sResponse.Take(4000).ToArray());
            sErrorMessage = new string(sErrorMessage.Take(4000).ToArray());

            UpdateCollectionRequestData updateRequest = new UpdateCollectionRequestData("FTD_M3ConnectionLogs");
            IDOUpdateItem updateItem = new IDOUpdateItem(UpdateAction.Insert);
            updateItem.Properties.Add("ModuleName", sModuleName);
            updateItem.Properties.Add("Request", sRequest);
            updateItem.Properties.Add("Response", sResponse);
            updateItem.Properties.Add("ErrorMessage", sErrorMessage);
            updateItem.Properties.Add("TimeTaken", sTimeTaken);
            updateRequest.Items.Add(updateItem);
            this.Context.Commands.UpdateCollection(updateRequest);
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

        #endregion


    }


    public class OperationsInfo
    {
        public string Facility { get; set; }
        public string WorkCenter { get; set; }
        public int Qty { get; set; }
        public string MONumber { get; set; }
        public string OPNumber { get; set; }
        public string ScheduleNo { get; set; }
        public string Item { get; set; }
        public string Status { get; set; }
        public string StartDate { get; set; }
        public string StyleNumber { get; set; }
        public int MaxQty { get; set; }
    }

    public class UpdateSchInfo
    {
        public string ScheduleNo { get; set; }
        public string Item { get; set; }
        public string MONumber { get; set; }
        public string OrigMONumber { get; set; }
        public int Qty { get; set; }
        public string OPNumber { get; set; }
        public int MaxQty { get; set; }

    }

    #region Bulk Update DTOs

    public class MWOOPERoot
    {
        public string program { get; set; }
        public int cono { get; set; }
        public string divi { get; set; }
        public bool excludeEmptyValues { get; set; }
        public bool rightTrim { get; set; }
        public int maxReturnedRecords { get; set; }
        public List<MWOOPETransaction> transactions { get; set; }
    }

    public class MWOOPETransaction
    {
        public string transaction { get; set; }
        public MWOOPERecord record { get; set; }
        public List<string> selectedColumns { get; set; }
    }

    public class MWOOPERecord
    {
        public int CONO { get; set; }
        public string FACI { get; set; }
        public string MFNO { get; set; }
        public string PRNO { get; set; }
        public string OPNO { get; set; }
        public string SCHN { get; set; }

    }


    public class UpdOperationRoot
    {
        public string program { get; set; }
        public int cono { get; set; }
        public string divi { get; set; }
        public bool excludeEmptyValues { get; set; }
        public bool rightTrim { get; set; }
        public int maxReturnedRecords { get; set; }
        public List<UpdOperationTransaction> transactions { get; set; }
    }

    public class UpdOperationTransaction
    {
        public string transaction { get; set; }
        public UpdOperationRecord record { get; set; }
        public List<string> selectedColumns { get; set; }
    }

    public class UpdOperationRecord
    {
        public int CONO { get; set; }
        public string FACI { get; set; }
        public string MFNO { get; set; }
        public string PRNO { get; set; }
        public string OPNO { get; set; }
        public string TXT1 { get; set; }

    }

    #endregion

}

