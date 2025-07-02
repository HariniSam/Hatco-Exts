using Mongoose.IDO;
using Mongoose.IDO.Protocol;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

namespace FTD_MOR
{
    public class MOR : IDOExtensionClass
    {
        public override void SetContext(IIDOExtensionClassContext context)
        {
            base.SetContext(context);
            Context.IDO.PostLoadCollection += new IDOEventHandler(IDO_PostLoadCollection);
        }

        void IDO_PostLoadCollection(object sender, IDOEventArgs args)
        {
            LoadCollectionResponseData responseData = args.ResponsePayload as LoadCollectionResponseData;
        }

        [IDOMethod(Flags = MethodFlags.CustomLoad)]
        public IDataReader LoadScheduleNumbers(string scheduleNo)
        {
            DataTable ScheduleList = new DataTable("ScheduleNos");
            ScheduleList.Columns.Add("SCHN");
            int maxrecs = 1000;
            string sfilter = ""; string origScheduleNo = "";

            //For F2 search
            if (!string.IsNullOrEmpty(scheduleNo) && scheduleNo.Contains("%"))
            {
                origScheduleNo = scheduleNo.Replace("%", "");
                sfilter = "SCHN LIKE '" + scheduleNo + "'";

            }
            else
            {
                origScheduleNo = scheduleNo;
            }



            //Get Schedule Nos
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS270MI/LstScheduleNo"));
            lstparms.Add(new Tuple<string, string>("SCHN", origScheduleNo));
            lstparms.Add(new Tuple<string, string>("maxrecs", maxrecs.ToString()));

            string sMethod = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethod);
            RootScheduleNo resultObj = JsonConvert.DeserializeObject<RootScheduleNo>(result);

            for (int i = 0; i < resultObj.results[0].records.Count; i++)
            {
                DataRow newRow = ScheduleList.NewRow();
                newRow["SCHN"] = resultObj.results[0].records[i].SCHN;
                ScheduleList.Rows.Add(newRow);
            }

            ScheduleList.DefaultView.RowFilter = sfilter;
            ScheduleList = ScheduleList.DefaultView.ToTable();

            DataView view = new DataView(ScheduleList);
            DataTable distinctValues = view.ToTable(true, "SCHN");
            return distinctValues.CreateDataReader();
        }

        [IDOMethod(Flags = MethodFlags.CustomLoad)]
        public IDataReader LoadWorkCenters(string scheduleNo, string facility, string workCenter)
        {
            DataTable WorkCenterList = new DataTable("WorkCenters");
            WorkCenterList.Columns.Add("VOPLGR");
            int maxrecs = 1000;
            string sfilter = "";

            //For F2 search
            if (!string.IsNullOrEmpty(workCenter) && workCenter.Contains("%"))
            {
                sfilter = "VOPLGR LIKE '" + workCenter + "'";
            }

            //Get Work Centers
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOOPE01"));
            lstparms.Add(new Tuple<string, string>("VOFACI", facility));
            lstparms.Add(new Tuple<string, string>("VOSCHN", scheduleNo));
            lstparms.Add(new Tuple<string, string>("maxrecs", maxrecs.ToString()));

            string sMethod = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethod);
            RootWorkCenter resultObj = JsonConvert.DeserializeObject<RootWorkCenter>(result);

            for (int i = 0; i < resultObj.results[0].records.Count; i++)
            {
                DataRow newRow = WorkCenterList.NewRow();
                newRow["VOPLGR"] = resultObj.results[0].records[i].VOPLGR;
                WorkCenterList.Rows.Add(newRow);
            }

            WorkCenterList.DefaultView.RowFilter = sfilter;
            WorkCenterList = WorkCenterList.DefaultView.ToTable();

            DataView view = new DataView(WorkCenterList);
            DataTable distinctValues = view.ToTable(true, "VOPLGR");
            return distinctValues.CreateDataReader();
        }


        [IDOMethod(Flags = MethodFlags.None)]
        public void LoadOperations(string facility, string scheduleNo, string workCenter, string m3User)
        {
            DeleteOperations(m3User);
            try
            {
                string manQty, ordQty, startDate, status;
                int manufactQty, orderQty, remainQty;

                int num = 1000;

                List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
                lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOOPE02"));
                lstparms.Add(new Tuple<string, string>("VOFACI", facility));
                lstparms.Add(new Tuple<string, string>("VOSCHN", scheduleNo));
                if (!string.IsNullOrEmpty(workCenter))
                {
                    lstparms.Add(new Tuple<string, string>("F_PLGR", workCenter));
                    lstparms.Add(new Tuple<string, string>("T_PLGR", workCenter));
                }
                lstparms.Add(new Tuple<string, string>("maxrecs", num.ToString()));


                string parms = this.BuildMethodParms(lstparms);
                string result = this.ProcessIONAPI(parms);
                RootOperation operation = JsonConvert.DeserializeObject<RootOperation>(result);

                UpdateCollectionRequestData updateRequest = new UpdateCollectionRequestData("FTD_MOR_Operations");

                for (int i = 0; i < operation.results[0].records.Count; i++)
                {
                    IDOUpdateItem updateItem = new IDOUpdateItem(UpdateAction.Insert);
                    updateItem.Properties.Add("M3User", m3User);
                    updateItem.Properties.Add("OperationNo", int.Parse(operation.results[0].records[i].VOOPNO));
                    updateItem.Properties.Add("WorkCenter", operation.results[0].records[i].VOPLGR);
                    updateItem.Properties.Add("Facility", operation.results[0].records[i].VOFACI);
                    updateItem.Properties.Add("ScheduleNo", operation.results[0].records[i].VOSCHN);

                    startDate = operation.results[0].records[i].VORSDT;
                    if (!string.IsNullOrEmpty(startDate) && !startDate.Equals("*"))
                        startDate = DateTime.ParseExact(startDate, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy'/'MM'/'dd");
                    updateItem.Properties.Add("StartDate", startDate);

                    updateItem.Properties.Add("UoM", operation.results[0].records[i].MMUNMS);
                    updateItem.Properties.Add("RunTime", operation.results[0].records[i].VOPITI);

                    status = operation.results[0].records[i].VOWOST;
                    updateItem.Properties.Add("Status", status);

                    updateItem.Properties.Add("ProductNo", operation.results[0].records[i].VOPRNO);
                    updateItem.Properties.Add("ManufactNo", operation.results[0].records[i].VOMFNO);
                    updateItem.Properties.Add("OperationName", operation.results[0].records[i].MMITDS);
                    updateItem.Properties.Add("WorkCenterDesc", operation.results[0].records[i].VOOPDS);
                    updateItem.Properties.Add("ScrapQty", operation.results[0].records[i].VOSCQA);


                    manQty = operation.results[0].records[i].VOMAQT;
                    ordQty = operation.results[0].records[i].VOORQT;
                    manufactQty = string.IsNullOrEmpty(manQty) ? 0 : int.Parse(manQty);
                    orderQty = string.IsNullOrEmpty(ordQty) ? 0 : int.Parse(ordQty);
                    remainQty = orderQty - manufactQty;

                    updateItem.Properties.Add("ManufactQty", manufactQty);
                    updateItem.Properties.Add("OrderQty", orderQty);


                    if (status.Equals("90"))
                    {
                        updateItem.Properties.Add("RemainingQty", 0);
                        updateItem.Properties.Add("Qty", 0);
                    }
                    else
                    {
                        updateItem.Properties.Add("RemainingQty", remainQty);
                        updateItem.Properties.Add("Qty", remainQty);
                    }


                    updateRequest.Items.Add(updateItem);

                    if (i % 100 == 0 && i > 0)
                    {
                        //update IDO
                        this.Context.Commands.UpdateCollection(updateRequest);
                        updateRequest = null;

                        //create new request
                        updateRequest = new UpdateCollectionRequestData("FTD_MOR_Operations");
                    }

                }

                UpdateCollectionResponseData response = this.Context.Commands.UpdateCollection(updateRequest);
            }
            catch (Exception ex)
            {
                CreateLog("LoadOperations", ex.StackTrace, ex.Message, "", "Operation Reporting");
                throw new Exception(ex.Message);
            }
        }

        [IDOMethod(Flags = MethodFlags.None)]
        public void DeleteOperations(string m3User)
        {
            try
            {
                string sfilter = "M3User = '" + m3User + "'";
                LoadCollectionResponseData loadresponse = this.Context.Commands.LoadCollection("FTD_MOR_Operations", "OperationNo,ProductNo,ManufactNo", sfilter, "", 0);

                UpdateCollectionRequestData deleteRequest = new UpdateCollectionRequestData("FTD_MOR_Operations");
                for (int i = 0; i < loadresponse.Items.Count; i++)
                {
                    IDOUpdateItem deleteItem = new IDOUpdateItem(UpdateAction.Delete, loadresponse.Items[i].ItemID);
                    deleteRequest.Items.Add(deleteItem);

                    if (i % 100 == 0 && i > 0)
                    {
                        //update IDO
                        this.Context.Commands.UpdateCollection(deleteRequest);
                        deleteRequest = null;

                        //create new request
                        deleteRequest = new UpdateCollectionRequestData("FTD_MOR_Operations");
                    }

                }

                UpdateCollectionResponseData response = this.Context.Commands.UpdateCollection(deleteRequest);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [IDOMethod(Flags = MethodFlags.None)]
        public void ConfirmOperations(string facility, string scheduleNo, string workCenter, string m3User, string company, string date)
        {

            try
            {
                string sfilter = "M3User = '" + m3User + "' AND Facility = '" + facility + "' AND ScheduleNo = '" + scheduleNo + "'";
                if (!string.IsNullOrEmpty(workCenter))
                    sfilter = sfilter + " AND WorkCenter = '" + workCenter + "'";
                LoadCollectionResponseData loadresponse = this.Context.Commands.LoadCollection("FTD_MOR_Operations", "OperationNo,ProductNo,ManufactNo,Qty", sfilter, "ProductNo", 0);

                //Get detail level of operations
                string result = GetOperationDetails(facility, scheduleNo, workCenter);
                RootOperation operation = JsonConvert.DeserializeObject<RootOperation>(result);

                string productNo = string.Empty, opNo = string.Empty, manufactNo = string.Empty, status = string.Empty, message = string.Empty;
                int qty = 0, remainQty = 0;
                double priceTimeQty = 0, runTime = 0, usedLabourRuntime = 0;
                UpdateCollectionRequestData updateRequest = new UpdateCollectionRequestData("FTD_MOR_Operations");

                for (int i = 0; i < loadresponse.Items.Count; i++)
                {

                    opNo = loadresponse[i, "OperationNo"].Value;
                    qty = int.Parse(loadresponse[i, "Qty"].Value);
                    message = string.Empty;

                    var itemList = operation.results[0].records.Where(x => x.VOOPNO.Equals(opNo)).OrderBy(y => y.VOOPNO);

                    //Report Operation
                    foreach (var item in itemList)
                    {
                        productNo = item.VOPRNO;
                        manufactNo = item.VOMFNO;
                        remainQty = int.Parse(item.V_REMA);
                        priceTimeQty = string.IsNullOrEmpty(item.VOCTCD) ? 0 : double.Parse(item.VOCTCD);
                        runTime = string.IsNullOrEmpty(item.VOPITI) ? 0 : double.Parse(item.VOPITI);


                        if (qty > 0 && qty >= remainQty)
                        {
                            qty = qty - remainQty;
                            usedLabourRuntime = (priceTimeQty == 0) ? (runTime * remainQty) / 60 : (runTime * remainQty) / (priceTimeQty * 60);
                            usedLabourRuntime = Math.Round(usedLabourRuntime, 2);

                            message = message + RptOperation(company, facility, productNo, manufactNo, opNo, date, remainQty, usedLabourRuntime);
                        }

                        else if (qty < remainQty)
                        {
                            remainQty = qty;
                            qty = qty - remainQty;
                            usedLabourRuntime = (priceTimeQty == 0) ? (runTime * remainQty) / 60 : (runTime * remainQty) / (priceTimeQty * 60);
                            usedLabourRuntime = Math.Round(usedLabourRuntime, 2);

                            message = message + RptOperation(company, facility, productNo, manufactNo, opNo, date, remainQty, usedLabourRuntime);

                        }
                        else
                            break;


                    }

                    //Update results                    
                    status = string.IsNullOrEmpty(message) ? "Success" : "Fail";

                    IDOUpdateItem updateItem = new IDOUpdateItem(UpdateAction.Update, loadresponse.Items[i].ItemID);
                    updateItem.Properties.Add("ResultStatus", status);
                    updateItem.Properties.Add("ResultMessage", message);
                    updateRequest.Items.Add(updateItem);
                }

                this.Context.Commands.UpdateCollection(updateRequest);


            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private string RptOperation(string company, string facility, string productNo, string manufactNo, string operationNo, string date, int qty, double usedLabourRuntime)
        {
            string message = string.Empty;
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>()
            {
                new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS070MI/RptOperation"),
                new Tuple<string, string>("CONO", company),
                new Tuple<string, string>("FACI", facility),
                new Tuple<string, string>("PRNO", productNo),
                new Tuple<string, string>("MFNO", manufactNo),
                new Tuple<string, string>("OPNO", operationNo),
                new Tuple<string, string>("RPDT", DateTime.ParseExact(date, "yyyy'/'MM'/'dd", CultureInfo.InvariantCulture).ToString("yyyyMMdd")),
                new Tuple<string, string>("MAQA", qty.ToString()),
                new Tuple<string, string>("UMAT", usedLabourRuntime.ToString()),
                new Tuple<string, string>("SCQA", "0"),
                new Tuple<string, string>("REND", "1"),
                new Tuple<string, string>("DSP1", "1"),
                new Tuple<string, string>("DSP2", "1"),
                new Tuple<string, string>("DSP3", "1"),
                new Tuple<string, string>("DSP4", "1"),
            };

            string parms = this.BuildMethodParms(lstparms);
            string result = this.ProcessIONAPI(parms);
            if (!string.IsNullOrEmpty(result))
            {
                var resultObj = JsonConvert.DeserializeObject<RootValid>(result);

                if (resultObj.nrOfFailedTransactions > 0)
                {
                    message = "Manufacture No " + manufactNo + " : " + resultObj.results[0].errorMessage + "| ";
                }
            }
            else
            {
                message = "Manufacture No " + manufactNo + " : " + "Empty response" + "| ";

            }

            return message;

        }



        private string GetOperationDetails(string facility, string scheduleNo, string workCenter)
        {
            int num = 1000;
            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>()
            {
                new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOOPE01"),
                new Tuple<string, string>("VOFACI", facility),
                new Tuple<string, string>("VOSCHN", scheduleNo),
                new Tuple<string, string>("F_PLGR", workCenter),
                new Tuple<string, string>("T_PLGR", workCenter),
                new Tuple<string, string>("maxrecs", num.ToString())
            };

            string parms = this.BuildMethodParms(lstparms);
            string result = this.ProcessIONAPI(parms);
            return result;
        }


        public string ProcessIONAPI(string sMethod)
        {
            string sso = "1";
            string serverId = "0";
            string suiteContext = "M3";
            string httpMethod = "GET";
            string methodName = sMethod;
            string parametes = "";
            string contentType = "application/xml";
            string timeout = "10000";

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
            string result = response.Parameters[9].Value;

            CreateLog(sMethod, result, "", "", "Operation Reporting");


            return result;
        }

        public string BuildMethodParms(List<Tuple<string, string>> lstInput)
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



    }

    #region DTOs
    public class RootOperation
    {
        public List<ResultOperation> results { get; set; }
        public bool wasTerminated { get; set; }
        public int nrOfSuccessfullTransactions { get; set; }
        public int nrOfFailedTransactions { get; set; }
    }

    public class ResultOperation
    {
        public string transaction { get; set; }
        public List<Operation> records { get; set; }
    }

    public class Operation
    {
        public string VOFACI { get; set; }
        public string VOPRNO { get; set; }
        public string VOMFNO { get; set; }
        public string VOOPNO { get; set; }
        public string VOPLGR { get; set; }
        public string VOORQT { get; set; }
        public string VOMAQT { get; set; }
        public string VORSDT { get; set; }
        public string VOSCHN { get; set; }
        public string VOPITI { get; set; }
        public string V_REMA { get; set; }
        public string MMITDS { get; set; }
        public string MMUNMS { get; set; }
        public string VOWOST { get; set; }
        public string VHRORC { get; set; }
        public string VOOPDS { get; set; }
        public string VOCTCD { get; set; }
        public string VOSCQA { get; set; }

    }

    public class RootScheduleNo
    {
        public List<ResultScheduleNo> results { get; set; }
        public bool wasTerminated { get; set; }
        public int nrOfSuccessfullTransactions { get; set; }
        public int nrOfFailedTransactions { get; set; }
    }

    public class ResultScheduleNo
    {
        public string transaction { get; set; }
        public List<ScheduleNo> records { get; set; }
    }

    public class ScheduleNo
    {
        public string SCHN { get; set; }

    }

    public class RootWorkCenter
    {
        public List<ResultWorkCenter> results { get; set; }
        public bool wasTerminated { get; set; }
        public int nrOfSuccessfullTransactions { get; set; }
        public int nrOfFailedTransactions { get; set; }
    }

    public class ResultWorkCenter
    {
        public string transaction { get; set; }
        public List<WorkCenter> records { get; set; }
    }

    public class WorkCenter
    {
        public string VOPLGR { get; set; }

    }

    public class ResultValid
    {
        public string transaction { get; set; }
        public List<object> records { get; set; }
        public string errorMessage { get; set; }
        public string errorType { get; set; }
        public string errorCode { get; set; }
        public string errorCfg { get; set; }
        public string errorField { get; set; }
    }

    public class RootValid
    {
        public List<ResultValid> results { get; set; }
        public bool wasTerminated { get; set; }
        public int nrOfSuccessfullTransactions { get; set; }
        public int nrOfFailedTransactions { get; set; }
    }


    #endregion
}
