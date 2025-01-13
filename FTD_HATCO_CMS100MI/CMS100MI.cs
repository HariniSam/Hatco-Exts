using Mongoose.IDO;
using Mongoose.IDO.Protocol;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FTD_HATCO_CMS100MI
{
    public class CMS100MI : IDOExtensionClass
    {
        private int retVal = -1;

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
        public int LstMWOHED02(string sFacility, string sScheduleNumber, string sItem, string sMO, out string sMONumber, out string sOrderQuantity,
            out string sManufacturedQuantity, out string sManQtyBySchedule, out string sOrdQtyBySchedule, out string sItemDescription, out string sUnitofMeasure, out string sMODetails, out string Infobar)
        {
            sOrderQuantity = "";
            sManufacturedQuantity = "";
            sManQtyBySchedule = "";
            sOrdQtyBySchedule = "";
            sItemDescription = "";
            sUnitofMeasure = "";
            sMONumber = "";
            sMODetails = "";
            Infobar = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOHED02"));
            lstparms.Add(new Tuple<string, string>("VHFACI", sFacility));
            lstparms.Add(new Tuple<string, string>("F_SCHN", sScheduleNumber));
            lstparms.Add(new Tuple<string, string>("T_SCHN", sScheduleNumber));
            lstparms.Add(new Tuple<string, string>("F_PRNO", sItem));
            lstparms.Add(new Tuple<string, string>("T_PRNO", sItem));
            lstparms.Add(new Tuple<string, string>("F_MFNO", sMO));
            lstparms.Add(new Tuple<string, string>("T_MFNO", sMO));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                int count = rss["results"][0]["records"].Count();

                if (count == 0)
                {
                    Infobar = "No match found";
                    return retVal;
                }
                else if (count == 1)
                {
                    sMONumber = rss["results"][0]["records"][0]["VHMFNO"].ToString();
                    sOrderQuantity = rss["results"][0]["records"][0]["VHORQT"].ToString();
                    sManufacturedQuantity = rss["results"][0]["records"][0]["VHMAQA"].ToString();
                    sItemDescription = rss["results"][0]["records"][0]["MMITDS"].ToString();
                    sUnitofMeasure = rss["results"][0]["records"][0]["MMUNMS"].ToString();

                }
                else if (count > 1)
                {
                    sMONumber = "*";
                    int iOrderQuantity = 0;
                    int iManufacturedQuantity = 0;
                    List<MOInfo> iMOList = new List<MOInfo>();

                    for (int i = 0; i < count; i++)
                    {
                        iOrderQuantity = iOrderQuantity + int.Parse(rss["results"][0]["records"][i]["VHORQT"].ToString());
                        iManufacturedQuantity = iManufacturedQuantity + int.Parse(rss["results"][0]["records"][i]["VHMAQA"].ToString());

                        MOInfo iMO = new MOInfo()
                        {
                            orderQty = int.Parse(rss["results"][0]["records"][i]["VHORQT"].ToString()),
                            manufactQty = int.Parse(rss["results"][0]["records"][i]["VHMAQA"].ToString()),
                            moNumber = rss["results"][0]["records"][i]["VHMFNO"].ToString()
                        };

                        iMOList.Add(iMO);
                    }

                    sOrderQuantity = iOrderQuantity.ToString();
                    sManufacturedQuantity = iManufacturedQuantity.ToString();
                    sItemDescription = rss["results"][0]["records"][0]["MMITDS"].ToString();
                    sUnitofMeasure = rss["results"][0]["records"][0]["MMUNMS"].ToString();
                    sMODetails = JsonConvert.SerializeObject(iMOList);
                }

                // Get total Qty by Schedule
                LstMWOHED02_3(sFacility, sScheduleNumber, ref sOrdQtyBySchedule, ref sManQtyBySchedule, ref Infobar);

                retVal = 0;
            }

            return retVal;
        }

        [IDOMethod(Flags = MethodFlags.None)]
        public int LstMWOHED02_2(string sFacility, string sScheduleNumber, string sItem, string sMONumber, out string sOrderQuantity,
            out string sManufacturedQuantity, out string Infobar)
        {
            sOrderQuantity = "";
            sManufacturedQuantity = "";
            Infobar = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOHED02"));
            lstparms.Add(new Tuple<string, string>("VHFACI", sFacility));
            lstparms.Add(new Tuple<string, string>("F_SCHN", sScheduleNumber));
            lstparms.Add(new Tuple<string, string>("T_SCHN", sScheduleNumber));
            lstparms.Add(new Tuple<string, string>("F_PRNO", sItem));
            lstparms.Add(new Tuple<string, string>("T_PRNO", sItem));
            lstparms.Add(new Tuple<string, string>("F_MFNO", sMONumber));
            lstparms.Add(new Tuple<string, string>("T_MFNO", sMONumber));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                sOrderQuantity = rss["results"][0]["records"][0]["VHORQT"].ToString();
                sManufacturedQuantity = rss["results"][0]["records"][0]["VHMAQA"].ToString();

                retVal = 0;
            }

            return retVal;
        }


        private void LstMWOHED02_3(string sFacility, string sScheduleNumber, ref string sOrderQuantity, ref string sManufacturedQuantity, ref string Infobar)
        {
            sOrderQuantity = "";
            sManufacturedQuantity = "";
            Infobar = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOHED02"));
            lstparms.Add(new Tuple<string, string>("VHFACI", sFacility));
            lstparms.Add(new Tuple<string, string>("F_SCHN", sScheduleNumber));
            lstparms.Add(new Tuple<string, string>("T_SCHN", sScheduleNumber));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                int iOrderQuantity = 0;
                int iManufacturedQuantity = 0;

                JObject rss = JObject.Parse(result);

                int count = rss["results"][0]["records"].Count();

                for (int i = 0; i < count; i++)
                {
                    iOrderQuantity = iOrderQuantity + int.Parse(rss["results"][0]["records"][i]["VHORQT"].ToString());
                    iManufacturedQuantity = iManufacturedQuantity + int.Parse(rss["results"][0]["records"][i]["VHMAQA"].ToString());

                }

                sOrderQuantity = iOrderQuantity.ToString();
                sManufacturedQuantity = iManufacturedQuantity.ToString();

            }
        }



        [IDOMethod(Flags = MethodFlags.None)]
        public int GetMONumber(string sFacility, string sScheduleNumber, string sItem, int Qty, out string sMONumber, out string Infobar)
        {
            sMONumber = "";
            Infobar = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOHED02"));
            lstparms.Add(new Tuple<string, string>("VHFACI", sFacility));
            lstparms.Add(new Tuple<string, string>("VHSCHN", sScheduleNumber));
            lstparms.Add(new Tuple<string, string>("VHPRNO", sItem));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                int count = rss["results"][0]["records"].Count();

                if (count > 1)
                {
                    sMONumber = "*";
                    int iOrderQuantity = 0;

                    for (int i = 0; i < count; i++)
                    {
                        iOrderQuantity = iOrderQuantity + int.Parse(rss["results"][0]["records"][i]["VHORQT"].ToString());
                        if (iOrderQuantity >= Qty)
                        {
                            sMONumber = rss["results"][0]["records"][i]["VHMFNO"].ToString();
                        }
                    }
                }

                retVal = 0;
            }
            return retVal;
        }

        [IDOMethod(Flags = MethodFlags.CustomLoad)]
        public IDataReader LstMWOOPE02(string Facility, string ScheduleNumber)
        {

            string Infobar = "";
            DataTable operationList = new DataTable("Operations");
            operationList.Columns.Add("VOFACI");
            operationList.Columns.Add("VOPRNO");
            operationList.Columns.Add("VOMFNO");
            operationList.Columns.Add("VOOPNO");
            operationList.Columns.Add("VOPLGR");
            operationList.Columns.Add("VOORQT");
            operationList.Columns.Add("VOMAQT");
            operationList.Columns.Add("VOSCHN");
            operationList.Columns.Add("V_REMA");
            operationList.Columns.Add("VOWOST");
            operationList.Columns.Add("Message");
            int maxrecs = 1000;

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOOPE02"));
            lstparms.Add(new Tuple<string, string>("VOFACI", Facility));
            lstparms.Add(new Tuple<string, string>("VOSCHN", ScheduleNumber));
            lstparms.Add(new Tuple<string, string>("maxrecs", maxrecs.ToString()));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                for (int i = 0; i < rss["results"][0]["records"].Count(); i++)
                {
                    DataRow row = operationList.NewRow();
                    row["VOFACI"] = rss["results"][0]["records"][i]["VOFACI"].ToString();
                    row["VOPRNO"] = rss["results"][0]["records"][i]["VOPRNO"].ToString();
                    row["VOMFNO"] = rss["results"][0]["records"][i]["VOMFNO"].ToString();
                    row["VOOPNO"] = rss["results"][0]["records"][i]["VOOPNO"].ToString();
                    row["VOPLGR"] = rss["results"][0]["records"][i]["VOPLGR"].ToString();
                    row["VOORQT"] = rss["results"][0]["records"][i]["VOORQT"].ToString();
                    row["VOMAQT"] = rss["results"][0]["records"][i]["VOMAQT"].ToString();
                    row["VOSCHN"] = rss["results"][0]["records"][i]["VOSCHN"].ToString();
                    row["V_REMA"] = rss["results"][0]["records"][i]["V_REMA"].ToString();
                    row["VOWOST"] = rss["results"][0]["records"][i]["VOWOST"].ToString();
                    row["Message"] = string.Empty;
                    //string ReferenceOrderCat = rss["results"][0]["records"][0]["VHRORC"].ToString();
                    //string Startdate = rss["results"][0]["records"][0]["VOSTDT"].ToString();
                    //string Description = rss["results"][0]["records"][0]["MMITDS"].ToString();
                    //string UnitofMeasure = rss["results"][0]["records"][0]["MMUNMS"].ToString();
                    //string RunTime = rss["results"][0]["records"][0]["VOPITI"].ToString();
                    operationList.Rows.Add(row);
                }
            }
            else
            {
                DataRow row = operationList.NewRow();
                row["Message"] = Infobar;
                operationList.Rows.Add(row);
            }

            return operationList.CreateDataReader(); ;
        }

        [IDOMethod(Flags = MethodFlags.CustomLoad)]
        public IDataReader LstMWOOPE03(string Facility, string ScheduleNumber)
        {

            string Infobar = "";
            DataTable operationList = new DataTable("Operations");
            operationList.Columns.Add("VOFACI");
            operationList.Columns.Add("VOPRNO");
            operationList.Columns.Add("VOMFNO");
            operationList.Columns.Add("VOOPNO");
            operationList.Columns.Add("VOPLGR");
            operationList.Columns.Add("VOORQT");
            operationList.Columns.Add("VOMAQT");
            operationList.Columns.Add("VOSCHN");
            operationList.Columns.Add("V_REMA");
            operationList.Columns.Add("VOWOST");
            operationList.Columns.Add("VHMAQA");
            operationList.Columns.Add("VOPITI");
            operationList.Columns.Add("VOCTCD");
            operationList.Columns.Add("Message");
            int maxrecs = 1000;

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/CMS100MI/LstMWOOPE03"));
            lstparms.Add(new Tuple<string, string>("VOFACI", Facility));
            lstparms.Add(new Tuple<string, string>("VOSCHN", ScheduleNumber));
            lstparms.Add(new Tuple<string, string>("maxrecs", maxrecs.ToString()));

            string sMethodList = BuildMethodParms(lstparms);
            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);

                for (int i = 0; i < rss["results"][0]["records"].Count(); i++)
                {
                    DataRow row = operationList.NewRow();
                    row["VOFACI"] = rss["results"][0]["records"][i]["VOFACI"].ToString();
                    row["VOPRNO"] = rss["results"][0]["records"][i]["VOPRNO"].ToString();
                    row["VOMFNO"] = rss["results"][0]["records"][i]["VOMFNO"].ToString();
                    row["VOOPNO"] = rss["results"][0]["records"][i]["VOOPNO"].ToString();
                    row["VOPLGR"] = rss["results"][0]["records"][i]["VOPLGR"].ToString();
                    row["VOORQT"] = rss["results"][0]["records"][i]["VOORQT"].ToString();
                    row["VOMAQT"] = rss["results"][0]["records"][i]["VOMAQT"].ToString();
                    row["VOSCHN"] = rss["results"][0]["records"][i]["VOSCHN"].ToString();
                    row["V_REMA"] = rss["results"][0]["records"][i]["V_REMA"].ToString();
                    row["VOWOST"] = rss["results"][0]["records"][i]["VOWOST"].ToString();
                    row["VHMAQA"] = rss["results"][0]["records"][i]["VHMAQA"].ToString();
                    row["VOPITI"] = rss["results"][0]["records"][i]["VOPITI"].ToString();
                    row["VOCTCD"] = rss["results"][0]["records"][i]["VOCTCD"].ToString();
                    row["Message"] = string.Empty;
                    //string ReferenceOrderCat = rss["results"][0]["records"][0]["VHRORC"].ToString();
                    //string Startdate = rss["results"][0]["records"][0]["VOSTDT"].ToString();
                    //string Description = rss["results"][0]["records"][0]["MMITDS"].ToString();
                    //string UnitofMeasure = rss["results"][0]["records"][0]["MMUNMS"].ToString();
                    //string RunTime = rss["results"][0]["records"][0]["VOPITI"].ToString();
                    operationList.Rows.Add(row);
                }
            }
            else
            {
                DataRow row = operationList.NewRow();
                row["Message"] = Infobar;
                operationList.Rows.Add(row);
            }

            return operationList.CreateDataReader(); ;
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

            string sso = "1";
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

            stopwatch.Stop();
            TimeSpan ts = stopwatch.Elapsed;
            string sTimeTaken = (ts.Hours > 0 ? ts.Hours + ":" : "") + (ts.Minutes > 0 ? ts.Minutes + ":" : "") + ts.Seconds + "." + ts.Milliseconds;

            CreateLog(sMethod, result, errMsg, sTimeTaken, "FG Reporting");


            return result;
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

    public class MOInfo
    {
        public int orderQty { get; set; }
        public int manufactQty { get; set; }
        public string moNumber { get; set; }
    }
}
