using Mongoose.IDO;
using Mongoose.IDO.Protocol;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace FTD_PMS070MI
{
    public class PMS070MI : IDOExtensionClass
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
        public int RptOperation(string company, string facility, string productNo, string manufactNo, string operationNo, string date, string localtime, int qty, int scrapQty, string usedLabourTime, out string Infobar)
        {
            Infobar = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/PMS070MI/RptOperation"));
            lstparms.Add(new Tuple<string, string>("CONO", company));
            lstparms.Add(new Tuple<string, string>("FACI", facility));
            lstparms.Add(new Tuple<string, string>("PRNO", productNo));
            lstparms.Add(new Tuple<string, string>("MFNO", manufactNo));
            lstparms.Add(new Tuple<string, string>("OPNO", operationNo));
            lstparms.Add(new Tuple<string, string>("RPDT", DateTime.ParseExact(date, "yyyy'/'MM'/'dd", CultureInfo.InvariantCulture).ToString("yyyyMMdd")));
            lstparms.Add(new Tuple<string, string>("RPTM", localtime));
            lstparms.Add(new Tuple<string, string>("MAQA", qty.ToString()));
            lstparms.Add(new Tuple<string, string>("SCQA", scrapQty.ToString()));
            lstparms.Add(new Tuple<string, string>("UMAT", usedLabourTime));
            lstparms.Add(new Tuple<string, string>("REND", "1"));
            lstparms.Add(new Tuple<string, string>("DSP1", "1"));
            lstparms.Add(new Tuple<string, string>("DSP2", "1"));
            lstparms.Add(new Tuple<string, string>("DSP3", "1"));
            lstparms.Add(new Tuple<string, string>("DSP4", "1"));

            string sMethodList = BuildMethodParms(lstparms);

            string result = ProcessIONAPI(sMethodList, out Infobar);


            return retVal;
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

            CreateLog(sMethod, result, errMsg);


            return result;
        }

        private void CreateLog(string sRequest, string sResponse, string sErrorMessage)
        {
            sRequest = new string(sRequest.Take(4000).ToArray());
            sResponse = new string(sResponse.Take(4000).ToArray());
            sErrorMessage = new string(sErrorMessage.Take(4000).ToArray());

            UpdateCollectionRequestData updateRequest = new UpdateCollectionRequestData("FTD_M3ConnectionLogs");
            IDOUpdateItem updateItem = new IDOUpdateItem(UpdateAction.Insert);
            updateItem.Properties.Add("Request", sRequest);
            updateItem.Properties.Add("Response", sResponse);
            updateItem.Properties.Add("ErrorMessage", sErrorMessage);
            updateRequest.Items.Add(updateItem);
            this.Context.Commands.UpdateCollection(updateRequest);
        }

    }
}
