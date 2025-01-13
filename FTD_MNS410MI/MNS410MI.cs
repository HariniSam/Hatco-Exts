using Mongoose.IDO;
using Mongoose.IDO.Protocol;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace FTD_MNS410MI
{
    public class MNS410MI : IDOExtensionClass
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
        public int GetUserAccess(string UserID, string UserRole, out string ValidTo, out string ValidFrom, out string Infobar)
        {
            Infobar = "";
            ValidTo = "";
            ValidFrom = "";


            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/MNS410MI/Get"));
            lstparms.Add(new Tuple<string, string>("USID", UserID));
            lstparms.Add(new Tuple<string, string>("ROLL", UserRole));
            string sMethodList = BuildMethodParms(lstparms);


            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                JObject rss = JObject.Parse(result);
                UserID = rss["results"][0]["records"][0]["USID"].ToString();
                ValidFrom = rss["results"][0]["records"][0]["FVDT"].ToString();
                ValidTo = rss["results"][0]["records"][0]["VTDT"].ToString();
                retVal = 0;
            }

            return retVal;
        }

        [IDOMethod(Flags = MethodFlags.None)]
        public int LstUserRoles(string UserID, out string UserRole, out string Infobar)
        {
            Infobar = "";
            UserRole = "";

            List<Tuple<string, string>> lstparms = new List<Tuple<string, string>>();
            lstparms.Add(new Tuple<string, string>("method", "/m3api-rest/v2/execute/MNS410MI/Lst"));
            lstparms.Add(new Tuple<string, string>("USID", UserID));
            string sMethodList = BuildMethodParms(lstparms);


            string result = ProcessIONAPI(sMethodList, out Infobar);

            if (string.IsNullOrEmpty(Infobar))
            {
                List<string> UserRoleList = new List<string>();
                JObject rss = JObject.Parse(result);
                for (int i = 0; i < rss["results"][0]["records"].Count(); i++)
                {
                    UserRoleList.Add(rss["results"][0]["records"][i]["ROLL"].ToString());
                }
                UserRole = string.Join(",", UserRoleList);
            }

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
}
