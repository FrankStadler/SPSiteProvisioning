using System;
using System.Net;
using OfficeDevPnP.Core;
using PnPAuthenticationManager = OfficeDevPnP.Core.AuthenticationManager;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;

namespace SPSiteProvisioning
{
    public static class SPSiteProvisioning
    {
        [FunctionName("SPSiteProvisioning")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Admin, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            dynamic data = await req.Content.ReadAsAsync<object>();

            string SPParentSiteUrl = data["SPParentSiteUrl"];
            string SPWebTemplate = data["SPWebTemplate"];
            string SPSiteTitle = data["SPSiteTitle"];
            string SPSiteDescription = data["SPSiteDescription"];
            string SPSiteURL = data["SPSiteURL"];
            int SPSiteLanguage = data["SPSiteLanguage"];

            log.Info($"SPParentSiteUrl = '{SPParentSiteUrl}'");
            log.Info($"SPWebTemplate = '{SPWebTemplate}'");
            log.Info($"SPSiteTitle = '{SPSiteTitle}'");
            log.Info($"SPSiteDescription = '{SPSiteDescription}'");
            log.Info($"SPSiteURL = '{SPSiteURL}'");
            log.Info($"SPSiteLanguage = '{SPSiteLanguage}'");

            string userName = System.Environment.GetEnvironmentVariable("SPUser", EnvironmentVariableTarget.Process);
            string password = System.Environment.GetEnvironmentVariable("SPPwd", EnvironmentVariableTarget.Process);

            var authenticationManager = new PnPAuthenticationManager();
            var clientContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(SPParentSiteUrl, userName, password);
            var pnpClientContext = PnPClientContext.ConvertFrom(clientContext);

            var webCreationInfo = new Microsoft.SharePoint.Client.WebCreationInformation();
            webCreationInfo.WebTemplate = SPWebTemplate;
            webCreationInfo.Title = SPSiteTitle;
            webCreationInfo.Description = SPSiteDescription;
            webCreationInfo.Url = SPSiteURL;
            webCreationInfo.Language = SPSiteLanguage;

            pnpClientContext.Web.Webs.Add(webCreationInfo);
            pnpClientContext.ExecuteQuery();

            return req.CreateResponse(HttpStatusCode.OK, "request done");
        }
    }
}