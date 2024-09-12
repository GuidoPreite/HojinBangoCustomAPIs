using Microsoft.Xrm.Sdk;
using System;
using System.Net;
using System.Text;

namespace HojinBangoCustomAPIs.Plugins
{
    public class CheckHojinBango : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            string number = context.InputParameters["Number"] as string;
            bool success = false;
            string companyName1 = "";
            string companyName2 = "";
            string address = "";
            if (!string.IsNullOrEmpty(number))
            {
                try
                {
                    string checkNumberUrl = "https://www.houjin-bangou.nta.go.jp/henkorireki-johoto.html?selHouzinNo=" + number;
                    WebClient webClient = new WebClient { Encoding = Encoding.UTF8 };
                    string pageContent = webClient.DownloadString(checkNumberUrl);
                    if (pageContent.Contains("<dd>" + number + "</dd>"))
                    {
                        string info = GetInsideContent(pageContent, "<dd>" + number + "</dd>", "<dd>", "</dl>");

                        string companyName1Placeholder = "<dt>商号又は名称</dt>";
                        string companyName2Placeholder = "<dt>商号又は名称（フリガナ）</dt>";
                        string addressPlaceholder = "<dt>本店又は主たる事務所の所在地</dt>";

                        companyName1 = GetInsideContent(info, companyName1Placeholder, "<dd>", "</dd>");
                        companyName2 = GetInsideContent(info, companyName2Placeholder, "<dd>", "</dd>");
                        address = GetInsideContent(info, addressPlaceholder, "<dd>", "</dd>");

                        success = true;
                    }
                }
                catch (Exception ex)
                {
                    address = ex.Message;
                }
            }
            context.OutputParameters["Success"] = success;
            context.OutputParameters["CompanyName1"] = companyName1;
            context.OutputParameters["CompanyName2"] = companyName2;
            context.OutputParameters["Address"] = address;
        }

        private string GetInsideContent(string content, string beginPlaceHolder, string startPlaceholder, string endPlaceholder)
        {
            string firstPart = content.Substring(content.IndexOf(beginPlaceHolder));
            string secondPart = firstPart.Substring(0, firstPart.IndexOf(endPlaceholder));
            string thirdPart = secondPart.Substring(secondPart.IndexOf(startPlaceholder) + startPlaceholder.Length);
            return thirdPart.Trim();
        }
    }
}