using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.Web;
using System.Security.Permissions;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Data.SqlClient;
using RECEIVER.Utils;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Collections;
using System.Data;


namespace ConsApp_Insert_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            using (SPSite siteCollection = new SPSite("portal"))
            {

                SPWebCollection sites = siteCollection.AllWebs;
                SPWeb site = siteCollection.OpenWeb();

                try
                {

                    SPListItemCollection listItems = site.GetList(siteCollection.Url + "/Lists/testSendEmail/AllItems.aspx").Items;

                    SPListItem item = listItems.Add();


                    string queryString = "SELECT DISTINCT [email] FROM [table] where [email] <> ''";
                    using (var manager = new ManagerSQL(site))
                    {
                        List<Dictionary<string, string>> result = manager.SelectAll(queryString.Replace(@"\\", @"\"));
                        //Dictionary<string, string> result = manager.Select(queryString.Replace(@"\\", @"\"));

                        item["Title"] = DateTime.Now; 
                        for (int i = 0; i < result.Count; i++)
                        {
                            item["textEmail"] += result[i]["email"] + ";";
                        }
                        item.Update();
                    }
                }
                catch (Exception ex)
                {
                    SPListItemCollection listItems = site.GetList(siteCollection.Url + "/Lists/testSendEmail/AllItems.aspx").Items;
                    SPListItem item = listItems.Add();
                    item["textEmail"] = "123@mail.ru";
                    item.Update();
                    Console.WriteLine(ex.Message);
                }

                finally
                {
                    if (site != null)
                        site.Dispose();
                }


            }
            //Console.Write("Press ENTER to continue");
            //Console.ReadLine();

        }
    }
}

