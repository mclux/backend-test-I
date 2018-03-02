using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Helpers;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using System.Security.Cryptography;
using LinqToTwitter;
using Spring.Social.Twitter.Api.Impl;
using Spring.Social.Twitter.Api;
using Newtonsoft.Json;

namespace DevCenterBot
{
    class Program
    {
        static void Main(string[] args)
        {
            TwitterHelper tw = new TwitterHelper();

            Console.Write("[+]Enter hashtag:");
            string hastTag = Console.ReadLine();
            Console.WriteLine("[+]searching tweets...");

            var response = tw.GetTweets(hastTag, 100);
            var obj = JObject.Parse(response);
            List<ExcelUserVM> lists = new List<ExcelUserVM>();
            foreach (var child in obj["statuses"].Children())
            {
                var childObj = child.First().Parent;
                var name = childObj["user"]["name"].ToString();
                var followerCount = childObj["user"]["followers_count"].ToString();
                lists.Add(new ExcelUserVM
                {
                    Name = name,
                    FollowerCount =Convert.ToInt32(followerCount)
                });
                Console.WriteLine("[+]Profile Name: {0}\n[+]No. of Followers:{1}",name, followerCount);
                Console.WriteLine("------------------------------------------");                               
            }
            ExcelHelper ex = new ExcelHelper();
            ex.ExportExcel(lists);
            Console.Read();            
        }        
    }
    
}
