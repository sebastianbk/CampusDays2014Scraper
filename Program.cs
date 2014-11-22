using HtmlAgilityPack;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CampusDays2014Scraper
{
    class Program
    {
        static void Main(string[] args)
        {
            var rootUrl = "http://channel9.msdn.com";
            RegexOptions options = RegexOptions.None;
            Regex regex = new Regex(@"\s+", options);
            Regex digitsRegex = new Regex(@"\d", options);

            var webClient = new WebClient();
            webClient.Encoding = Encoding.UTF8;
            var frontPage = webClient.DownloadString(rootUrl + "/Events/Microsoft-Campus-Days/Microsoft-Campus-Days-2014?sort=status&direction=asc");
            var doc = new HtmlDocument();
            doc.LoadHtml(frontPage);

            var sessions = new List<Session>();

            foreach (var link in doc.DocumentNode.SelectNodes("//a[@class=\"title\"]"))
            {
                string sessionUrl = string.Empty;

                try
                {
                    sessionUrl = link.Attributes.SingleOrDefault(a => a.Name == "href").Value;
                }
                catch
                {
                }

                try
                {
                    var sessionPage = webClient.DownloadString(rootUrl + sessionUrl);
                    var sessionDoc = new HtmlDocument();
                    sessionDoc.LoadHtml(WebUtility.HtmlDecode(sessionPage));

                    var session = new Session();
                    try
                    {
                        session.SessionCode = sessionDoc.DocumentNode.SelectSingleNode("//li[@class=\"code\"]").InnerText.Trim();
                    }
                    catch
                    {
                        session.SessionCode = "General";
                    }

                    if (sessions.Select(s => s.SessionCode).Contains(session.SessionCode))
                        continue;

                    session.Link = rootUrl + sessionUrl;
                    session.Title = sessionDoc.DocumentNode.SelectSingleNode("//h1").InnerText.Trim();
                    session.DateTime = regex.Replace(sessionDoc.DocumentNode.SelectSingleNode("//li[@class=\"date\"]").InnerText.Replace("Date: ", "").Trim(), " ");
                    session.Day = sessionDoc.DocumentNode.SelectSingleNode("//li[@class=\"day\"]").InnerText.Trim();
                    session.Location = sessionDoc.DocumentNode.SelectSingleNode("//li[@class=\"room\"]").InnerText.Trim();
                    session.Speakers = regex.Replace(sessionDoc.DocumentNode.SelectSingleNode("//li[@class=\"speakers\"]").InnerText.Replace("Speakers: ", "").Trim(), " ");
                    session.Description = regex.Replace(sessionDoc.DocumentNode.SelectSingleNode("//div[@id=\"entry-body\"]").InnerText.Trim(), " ");
                    session.TrackCode = digitsRegex.Replace(session.SessionCode, "").ToLower();

                    switch (digitsRegex.Replace(session.SessionCode, "").ToUpper())
                    {
                        case "DATA":
                            session.Track = "Datacenter Modernisering I & II";
                            break;
                        case "SQL":
                            session.Track = "SQL Server & Business Intelligence";
                            break;
                        case "WIN":
                            session.Track = "Windows Client";
                            break;
                        case "WEB":
                            session.Track = "Microsoft Azure og Webudvikling";
                            break;
                        case "VS":
                            session.Track = "Visual Studio og Team Foundation Server";
                            break;
                        case "APP":
                            session.Track = "App Development";
                            break;
                        case "PROD":
                            session.Track = "Produktivitet & Samarbejde";
                            break;
                        case "UC":
                            session.Track = "Microsoft Unified Communications";
                            break;
                        case "NAV":
                            session.Track = "Microsoft Dynamics C5 & NAV";
                            break;
                        case "CRM":
                            session.Track = "Microsoft Dynamics CRM";
                            break;
                        case "AX":
                            session.Track = "Microsoft Dynamics AX";
                            break;
                        default:
                            session.Track = "General";
                            session.TrackCode = "general";
                            break;
                    }

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(session.Title);
                    Console.ResetColor();
                    Console.WriteLine(session.DateTime);
                    Console.WriteLine(session.Day);
                    Console.WriteLine(session.Speakers);

                    sessions.Add(session);
                }
                catch (Exception)
                {
                    continue;
                }
            }

            Console.WriteLine();

            foreach (var track in sessions.Select(s => s.Track).Distinct())
            {
                Console.WriteLine(track);
            }

            var file = new System.IO.StreamWriter(@"C:\Temp\campusdays2014.json", false);
            file.WriteLine(JsonConvert.SerializeObject(sessions));
            file.Close();

            Console.ReadKey();
        }
    }
}
