using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using HtmlAgilityPack;
using SendGrid;
using SendGrid.Helpers.Mail;

namespace QueryCollaboratorUsers
{
    class Program
    {
        private static List<User> _activeUsers = new List<User>();

        // how many users can be online at the same time, if exceeds this number, then the exe will logoff the users
        private const int MaxOnlineUser = 3;

        // the user inactive for 10 minutes will be logoff
        private const double InactiveDurationThresholdInSeconds = 15 * 60;

        static void Main(string[] args)
        {
            var activeUserPageHtmlString = GetActiveUserPage();
            ParseActiveUserPage(activeUserPageHtmlString);

            var logoffUsersList = LogoffUsersByInactiveDuration();
            var forceLogoffList = LogoffUsersByOrder();

            logoffUsersList.AddRange(forceLogoffList);

            if (logoffUsersList.Any())
            {
                //SendEmailsForLogoffUsers(logoffUsersList);
                AddLogs(logoffUsersList);
            }

            // sleeps to let us see the console outputs
            System.Threading.Thread.Sleep(5000);
            Environment.Exit(0);
        }

        private static string GetActiveUserPage()
        {
            /*
            * How to get the token
            * http://svcndalcodeco:8080/services/json/v1
            * [
                {"command" : "Examples.checkLoggedIn"},
                {"command" : "SessionService.getLoginTicket",
                    "args":{"login":"username","password":"pwd"}},
                {"command" : "Examples.checkLoggedIn"}
            ]
            */
            // get active users
            var proc = new Process
            {
                StartInfo =
        {
          FileName = "C:\\Program Files\\Collaborator Client\\ccollab.exe",
          UseShellExecute = false,
          RedirectStandardInput = false,
          RedirectStandardOutput = true,
          RedirectStandardError = true,
          CreateNoWindow = true,
          Arguments = $"admin wget \"go?formSubmitteduserOpts=1&collaborator.security.token=a036ffcb5579453aedc7c7cfecd446a2&pv_component=ErrorsAndMessages&pv_pageNumber=1&pv_itemsPerPage=100&pv_step=AdminUsers&page=Admin&pv_ErrorsAndMessages_fingerPrint=861101&offsetX=0&offsetY=0&userListSort=LOGIN_ASC&userListFilter=SHOW_ACTIVE_PER_LAST_HOUR&searchByUserName=\""
        }
            };

            proc.Start();
            var outputResult = proc.StandardOutput.ReadToEnd();
            var error = proc.StandardError.ReadToEnd();
            proc.WaitForExit();

            return outputResult;
        }

        private static void ParseActiveUserPage(string activeUserPageHtmlString)
        {
            // parse the HTML
            var doc = new HtmlDocument();
            doc.LoadHtml(activeUserPageHtmlString);

            var wizardPanelTableNode = doc.DocumentNode.SelectSingleNode("//table");
            var userListTableNode = wizardPanelTableNode.SelectSingleNode("//tr/td[2]/div[2]/div[1]/div[2]/div[2]/table[2]");
            foreach (HtmlNode rowNode in userListTableNode.SelectNodes("tr"))
            {
                var cellNodes = rowNode.SelectNodes("td");
                var lastActivityTime = ElementParser.GetLastActivityTime(cellNodes[7].InnerText);
                _activeUsers.Add(new User
                {
                    DisplayName = cellNodes[5].InnerText,
                    LastActivityTime = lastActivityTime,
                    LogOffUri = ElementParser.GetLogUri(cellNodes[1]),
                    Email = ElementParser.GetEmail(cellNodes[5])
                });
            }

            // sort by datetime asc
            _activeUsers = _activeUsers.OrderBy(x => x.LastActivityTime).ToList();

            // outputs to the console
            Console.WriteLine("Belows are the currect active users:");
            foreach (var user in _activeUsers)
            {
                Console.WriteLine($"{user.DisplayName}\t\t{user.LastActivityTime}");
            }
        }

        // logoff the user who is inactive for 20 minuts, and return the logoff user list
        private static List<User> LogoffUsersByInactiveDuration()
        {
            var logoffList = new List<User>();
            foreach (var user in _activeUsers)
            {
                var duration = (DateTime.Now - user.LastActivityTime).TotalSeconds;
                if (duration >= InactiveDurationThresholdInSeconds)
                {
                    logoffList.Add(user);
                }
            }

            foreach (var userTobeLogoff in logoffList)
            {
                LogoffUser(userTobeLogoff);
                _activeUsers.Remove(userTobeLogoff);
            }

            return logoffList;
        }

        private static List<User> LogoffUsersByOrder()
        {
            var logoffList = new List<User>();
            if (_activeUsers.Count <= MaxOnlineUser) return logoffList;

            do
            {
                var userTobeLogOff = _activeUsers[0];
                logoffList.Add(userTobeLogOff);
                LogoffUser(userTobeLogOff);

                // remove the user from the list
                _activeUsers.RemoveAt(0);
            } while (_activeUsers.Count > MaxOnlineUser);
            return logoffList;
        }

        private static void LogoffUser(User userTobeLogoff)
        {
            var logoffProc = new Process
            {
                StartInfo =
        {
          FileName = "C:\\Program Files\\Collaborator Client\\ccollab.exe",
          UseShellExecute = false,
          RedirectStandardInput = false,
          RedirectStandardOutput = true,
          RedirectStandardError = true,
          CreateNoWindow = true,
          Arguments = $"admin wget \"{userTobeLogoff.LogOffUri}\""
        }
            };

            logoffProc.Start();
            var logoffOutputResult = logoffProc.StandardOutput.ReadToEnd();
            var logoffError = logoffProc.StandardError.ReadToEnd();
            logoffProc.WaitForExit();
        }


        private static void AddLogs(List<User> logoffUsersList)
        {
            var logContent = new StringBuilder();
            logContent.AppendLine($"Time: {DateTime.Now.ToString("O")}");
            foreach (var user in logoffUsersList)
            {
                var item = $"Display Name: {user.DisplayName}, Last Activity Time: {user.LastActivityTime}";
                logContent.AppendLine(item);
            }

            var filename = @"c:\temp\kickCollaboratorUsers.log";
            File.WriteAllText(filename, logContent.ToString());
        }

        private static void SendEmailsForLogoffUsers(List<User> logoffUsers)
        {
            try
            {
                // Create new send grid client

                var exeRunningPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var configFile = Path.Combine(exeRunningPath, "qcuk.ky");
                var data = File.ReadAllLines(configFile);
                var sendGridKey = data[0];

                var toEmailFullAddressNames = new List<ToEmailFullAddressName>();
                for (int lineIndex = 1; lineIndex < data.Length;)
                {
                    var email = data[lineIndex++];
                    var name = data[lineIndex++];
                    toEmailFullAddressNames.Add(new ToEmailFullAddressName(email, name));
                }

                var client = new SendGridClient(sendGridKey);

                var textEmailBody = new StringBuilder();
                var htmlEmailBody = new StringBuilder();
                foreach (var user in logoffUsers)
                {
                    var item = $"Display Name: {user.DisplayName}, Last Activity Time: {user.LastActivityTime}";
                    textEmailBody.AppendLine(item);
                    textEmailBody.AppendLine();

                    htmlEmailBody.AppendLine(item);
                    textEmailBody.AppendLine("<br>");

                    // send email to the logoff user?
                    toEmailFullAddressNames.Add(new ToEmailFullAddressName(user.Email, user.DisplayName));
                }

                // Add email message info
                const string promptMessage = "The below accounts are logoff to release the license seats";
                var msg = new SendGridMessage
                {
                    From = new EmailAddress("CollaboratorAdmin@beckman.com", "Collaborator Admin"),
                    Subject = promptMessage,
                    PlainTextContent = promptMessage + Environment.NewLine + textEmailBody,
                    HtmlContent = promptMessage + "<br>" + htmlEmailBody
                };

                // outputs to console
                Console.WriteLine(msg.PlainTextContent);

                var toEmailAddressList = new List<EmailAddress>();
                foreach (var toEachEmail in toEmailFullAddressNames)
                {
                    var toEmailAddress = new EmailAddress(toEachEmail.ToEmailAddress, toEachEmail.ToName);
                    toEmailAddressList.Add(toEmailAddress);
                }
                msg.AddTos(toEmailAddressList);

                // Call async send email in sync way
                var response = client.SendEmailAsync(msg).Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine("SendEmail: {0}", ex);
            }
        }
    }


    class User
    {
        public string DisplayName { get; set; }
        public DateTime LastActivityTime { get; set; }
        public string LogOffUri { get; set; }
        public string Email { get; set; }
    }
}
