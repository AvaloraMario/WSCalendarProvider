using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.Graph;
using WSCalendarProvider.Models;
using Microsoft.Identity.Client;
using System.Security.Authentication;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Configuration;
using System.IO;

namespace WSCalendarProvider.Controllers
{
    [RoutePrefix("calendar")]
    public class CalendarController : ApiController
    {
        static IConfidentialClientApplication app;
        string Secret = GetAppSettingsValues("Secret");
        string TenantId = GetAppSettingsValues("TenantId");
        string ClientId = GetAppSettingsValues("ClientId");
        string RedirectUri = GetAppSettingsValues("RedirectUri");

        public static string[] Scopes = { GetAppSettingsValues("Scope") };

        [Route("GetCalendar")]
        [HttpGet]
        public async Task<string> GetCalendar(string sala)
        {
            try
            {
                var authorization = Request.Headers.Authorization;
                if (authorization != null)
                {
                    List<PairValue> Photos;

                    app = ConfidentialClientApplicationBuilder.Create(ClientId)
                        .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                        .WithClientSecret(Secret)
                        .WithRedirectUri(RedirectUri)
                        .WithTenantId(TenantId)
                        .Build();

                    AuthenticationResult result;
                    var accounts = await app.GetAccountsAsync();
                    try
                    {
                        if (accounts.Any())
                        {
                            IAccount account = accounts.FirstOrDefault();
                            result = await app.AcquireTokenSilent(Scopes, account).ExecuteAsync();
                        }
                        else
                        {
                            result = await app.AcquireTokenForClient(Scopes)
                                              .ExecuteAsync();
                        }
                    }
                    catch (Exception ex)
                    {
                        ToLog.RegisterLog.WriteLog("", ex, "1- GetCalendar", "WSCalendar", "", false);
                        return ex.Message;
                    }

                    if (result == null)
                    {
                        ToLog.RegisterLog.WriteLog("", "Problemas al obtener el token", "GetCalendar", "WSCalendar", "", false);
                        throw new ArgumentException("Problems to take token");
                    }

                    GraphServiceClient client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(async (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("bearer", result.AccessToken);
                        }));

                    var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("startDateTime", DateTime.Today.ToString("yyyy-MM-ddTHH:mm")),
                        new QueryOption("endDateTime",DateTime.Today.AddDays(1).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm")),
                        new QueryOption("Prefer","Romance Standard Time")
                    };
                    var events = await client.Users[sala].CalendarView
                        .Request(queryOptions)
                        .Select(c => new
                        {
                            c.Start,
                            c.End,
                            c.Location,
                            c.Organizer,
                            c.Attendees,
                            c.Subject,
                        }).Top(100)
                        .GetAsync();
                    var ateendes = events.Select(e => e.Attendees.Select(a => a.EmailAddress).Distinct()).Distinct();
                    Photos = new List<PairValue>();

                    foreach (var emailAddress in ateendes)
                    {
                        ToLog.RegisterLog.WriteLog("", "5- Ateendes " + ateendes.Count(), "GetCalendar", "WSCalendar", "", false);
                        foreach (var email in emailAddress)
                        {
                            try
                            {
                                var photo = await client.Users[email.Address].Photo
                                                        .Content
                                                        .Request()
                                                        .GetAsync();
                                Photos.Add(new PairValue() { Key = email.Address, Value = Utils.Util.ReadStreamToBase64(photo) });
                            }
                            catch (Exception ex)
                            {
                                 string name = email.Name.Trim();
                                name = name.Substring(0, 1) + name.Substring(name.IndexOf(' '), 2);
                                Photos.Add(new PairValue() { Key = email.Address, Value = name });
                            }
                        }
                    }
                    ToLog.RegisterLog.WriteLog("", JsonConvert.SerializeObject(events) , "WSCalendar", "GetCalendar Json de CalendarView (sin código de Photos) ", "", false);
                    return JsonConvert.SerializeObject(events) + "^" + JsonConvert.SerializeObject(Photos);
                }
                else
                {
                    throw new System.Security.Authentication.AuthenticationException("Not authorize");
                }
            }
            catch (Exception ex)
            {
                ToLog.RegisterLog.WriteLog("", ex, "WSCalendar", "GetCalendar Final exception", "", false);
                return ex.Message;
            }

        }

        [Route("RoomList")]
        [HttpGet]
        public async Task<string> GetRoomList()
        {
            ToLog.RegisterLog.WriteLog("", "1- Entrando en GetRoomList", "WSCalendar", "GetRoomList", "", false);
            try
            {
                var authorization = Request.Headers.Authorization;
                ToLog.RegisterLog.WriteLog("", "2- Obteniendo autorización ", "WSCalendar", "GetRoomList", "", false);
                if (authorization != null)
                {
                    ToLog.RegisterLog.WriteLog("", "3- Authorization not null ", "WSCalendar", "GetRoomList", "", false);
                    app = ConfidentialClientApplicationBuilder.Create(ClientId)
                        .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                        .WithClientSecret(Secret)
                        .WithRedirectUri(RedirectUri)
                        .WithTenantId(TenantId)
                        .Build();
                    ToLog.RegisterLog.WriteLog("", "4- Aplicación instanciada", "WSCalendar", "GetRoomList", "", false);
                    AuthenticationResult result;
                    var accounts = await app.GetAccountsAsync();
                    ToLog.RegisterLog.WriteLog("", "5- Accounts " + accounts.Count(), "WSCalendar", "GetRoomList", "", false) ;
                    try
                    {
                        if (accounts.Any())
                        {
                            IAccount account = accounts.FirstOrDefault();
                            ToLog.RegisterLog.WriteLog("", "6- Account " + account.Username + " " + account.Environment, "WSCalendar", "GetRoomList", "", false);
                            result = await app.AcquireTokenSilent(Scopes, account).ExecuteAsync();
                            ToLog.RegisterLog.WriteLog("", "7- Result Account " , "WSCalendar", "GetRoomList", "", false);

                        }
                        else
                        {
                            ToLog.RegisterLog.WriteLog("", "8- Accounts = 0 ", "WSCalendar", "GetRoomList", "", false);

                            result = await app.AcquireTokenForClient(Scopes)
                                              .ExecuteAsync();
                             ToLog.RegisterLog.WriteLog("", "9- Accounts = 0 " , "WSCalendar", "GetRoomList", "", false);
                        }
                    }
                    catch (Exception ex)
                    {
                        ToLog.RegisterLog.WriteLog("", ex, "WSCalendar", "GetRoomList", "", false);
                        return ex.Message;
                    }

                    if (result == null)
                    {
                        ToLog.RegisterLog.WriteLog("", "Problemas al obtener el token", "WSCalendar", "GetRoomList", "", false);
                        throw new ArgumentException("Problems to take token");
                    }
                    else
                    {
                        ToLog.RegisterLog.WriteLog("", "10- Result != null ", "WSCalendar", "GetRoomList", "", false); }
                    using (HttpClient client = new HttpClient())
                    {
                        ToLog.RegisterLog.WriteLog("", "11- Entrando en using (HttpClient client = new HttpClient())", "WSCalendar", "GetRoomList", "", false);
                        client.DefaultRequestHeaders.Clear();
                        ToLog.RegisterLog.WriteLog("", "12- Después de Clear ", "WSCalendar", "GetRoomList", "", false);

                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                        ToLog.RegisterLog.WriteLog("", "13- Después de lient.DefaultRequestHeaders.Authorization ", "WSCalendar", "GetRoomList", "", false);
                        client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                        ToLog.RegisterLog.WriteLog("", "14- Después de  client.DefaultRequestHeaders.Accept.Ad ", "WSCalendar", "GetRoomList", "", false);
                        //var response = await client.GetAsync($"https://graph.microsoft.com/beta/users/{calendar}/findRoomLists");
                        var response = await client.GetAsync("https://graph.microsoft.com/v1.0/places/microsoft.graph.room");
                        ToLog.RegisterLog.WriteLog("", "15- Después de llamada a app graph ", "WSCalendar", "GetRoomList", "", false);

                        using (HttpContent httpContent = response.Content)
                        {
                            ToLog.RegisterLog.WriteLog("", "16- Obteniendo Content ", "WSCalendar", "GetRoomList", "", false);
                            var value = await httpContent.ReadAsStringAsync();
                            ToLog.RegisterLog.WriteLog("", "17- VAlue " , "WSCalendar", "GetRoomList", "", false);
                            return JsonConvert.DeserializeObject(value.ToString()).ToString();
                        }
                    }
                }
                else
                {
                    ToLog.RegisterLog.WriteLog("", "18- Not authorize " , "WSCalendar", "GetRoomList", "", false);

                    throw new System.Security.Authentication.AuthenticationException("Not authorize");
                }
            }
            catch (Exception ex)
            {
                ToLog.RegisterLog.WriteLog("","19- "+ ex, "WSCalendar", "GetRoomList", "", false);
                return ex.Message;
            }
        }

        [Route("AddEvent")]
        [HttpPost]
        public async Task<bool> AddEvent(JObject json)
        {
            var subject = string.Empty;
            var location = string.Empty;
            var start = string.Empty;
            var end = string.Empty;
            var calendar = string.Empty;

            try
            {
                var authorization = Request.Headers.Authorization;

                if (authorization != null)
                {
                    subject = json.SelectToken("$..Subject").ToString();
                    location = json.SelectToken("$..Location").ToString();
                    start = json.SelectToken("$..Start").ToString();
                    end = json.SelectToken("$..End").ToString();
                    calendar = json.SelectToken("$..Calendar").ToString();

                    app = ConfidentialClientApplicationBuilder.Create(ClientId)
                        .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                        .WithClientSecret(Secret)
                        .WithRedirectUri(RedirectUri)
                        .WithTenantId(TenantId)
                        .Build();

                    AuthenticationResult result;
                    var accounts = await app.GetAccountsAsync();
                    try
                    {
                        if (accounts.Any())
                        {
                            IAccount account = accounts.FirstOrDefault();
                            result = await app.AcquireTokenSilent(Scopes, account).ExecuteAsync();
                        }
                        else
                        {
                            result = await app.AcquireTokenForClient(Scopes)
                                              .ExecuteAsync();
                        }
                    }
                    catch (Exception ex)
                    {
                        ToLog.RegisterLog.WriteLog("", ex, "AddEvents", "WSCalendar", $"Subject:{subject}, Location:{location}, Start:{start}, End:{end}", false);
                        return false;
                    }

                    if (result == null)
                    {
                        ToLog.RegisterLog.WriteLog("", "Problemas al obtener el token", "AddEvents", "WSCalendar", $"Subject:{subject}, Location:{location}, Start:{start}, End:{end}", false);
                        throw new ArgumentException("Problems to take token");
                    }

                    DateTime dateTimeStart = Convert.ToDateTime(start);
                    DateTime dateTimeEnd = Convert.ToDateTime(end);
                    var stringStart = dateTimeStart.ToString("yyyy-MM-ddTHH:mm");
                    var stringEnd = dateTimeEnd.ToString("yyyy-MM-ddTHH:mm");

                    GraphServiceClient client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(async (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("bearer", result.AccessToken);
                        }));

                    var @event = new Event()
                    {
                        Subject = subject,
                        Location = new Location()
                        {
                            DisplayName = location
                        },
                        Start = new DateTimeTimeZone()
                        {
                            DateTime = stringStart,
                            TimeZone = "Romance Standard Time"
                        },
                        End = new DateTimeTimeZone()
                        {
                            DateTime = stringEnd,
                            TimeZone = "Romance Standard Time"
                        },
                        Attendees = new List<Attendee>()
                        {
                            new Attendee()
                            {
                                EmailAddress = new EmailAddress()
                                {
                                    Address=calendar,
                                    Name = calendar.Split('@')[0]
                                }
                            }
                        }
                    };

                    var a = await client.Users[calendar].Calendar.Events
                        .Request()
                        .AddAsync(@event);

                    if (a == null)
                        return false;

                    return true;
                }
                else
                {
                    throw new System.Security.Authentication.AuthenticationException("Not authorize");
                }
            }
            catch (Exception ex)
            {
                ToLog.RegisterLog.WriteLog("", ex, "AddEvents", "WSCalendar", $"Subject:{subject}, Location:{location}, Start:{start}, End:{end}", false);
                throw ex;
            }
        }

        [Route("SendLog")]
        [HttpPost]
        public bool SendEmail(JObject json)
        {
            string subject = string.Empty;
            string from = string.Empty;
            string to = string.Empty;
            string attached = string.Empty;

            var authorization = Request.Headers.Authorization;

            try
            {
                if (authorization != null)
                {

                    subject = json.SelectToken("$..subject").ToString();
                    from = json.SelectToken("$..from").ToString();
                    to = json.SelectToken("$..to").ToString();
                    attached = json.SelectToken("$..attached").ToString();

                    using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(attached)))
                    {
                        String mailTo = GetAppSettingsValues("mailTo");
                        String userName = GetAppSettingsValues("userName");
                        String password = GetAppSettingsValues("pass");
                        MailMessage msg = new MailMessage();
                        msg.To.Add(new MailAddress(mailTo));
                        msg.From = new MailAddress(userName);
                        msg.Subject = "Test Office 365 Account";
                        msg.Attachments.Add(new System.Net.Mail.Attachment(ms, DateTime.Now + "log.txt"));


                        using (SmtpClient client = new SmtpClient())
                        {
                            client.Host = "avalora.com";
                            client.Credentials = new System.Net.NetworkCredential(userName, password);
                            client.Port = 465;
                            client.EnableSsl = true;
                            client.Send(msg);
                        }
                    }
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                ToLog.RegisterLog.WriteLog("", ex, "SendEmail", "WSCalendar", $"Subject:{subject}, from:{from}, to:{to}, Attached:{attached}", false);
                throw ex;
            }
        }

        private static string GetAppSettingsValues(string key)
        {
            return string.IsNullOrWhiteSpace(ConfigurationManager.AppSettings[key]) ? string.Empty : ConfigurationManager.AppSettings[key].ToString();
        }

        [Route("WriteLog")]
        [HttpPost]
        public async Task<bool> WriteLog(JObject json)
        {
            try
            {
                ToLog.RegisterLog.WriteLog("", json.SelectToken("MessageLog").ToString(), "Reserva de Salas", json.SelectToken("MethodName").ToString(), "", false);
                return true;
            }
            catch (Exception ex)
            {
                ToLog.RegisterLog.WriteLog("", ex.Message, "WriteLog", "WSCalendar","", false);
                return false;
            }
        }
    }
}
