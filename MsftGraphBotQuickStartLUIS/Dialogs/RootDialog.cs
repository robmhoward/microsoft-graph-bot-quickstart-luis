using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using BotAuth.Models;
using System.Configuration;
using BotAuth.Dialogs;
using BotAuth.AADv2;
using System.Threading;
using System.Net.Http;
using BotAuth;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace MsftGraphBotQuickStart.Dialogs
{
    [LuisModel("d1f0646d-4927-4aac-8f44-8d8e8da84965", "782118274fc846a793b16d7ebdf8770a")]
    [Serializable]
    public class RootDialog : LuisDialog<IMessageActivity>
    {
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            await context.PostAsync("I didn't understand your query...I'm just a simple bot that searches OneDrive. Try a query similar to these:<br/>'find all music'<br/>'find all .pptx files'<br/>'search for mydocument.docx'");
        }

        [LuisIntent("BlockCalendar")]
        public async Task BlockCalendar(IDialogContext context, LuisResult result)
        {
            await context.PostAsync("Block calendar goes here");
        }

        private class When
        {
            public DateTime start;
            public DateTime end;
        }

        private int DayOfWeekIndex(DayOfWeek dayOfWeek)
        {
            switch (dayOfWeek)
            {
                case DayOfWeek.Sunday:
                    return 0;
                case DayOfWeek.Monday:
                    return 1;
                case DayOfWeek.Tuesday:
                    return 2;
                case DayOfWeek.Wednesday:
                    return 3;
                case DayOfWeek.Thursday:
                    return 4;
                case DayOfWeek.Friday:
                    return 5;
                case DayOfWeek.Saturday:
                    return 6;
                default:
                    return -1;
            }
        }

        private When GetWhen(EntityRecommendation when)
        {
            When value = new When();

            var today = DateTime.Today.ToLocalTime(); //DateTime.Today;
            int dayOfWeekIndex = DayOfWeekIndex(today.DayOfWeek);

            switch (when.Entity.ToLower())
            {
                case "today":
                    value.start = DateTime.Today;
                    value.end = DateTime.Today.AddDays(1);
                    break;
                case "tomorrow":
                    value.start = DateTime.Today.AddDays(1);
                    value.end = DateTime.Today.AddDays(2);
                    break;
                case "monday":
                    value.start = today.AddDays((8 - dayOfWeekIndex) % 7);
                    value.end = today.AddDays((8 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "tuesday":
                    value.start = today.AddDays((9 - dayOfWeekIndex) % 7);
                    value.end = today.AddDays((9 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "wednesday":
                    value.start = today.AddDays((10 - dayOfWeekIndex) % 7);
                    value.end = today.AddDays((10 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "thursday":
                    value.start = today.AddDays((11 - dayOfWeekIndex) % 7);
                    value.end = today.AddDays((11 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "friday":
                    value.start = today.AddDays((12 - dayOfWeekIndex) % 7 );
                    value.end = today.AddDays((12 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "saturday":
                    value.start = today.AddDays((13 - dayOfWeekIndex) % 7);
                    value.end = today.AddDays((13 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "sunday":
                    value.start = today.AddDays((7 - dayOfWeekIndex) % 7);
                    value.end = today.AddDays((7 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "next week":
                    value.start = today.AddDays((7 - dayOfWeekIndex) % 7);
                    value.end = today.AddDays((7 - dayOfWeekIndex) % 7 + 8);
                    break;
                case "this week":
                    value.start = today;
                    value.end = today.AddDays((7 - dayOfWeekIndex) % 7 + 1);
                    break;
                default:
                    value.start = today;
                    value.end = today.AddDays(1);
                    break;
            }

            return value;
        }

        private async Task GetFocusTime(IDialogContext context, LuisResult result)
        {
            if (result.Entities.Count > 0 && result.Entities[0].Type == "When")
            {
                When availabilityDates = GetWhen(result.Entities[0]);
                var query = "https://graph.microsoft.com/v1.0/me/calendarview?startdatetime={0}&enddatetime={1}&$select=location,subject,start,end&$orderby=start/datetime&$filter=categories/any(a:a%20eq%20'Focus%20Time')";
                query = string.Format(query, availabilityDates.start.ToUniversalTime().ToString(), availabilityDates.end.ToUniversalTime().ToString());
                // save the query so we can run it after authenticating
                context.ConversationData.SetValue<string>("When", result.Entities[0].Entity);
                context.ConversationData.SetValue<string>("GraphQuery", query);
                // Initialize AuthenticationOptions with details from AAD v2 app registration (https://apps.dev.microsoft.com)
                AuthenticationOptions options = new AuthenticationOptions()
                {
                    Authority = ConfigurationManager.AppSettings["aad:Authority"],
                    ClientId = ConfigurationManager.AppSettings["aad:ClientId"],
                    ClientSecret = ConfigurationManager.AppSettings["aad:ClientSecret"],
                    Scopes = new string[] { "Files.Read", "Calendars.Read" },
                    RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
                };

                // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
                await context.Forward(new AuthDialog(new MSALAuthProvider(), options), async (IDialogContext authContext, IAwaitable<AuthResult> authResult) =>
                {
                    var tokenInfo = await authResult;

                    var json = await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, authContext.ConversationData.GetValue<string>("GraphQuery"));
                    var items = (JArray)json.SelectToken("value");
                    var reply = ((Activity)authContext.Activity).CreateReply();
                    if (items.Count > 0)
                    {
                        reply.Text = "";
                        foreach (var item in items)
                        {
                            string itemString = "* " + item.SelectToken("start").SelectToken("dateTime").Value<DateTime>().ToLocalTime().ToShortTimeString() + " to " + item.SelectToken("end").SelectToken("dateTime").Value<DateTime>().ToLocalTime().ToShortTimeString() + "\r";
                            reply.Text += itemString;
                        }
                    }
                    else
                    {
                        reply.Text = "You have no availability for " + authContext.ConversationData.GetValue<string>("When");
                    }

                    ConnectorClient client = new ConnectorClient(new Uri(authContext.Activity.ServiceUrl));
                    await client.Conversations.ReplyToActivityAsync(reply);

                }, context.Activity, CancellationToken.None);
            }
            else
            {
                await None(context, result);
            }
        }

        private List<When> FindScheduleGaps(JArray meetings, When timePeriod, bool treatFocusTimeAsAvailable)
        {
            List<When> gaps = new List<When>();
            gaps.Add(timePeriod);

            foreach (var meeting in meetings)
            {

            }

            return gaps;
        }

        [LuisIntent("GetAvailability")]
        public async Task GetAvailability(IDialogContext context, LuisResult result)
        {
            if (result.Entities.Count > 0 && result.Entities[0].Type == "When")
            {
                When availabilityDates = GetWhen(result.Entities[0]);
                var query = "https://graph.microsoft.com/v1.0/me/calendarview?startdatetime={0}&enddatetime={1}&$select=location,subject,start,end&$orderby=start/datetime";
                query = string.Format(query, availabilityDates.start.ToUniversalTime().ToString(), availabilityDates.end.ToUniversalTime().ToString());
                // save the query so we can run it after authenticating
                context.ConversationData.SetValue<string>("When", result.Entities[0].Entity);
                context.ConversationData.SetValue<string>("GraphQuery", query);
                // Initialize AuthenticationOptions with details from AAD v2 app registration (https://apps.dev.microsoft.com)
                AuthenticationOptions options = new AuthenticationOptions()
                {
                    Authority = ConfigurationManager.AppSettings["aad:Authority"],
                    ClientId = ConfigurationManager.AppSettings["aad:ClientId"],
                    ClientSecret = ConfigurationManager.AppSettings["aad:ClientSecret"],
                    Scopes = new string[] { "Files.Read", "Calendars.Read" },
                    RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
                };

                // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
                await context.Forward(new AuthDialog(new MSALAuthProvider(), options), async (IDialogContext authContext, IAwaitable<AuthResult> authResult) =>
                {
                    var tokenInfo = await authResult;

                    var json = await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, authContext.ConversationData.GetValue<string>("GraphQuery"));
                    var items = (JArray)json.SelectToken("value");

                    List<When> gaps = FindScheduleGaps(items, availabilityDates, true);


                    var reply = ((Activity)authContext.Activity).CreateReply();
                    if (gaps.Count > 0)
                    {
                        reply.Text = "";
                        foreach (var item in gaps)
                        {
                            string itemString = "* " + item.start.ToLocalTime().ToShortTimeString() + " to " + item.end.ToLocalTime().ToShortTimeString() + "\r";
                            reply.Text += itemString;
                        }
                    }
                    else
                    {
                        reply.Text = "You have no availability for " + authContext.ConversationData.GetValue<string>("When");
                    }

                    ConnectorClient client = new ConnectorClient(new Uri(authContext.Activity.ServiceUrl));
                    await client.Conversations.ReplyToActivityAsync(reply);

                }, context.Activity, CancellationToken.None);
            }
            else
            {
                await None(context, result);
            }
        }

        [LuisIntent("NextMeeting")]
        public async Task NextMeeting(IDialogContext context, LuisResult result)
        {
            await context.PostAsync("Next meeting goes here");
        }

        [LuisIntent("ScheduleTime")]
        public async Task ScheduleTime(IDialogContext context, LuisResult result)
        {
            await context.PostAsync("Schedule time goes here");
        }

        [LuisIntent("SearchFiles")]
        public async Task SearchFiles(IDialogContext context, LuisResult result)
        {
            // makes sure we got at least one entity from LUIS
            if (result.Entities.Count == 0)
                await None(context, result);
            else
            {
                var query = "https://graph.microsoft.com/v1.0/me/drive/search(q='{0}')?$select=id,name,size,webUrl&$top=5";
                // we will assume only one entity, but LUIS can handle multiple entities
                if (result.Entities[0].Type == "FileName")
                {
                    // perform a search for the filename
                    query = String.Format(query, result.Entities[0].Entity.Replace(" . ", "."));
                }
                else if (result.Entities[0].Type == "FileType")
                {
                    // perform search based on filetype...but clean up the filetype first
                    var fileType = result.Entities[0].Entity.Replace(" . ", ".").Replace(". ", ".").ToLower();
                    List<string> images = new List<string>() { "images", "pictures", "pics", "photos", "image", "picture", "pic", "photo" };
                    List<string> presentations = new List<string>() { "powerpoints", "presentations", "decks", "powerpoints", "presentation", "deck", ".pptx", ".ppt", "pptx", "ppt" };
                    List<string> documents = new List<string>() { "documents", "document", "word", "doc", ".docx", ".doc", "docx", "doc" };
                    List<string> workbooks = new List<string>() { "workbooks", "workbook", "excel", "spreadsheet", "spreadsheets", ".xlsx", ".xls", "xlsx", "xls" };
                    List<string> music = new List<string>() { "music", "songs", "albums", ".mp3", "mp3" };
                    List<string> videos = new List<string>() { "video", "videos", "movie", "movies", ".mp4", "mp4", ".mov", "mov", ".avi", "avi" };

                    if (images.Contains(fileType))
                        query = String.Format(query, ".png .jpg .jpeg .gif");
                    else if (presentations.Contains(fileType))
                        query = String.Format(query, ".pptx .ppt");
                    else if (documents.Contains(fileType))
                        query = String.Format(query, ".docx .doc");
                    else if (workbooks.Contains(fileType))
                        query = String.Format(query, ".xlsx .xls");
                    else if (music.Contains(fileType))
                        query = String.Format(query, ".mp3");
                    else if (videos.Contains(fileType))
                        query = String.Format(query, ".mp4 .avi .mov");
                    else
                        query = String.Format(query, fileType);
                }

                // save the query so we can run it after authenticating
                context.ConversationData.SetValue<string>("GraphQuery", query);
                // Initialize AuthenticationOptions with details from AAD v2 app registration (https://apps.dev.microsoft.com)
                AuthenticationOptions options = new AuthenticationOptions()
                {
                    Authority = ConfigurationManager.AppSettings["aad:Authority"],
                    ClientId = ConfigurationManager.AppSettings["aad:ClientId"],
                    ClientSecret = ConfigurationManager.AppSettings["aad:ClientSecret"],
                    Scopes = new string[] { "Files.Read" },
                    RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
                };

                // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
                await context.Forward(new AuthDialog(new MSALAuthProvider(), options), async (IDialogContext authContext, IAwaitable<AuthResult> authResult) =>
                {
                    var tokenInfo = await authResult;

                    // Get the users profile photo from the Microsoft Graph
                    var json = await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, authContext.ConversationData.GetValue<string>("GraphQuery"));
                    var items = (JArray)json.SelectToken("value");
                    var reply = ((Activity)authContext.Activity).CreateReply();
                    foreach (var item in items)
                    { 
                        // we could get thumbnails for each item using the id, but will keep it simple
                        ThumbnailCard card = new ThumbnailCard()
                        {
                            Title = item.Value<string>("name"),
                            Subtitle = $"Size: {item.Value<int>("size").ToString()}",
                            Text = $"Download: {item.Value<string>("webUrl")}"
                        };
                        reply.Attachments.Add(card.ToAttachment());
                    }

                    reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                    ConnectorClient client = new ConnectorClient(new Uri(authContext.Activity.ServiceUrl));
                    await client.Conversations.ReplyToActivityAsync(reply);

                }, context.Activity, CancellationToken.None);
            }
        }
    }
}