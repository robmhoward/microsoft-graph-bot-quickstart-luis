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

        #region configuration properties

        private static TimeZoneInfo defaultTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");

        private AuthenticationOptions authenticationOptions = new AuthenticationOptions()
        {
            Authority = ConfigurationManager.AppSettings["aad:Authority"],
            ClientId = ConfigurationManager.AppSettings["aad:ClientId"],
            ClientSecret = ConfigurationManager.AppSettings["aad:ClientSecret"],
            Scopes = new string[] { "Files.Read", "Calendars.ReadWrite", "Calendars.Read.Shared", "MailboxSettings.ReadWrite", "People.Read" },
            RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
        };

        #endregion

        #region classes 

        private class When
        {
            public DateTime start;
            public DateTime end;
        }

        [Serializable]
        public class Person
        {
            public string displayName { get; set; }
            public string id { get; set; }
            public string email { get; set; }
            public override string ToString()
            {
                return displayName;
            }
        }

        #endregion

        #region helper functions


        private async Task SearchForPeople(IDialogContext context, IAwaitable<AuthResult> authResult)
        {
            var tokenInfo = await authResult;
            List<string> peopleToMeetWith = context.ConversationData.GetValue<List<string>>("WhoToLookup");


            var personQuery = string.Format("https://graph.microsoft.com/beta/me/people?$search={0}", peopleToMeetWith[0]);
            var json = await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, personQuery);
            var items = (JArray)json.SelectToken("value");
            List<Person> peopleChoices = new List<Person>();

            foreach (var item in items)
            {
                Person person = new Person();
                person.displayName = item.Value<string>("displayName");
                person.email = item.SelectToken("emailAddresses")[0].Value<string>("address");
                person.id = item.Value<string>("id");

                peopleChoices.Add(person);
            }

            string promptText = string.Format("Which '{0}'?", peopleToMeetWith[0]);
            peopleToMeetWith.RemoveAt(0);
            context.ConversationData.SetValue<List<string>>("WhoToLookup", peopleToMeetWith);
            PromptDialog.Choice<Person>(context, this.ChoosePerson, peopleChoices, promptText);
        }

        private async Task ChoosePerson(IDialogContext context, IAwaitable<Person> argument)
        {
            var person = await argument;
            List<Person> whoToSchedule = context.ConversationData.GetValue<List<Person>>("WhoToSchedule");
            whoToSchedule.Add(person);
            context.ConversationData.SetValue<List<Person>>("WhoToSchedule", whoToSchedule);

            List<string> whoToLookup = context.ConversationData.GetValue<List<string>>("WhoToLookup");
            if (whoToLookup.Count > 0)
            {
                await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), SearchForPeople, context.Activity, CancellationToken.None);
            }
            else
            {
                await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), ScheduleTime, context.Activity, CancellationToken.None);
            }

        }

        // Takes a JArray representing meetings and a When describing a time period; returns a list of available times during the When 
        private List<When> FindScheduleGaps(JArray meetings, When timePeriod, bool treatFocusTimeAsAvailable)
        {
            List<When> gaps = FindWorkingHoursInTimePeriod(timePeriod);

            foreach (var meeting in meetings)
            {
                When fill = new When();
                fill.start = DateTime.SpecifyKind(meeting.SelectToken("start").SelectToken("dateTime").Value<DateTime>(), DateTimeKind.Utc);
                fill.end = DateTime.SpecifyKind(meeting.SelectToken("end").SelectToken("dateTime").Value<DateTime>(), DateTimeKind.Utc);
                bool isFocusTime = false;

                if (treatFocusTimeAsAvailable)
                {
                    IEnumerable<string> categories = meeting.SelectToken("categories").Values<string>();

                    foreach (var category in categories)
                    {
                        if (category == "Focus Time")
                        {
                            isFocusTime = true;
                            break;
                        }
                    }
                }
                List<When> gapsToAdd = new List<When>();
                List<When> gapsToRemove = new List<When>();
                if (!isFocusTime)
                {
                    foreach (var gap in gaps)
                    {
                        if (fill.start < gap.end && fill.end > gap.start)
                        {
                            // Fill covers full gap
                            if (fill.start <= gap.start && fill.end >= gap.end)
                            {
                                gapsToRemove.Add(gap);
                            }
                            // Fill starts before gap starts and ends before gap ends
                            else if (fill.start <= gap.start && fill.end < gap.end)
                            {
                                gap.start = fill.end;
                            }
                            // Fill starts after gap starts and ends after gap ends 
                            else if (fill.start > gap.start && fill.end >= gap.end)
                            {
                                gap.end = fill.start;
                            }
                            //Fill falls in the middle of the gap
                            else if (fill.start > gap.start && fill.end < gap.end)
                            {
                                When newGap = new When();
                                newGap.end = gap.end;
                                newGap.start = fill.end;
                                gap.end = fill.start;
                                gapsToAdd.Add(newGap);
                            }
                        }
                    }

                    // add and remove all the gaps at the end, after we're done iterating through the collection
                    foreach (var gap in gapsToRemove)
                    {
                        gaps.Remove(gap);
                    }

                    foreach (var gap in gapsToAdd)
                    {
                        gaps.Add(gap);
                    }
                }
            }

            return gaps;
        }

        // Takes a When describing a potentially multi-day period, and returns a list of Whens that represent working hours within that time
        private List<When> FindWorkingHoursInTimePeriod(When timePeriod)
        {
            List<When> results = new List<When>();
            for (var i = 0; i < (timePeriod.end - timePeriod.start).Days; i++)
            {
                When workingHours = new When();
                workingHours.start = timePeriod.start.AddDays(i).AddHours(8);
                workingHours.end = workingHours.start.AddHours(9);
                results.Add(workingHours);
            }

            return results;
        }

        // Returns an index representing a DayOfWeek as an integer, starting with 0 for Sunday
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

        // Takes a list of UTC Whens and returns a reply expressing those Whens with a short time string
        private Activity CreateReplyFromWhenList(List<When> whenList, string whenString, IDialogContext authContext, TimeZoneInfo timeZoneInfo)
        {
            var reply = ((Activity)authContext.Activity).CreateReply();

            if (whenList.Count > 0)
            {
                reply.Text = "";
                foreach (var item in whenList)
                {
                    string itemString = "* " + TimeZoneInfo.ConvertTimeFromUtc(item.start, timeZoneInfo).ToShortTimeString() + " to " + TimeZoneInfo.ConvertTimeFromUtc(item.end, timeZoneInfo).ToShortTimeString() + "\r";
                    reply.Text += itemString;
                }
            }
            else
            {
                reply.Text = "You have no availability for " + whenString;
            }

            return reply;
        }

        // Takes a string describing "when" and the TimeZoneInfo of that description and returns a When object with UTC start and end DateTime
        private When GetWhen(string when, TimeZoneInfo timeZoneInfoOfWhen)
        {
            When value = new When();
            var today = (TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, timeZoneInfoOfWhen)).Date;
            int dayOfWeekIndex = DayOfWeekIndex(today.DayOfWeek);
            var utcStartOfToday = TimeZoneInfo.ConvertTimeToUtc(today, timeZoneInfoOfWhen);

            switch (when.ToLower())
            {
                case "today":
                    value.start = utcStartOfToday;
                    value.end = utcStartOfToday.AddDays(1);
                    break;
                case "tomorrow":
                    value.start = utcStartOfToday.AddDays(1);
                    value.end = utcStartOfToday.AddDays(2);
                    break;
                case "monday":
                    value.start = utcStartOfToday.AddDays((8 - dayOfWeekIndex) % 7);
                    value.end = utcStartOfToday.AddDays((8 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "tuesday":
                    value.start = utcStartOfToday.AddDays((9 - dayOfWeekIndex) % 7);
                    value.end = utcStartOfToday.AddDays((9 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "wednesday":
                    value.start = utcStartOfToday.AddDays((10 - dayOfWeekIndex) % 7);
                    value.end = utcStartOfToday.AddDays((10 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "thursday":
                    value.start = utcStartOfToday.AddDays((11 - dayOfWeekIndex) % 7);
                    value.end = utcStartOfToday.AddDays((11 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "friday":
                    value.start = utcStartOfToday.AddDays((12 - dayOfWeekIndex) % 7);
                    value.end = utcStartOfToday.AddDays((12 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "saturday":
                    value.start = utcStartOfToday.AddDays((13 - dayOfWeekIndex) % 7);
                    value.end = utcStartOfToday.AddDays((13 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "sunday":
                    value.start = utcStartOfToday.AddDays((7 - dayOfWeekIndex) % 7);
                    value.end = utcStartOfToday.AddDays((7 - dayOfWeekIndex) % 7 + 1);
                    break;
                case "next week":
                    value.start = utcStartOfToday.AddDays((7 - dayOfWeekIndex) % 7);
                    value.end = utcStartOfToday.AddDays((7 - dayOfWeekIndex) % 7 + 8);
                    break;
                case "this week":
                    value.start = utcStartOfToday;
                    value.end = utcStartOfToday.AddDays((7 - dayOfWeekIndex) % 7 + 1);
                    break;
                default:
                    value.start = utcStartOfToday;
                    value.end = utcStartOfToday.AddDays(1);
                    break;
            }

            return value;
        }

        private async Task<TimeZoneInfo> GetTimeZoneInfo(AuthResult tokenInfo, IDialogContext authContext)
        {
            TimeZoneInfo timeZoneInfo = null;
            string timeZoneName = null;

            if (!authContext.ConversationData.TryGetValue<string>("TimeZoneName", out timeZoneName)) {
                var mailboxSettingsJson = await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, "https://graph.microsoft.com/v1.0/me?$select=mailboxSettings");
                timeZoneName = mailboxSettingsJson.SelectToken("mailboxSettings").SelectToken("timeZone").Value<string>();
                timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(timeZoneName);
                authContext.ConversationData.SetValue<string>("TimeZoneName", timeZoneName);
            } else
            {
                timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(timeZoneName);
            }

            return timeZoneInfo;
        }

        #endregion

        #region intents

        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            await context.PostAsync("I didn't understand your query.  Here are a few things I know how to do:<br/>'find all music'<br/>'find all .pptx files'<br/>'search for mydocument.docx'<br/>where's my next meeting?<br/>when am I available on Tuesday?<br/>block my calendar tomorrow");
        }

        [LuisIntent("BlockCalendar")]
        public async Task BlockCalendar(IDialogContext context, LuisResult result)
        {
            await context.PostAsync("Block calendar goes here");
        }

        [LuisIntent("GetAvailability")]
        public async Task GetAvailability(IDialogContext context, LuisResult result)
        {
            if (result.Entities.Count > 0 && result.Entities[0].Type == "When")
            {
                
                // save the query so we can run it after authenticating
                context.ConversationData.SetValue<string>("When", result.Entities[0].Entity);

                // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
                await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), async (IDialogContext authContext, IAwaitable<AuthResult> authResult) =>
                {
                    var tokenInfo = await authResult;
                    string when = authContext.ConversationData.GetValue<string>("When");

                    TimeZoneInfo timeZoneInfo = await GetTimeZoneInfo(tokenInfo, authContext);

                    When availabilityDates = GetWhen(when, timeZoneInfo);

                    var calendarQuery = "https://graph.microsoft.com/v1.0/me/calendarview?startdatetime={0}&enddatetime={1}&$select=location,subject,start,end,categories&$orderby=start/datetime&$filter=showAs%20eq%20'busy'";
                    calendarQuery = string.Format(calendarQuery, availabilityDates.start.ToString(), availabilityDates.end.ToString());
                    var items = (JArray)(await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, calendarQuery)).SelectToken("value");

                    List<When> gaps = FindScheduleGaps(items, availabilityDates, true);

                    var reply = CreateReplyFromWhenList(gaps, when, authContext, timeZoneInfo);

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
            var query = "https://graph.microsoft.com/v1.0/me/calendarview?startdatetime={0}&enddatetime={1}&$top=1&$select=location,subject,start&$orderby=start/datetime&$filter=showAs%20eq%20'busy'";
            query = string.Format(query, DateTime.UtcNow.ToString(), DateTime.UtcNow.AddDays(1).ToString());
            // save the query so we can run it after authenticating
            context.ConversationData.SetValue<string>("GraphQuery", query);

            // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
            await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), async (IDialogContext authContext, IAwaitable<AuthResult> authResult) =>
            {
                var tokenInfo = await authResult;

                TimeZoneInfo timeZoneInfo = await GetTimeZoneInfo(tokenInfo, authContext);

                var json = await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, authContext.ConversationData.GetValue<string>("GraphQuery"));

                var nextMeeting = ((JArray)json.SelectToken("value"))[0];
                var responseText = string.Format("Your next meeting '{0}' is at {1} in {2}",
                    nextMeeting.SelectToken("subject").Value<string>(),
                    TimeZoneInfo.ConvertTimeFromUtc(nextMeeting.SelectToken("start").SelectToken("dateTime").Value<DateTime>(), timeZoneInfo).ToShortTimeString(),
                    nextMeeting.SelectToken("location").SelectToken("displayName").Value<string>());

                var reply = ((Activity)authContext.Activity).CreateReply(responseText);

                ConnectorClient client = new ConnectorClient(new Uri(authContext.Activity.ServiceUrl));
                await client.Conversations.ReplyToActivityAsync(reply);

            }, context.Activity, CancellationToken.None);
        }

        [LuisIntent("ScheduleTime")]
        public async Task ScheduleTime(IDialogContext context, LuisResult result)
        {
            if (result.Entities.Count > 0 && result.Entities[0].Type == "Person")
            {

                List<string> peopleToMeetWith = new List<string>();
                foreach(var entityRecommendation in result.Entities)
                {
                    if (entityRecommendation.Type == "Person")
                    {
                        peopleToMeetWith.Add(entityRecommendation.Entity);
                    }
                }

                // save the query so we can run it after authenticating
                context.ConversationData.SetValue<List<string>>("WhoToLookup", peopleToMeetWith);
                context.ConversationData.SetValue<List<Person>>("WhoToSchedule", new List<Person>());

                // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
                await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), SearchForPeople, context.Activity, CancellationToken.None);

            }
            else
            {
                await None(context, result);
            }
        }


        public class FindMeetingTimesBody
        {
            public List<Attendee> attendees { get; set; }
            public string meetingDuration { get; set; }
        }

        public class Attendee
        {
            public EmailAddress emailAddress { get; set; }
            public string type { get; set; }
        }

        public class EmailAddress
        {
            public string address { get; set; }
        }

        private async Task ScheduleTime(IDialogContext context, IAwaitable<AuthResult> authResult)
        {
            var tokenInfo = await authResult;

            List<Person> whoToSchedule = context.ConversationData.GetValue<List<Person>>("WhoToSchedule");
            List<Attendee> attendees = new List<Attendee>();

            foreach (var person in whoToSchedule)
            {
                Attendee attendee = new Attendee();
                EmailAddress email = new EmailAddress();
                email.address = person.email;
                attendee.emailAddress = email;
                attendee.type = "required";
                attendees.Add(attendee);
            }

            FindMeetingTimesBody body = new FindMeetingTimesBody();
            body.attendees = attendees;
            body.meetingDuration = "PT1H";

            TimeZoneInfo timeZoneInfo = await GetTimeZoneInfo(tokenInfo, context);

            var json = await new HttpClient().PostWithAuthAsync<FindMeetingTimesBody>(tokenInfo.AccessToken, "https://graph.microsoft.com/v1.0/me/findMeetingTimes", body);
            var items = (JArray)json.SelectToken("meetingTimeSuggestions");

            List<MeetingSuggestion> timeChoices = new List<MeetingSuggestion>();

            foreach (var item in items)
            {
                MeetingSuggestion suggestion = new MeetingSuggestion();
                suggestion.Start = item.SelectToken("meetingTimeSlot.start.dateTime").Value<DateTime>();
                suggestion.End= item.SelectToken("meetingTimeSlot.end.dateTime").Value<DateTime>();

                var availabilityString = "";
                var availability = (JArray)item.SelectToken("attendeeAvailability");
                foreach (var availabilityItem in availability)
                {
                    var status = availabilityItem.SelectToken("availability").Value<string>();
                    if (status != "free")
                    {
                        availabilityString += $" {availabilityItem.SelectToken("attendee.emailAddress.address").Value<string>()} is {status} "; 
                    }
                }

                if (availabilityString == "")
                {
                    availabilityString = "everyone is available";
                }

                suggestion.Description = $"{TimeZoneInfo.ConvertTimeFromUtc(suggestion.Start, timeZoneInfo).ToString()} {availabilityString}"; 
                timeChoices.Add(suggestion);
            }

            PromptDialog.Choice<MeetingSuggestion>(context, this.ChooseTime, timeChoices, "Which time would you like to choose?");
        }

        private async Task ChooseTime(IDialogContext context, IAwaitable<MeetingSuggestion> argument)
        {
            context.ConversationData.SetValue<MeetingSuggestion>("SchedulingSuggestion", await argument);

            // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
            await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), this.ScheduleMeeting, context.Activity, CancellationToken.None);

        }

        private async Task ScheduleMeeting(IDialogContext context, IAwaitable<AuthResult> authResult)
        {

            var tokenInfo = await authResult;

            TimeZoneInfo timeZoneInfo = await GetTimeZoneInfo(tokenInfo, context);
            MeetingSuggestion suggestion = context.ConversationData.GetValue<MeetingSuggestion>("SchedulingSuggestion");
            List<Person> whoToSchedule = context.ConversationData.GetValue<List<Person>>("WhoToSchedule");
            List<Attendee> attendees = new List<Attendee>();

            foreach (var person in whoToSchedule)
            {
                Attendee attendee = new Attendee();
                EmailAddress email = new EmailAddress();
                email.address = person.email;
                attendee.emailAddress = email;
                attendee.type = "required";
                attendees.Add(attendee);
            }

            MeetingBody body = new MeetingBody();
            body.subject = "Testing - please ignore";
            body.body = new EmailBody();
            body.body.contentType = "HTML";
            body.body.content = "Meeting scheduled with Rob's Demo Bot. This is probably for a demo.  Sorry for the spam :).";
            body.start = new MeetingTime();
            body.start.timeZone = timeZoneInfo.StandardName;
            body.start.dateTime = TimeZoneInfo.ConvertTimeFromUtc(suggestion.Start, timeZoneInfo).ToString();
            body.end = new MeetingTime();
            body.end.timeZone = timeZoneInfo.StandardName;
            body.end.dateTime = TimeZoneInfo.ConvertTimeFromUtc(suggestion.End, timeZoneInfo).ToString();
            body.location = new EmailLocation();
            body.location.displayName = "TBD";
            body.attendees = attendees;

            var json = await new HttpClient().PostWithAuthAsync<MeetingBody>(tokenInfo.AccessToken, "https://graph.microsoft.com/v1.0/me/events", body);
            if (json != null)
            {
                await context.PostAsync($"Sent an invite to your meeting at {suggestion}.");
            } else
            {
                await context.PostAsync($"It looks like something went wrong.");
            }
            
        }

        public class MeetingBody
        {
            public string subject { get; set; }
            public EmailBody body { get; set; }
            public EmailLocation location { get; set; }
            public MeetingTime start { get; set; }
            public MeetingTime end { get; set; }
            public List<Attendee> attendees { get; set; }
        }

        public class EmailBody
        {
            public string contentType { get; set; }
            public string content { get; set; }
        }

        public class EmailLocation
        {
            public string displayName { get; set; }
        }

        public class MeetingTime
        {
            public string dateTime { get; set; }
            public string timeZone { get; set; }
        }

        [Serializable]
        public class MeetingSuggestion
        {
            public DateTime Start { get; set; }
            public DateTime End { get; set; }
            public string Description { get; set; }
            public override string ToString()
            {
                return this.Description;
            }
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

                // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
                await context.Forward(new AuthDialog(new MSALAuthProvider(), authenticationOptions), async (IDialogContext authContext, IAwaitable<AuthResult> authResult) =>
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

        #endregion 
    }
}