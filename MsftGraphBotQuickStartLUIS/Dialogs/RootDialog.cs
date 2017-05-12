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
    [LuisModel("c138dfee-d228-416f-953e-5307c2b236ac", "25670274c6ad4180a3e4be0e8f4d91d5")]
    [Serializable]
    public class RootDialog : LuisDialog<IMessageActivity>
    {
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            await context.PostAsync("I didn't understand your query...I'm just a simple bot that searches OneDrive. Try a query similar to these:<br/>'find all music'<br/>'find all .pptx files'<br/>'search for mydocument.docx'");
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