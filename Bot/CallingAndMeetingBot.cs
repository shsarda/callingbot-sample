using CallingMeetingBot.Authentication;
using CallingMeetingBot.Extenstions;
using CallingMeetingBot.Utility;
using Microsoft.AspNetCore.Http;
using Microsoft.Graph.Communications.Client.Authentication;
using Microsoft.Graph.Communications.Common.Telemetry;
using Microsoft.Graph.Communications.Core.Notifications;
using Microsoft.Graph.Communications.Core.Serialization;
using System;
using System.Net;
using System.Threading.Tasks;
using System.Net.Http;
using Microsoft.Graph;
using System.Diagnostics;
using System.Collections.Generic;
using Microsoft.Graph.Communications.Common.Transport;
using System.Net.Http.Headers;
using System.Linq;
using Microsoft.Graph.Communications.Client.Transport;
using System.Text;
using CallingMeetingBot.Model;

namespace CallingMeetingBot.Bot
{
    public class CallingAndMeetingBot
    {
        private readonly BotOptions options;

        public IGraphLogger GraphLogger { get; }

        private IRequestAuthenticationProvider AuthenticationProvider { get; }

        private INotificationProcessor NotificationProcessor { get; }

        private CommsSerializer Serializer { get; }

        private GraphServiceClient RequestBuilder { get; }

        private GraphServiceClient GraphClient_DelegatedAuth { get; set; }

        private IGraphClient GraphApiClient { get; }

        public CallingAndMeetingBot(BotOptions options, IGraphLogger graphLogger)
        {
            this.options = options;
            this.GraphLogger = graphLogger;
            var name = this.GetType().Assembly.GetName().Name;
            this.AuthenticationProvider = new AuthenticationProvider(name, options.AppId, options.AppSecret, graphLogger);

            this.Serializer = new CommsSerializer();
            var authenticationWrapper = new AuthenticationWrapper(this.AuthenticationProvider);
            this.NotificationProcessor = new NotificationProcessor(Serializer);
            this.NotificationProcessor.OnNotificationReceived += this.NotificationProcessor_OnNotificationReceived;

            this.RequestBuilder = new GraphServiceClient(options.PlaceCallEndpointUrl.AbsoluteUri, authenticationWrapper);

            var defaultProperties = new List<IGraphProperty<IEnumerable<string>>>();
            using (HttpClient tempClient = GraphClientFactory.Create(authenticationWrapper))
            {
                defaultProperties.AddRange(tempClient.DefaultRequestHeaders.Select(header => GraphProperty.RequestProperty(header.Key, header.Value)));
            }

            // graph client
            var productInfo = new ProductInfoHeaderValue(
                typeof(CallingAndMeetingBot).Assembly.GetName().Name,
                typeof(CallingAndMeetingBot).Assembly.GetName().Version.ToString());
            this.GraphApiClient = new GraphAuthClient(
                this.GraphLogger,
                this.Serializer.JsonSerializerSettings,
                new HttpClient(),
                this.AuthenticationProvider,
                productInfo,
                defaultProperties);
        }

        /// <summary>
        /// Process "/callback" notifications asyncronously. 
        /// </summary>
        /// <param name="request"></param>
        /// <param name="response"></param>
        /// <returns></returns>
        public async Task ProcessNotificationAsync(
            HttpRequest request,
            HttpResponse response)
        {
            try
            {
                var httpRequest = request.CreateRequestMessage();
                var results = await this.AuthenticationProvider.ValidateInboundRequestAsync(httpRequest).ConfigureAwait(false);
                if (results.IsValid)
                {
                    var httpResponse = await this.NotificationProcessor.ProcessNotificationAsync(httpRequest).ConfigureAwait(false);
                    await httpResponse.CreateHttpResponseAsync(response).ConfigureAwait(false);
                }
                else
                {
                    var httpResponse = httpRequest.CreateResponse(HttpStatusCode.Forbidden);
                    await httpResponse.CreateHttpResponseAsync(response).ConfigureAwait(false);
                }


            }
            catch (Exception e)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await response.WriteAsync(e.ToString()).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Called when INotificationProcessor recieves notification.
        /// </summary>
        /// <param name="args"></param>
        private void NotificationProcessor_OnNotificationReceived(NotificationEventArgs args)
        {
            _ = NotificationProcessor_OnNotificationReceivedAsync(args).ForgetAndLogExceptionAsync(
                this.GraphLogger,
                $"Error processing notification {args.Notification.ResourceUrl} with scenario {args.ScenarioId}");
        }

        private async Task NotificationProcessor_OnNotificationReceivedAsync(NotificationEventArgs args)
        {
            this.GraphLogger.CorrelationId = args.ScenarioId;
            if (args.ResourceData is Call call)
            {
                if (args.ChangeType == ChangeType.Created && call.State == CallState.Incoming)
                {
                    await this.BotAnswerIncomingCallAsync(call.Id, args.TenantId, args.ScenarioId).ConfigureAwait(false);
                    
                    // graph client with delegated auth provider
                    GraphClient_DelegatedAuth = GraphUtils.CreateGraphServiceClient_DelegatedAuth(options, args.TenantId);
                }
            }
            else if (args.Notification.ResourceUrl.EndsWith("/participants") && args.ResourceData is List<object> participantObjects)
            {
                this.GraphLogger.Log(TraceLevel.Info, "Total count of participants found in this roster is " + participantObjects.Count());

                await GetListOfParticipantsInCall(participantObjects, args.TenantId);
            }
        }

        private async Task BotAnswerIncomingCallAsync(string callId, string tenantId, Guid scenarioId)
        {
            var answerRequest = this.RequestBuilder.Communications.Calls[callId].Answer(
                callbackUri: new Uri(options.BotBaseUrl, "callback").ToString(),
                mediaConfig: new ServiceHostedMediaConfig
                {
                    PreFetchMedia = new List<MediaInfo>()
                    {
                        new MediaInfo()
                        {
                            Uri = new Uri(options.BotBaseUrl, "audio/speech.wav").ToString(),
                            ResourceId = Guid.NewGuid().ToString(),
                        }
                    }
                },
                acceptedModalities: new List<Modality> { Modality.Audio }).Request();
            await this.GraphApiClient.SendAsync(answerRequest, RequestType.Create, tenantId, scenarioId).ConfigureAwait(false);
        }

        private async Task<List<ParticipantDetails>> GetListOfParticipantsInCall(List<object> participantObjects, string argsTenantId)
        {
            var participantDetailsList = new HashSet<ParticipantDetails>();
            foreach (var participantObject in participantObjects)
            {
                var participant = participantObject as Participant;
                var participantDetailsObject = new ParticipantDetails();

                // Identity User object for bot is null
                if (participant?.Info?.Identity?.User != null) 
                {
                    string aadUserId = participant.Info.Identity.User.Id;
                    string tenantId = (string)participant.Info.Identity.User.AdditionalData["tenantId"];
                    string participantName = participant.Info.Identity.User.DisplayName.ToString();

                    participantDetailsObject.Name = participantName;
                    participantDetailsObject.UserId = aadUserId;
                    participantDetailsObject.TenantId = tenantId;
                    try
                    {
                        if (tenantId != null && tenantId == argsTenantId)
                        {
                            var response = await GraphClient_DelegatedAuth.Users[aadUserId].Request().GetAsync();
                            string participantMail = response.Mail;
                            participantDetailsObject.EmailId = participantMail;
                        }
                    }
                    catch (Exception e)
                    {
                        this.GraphLogger.Log(TraceLevel.Error, "Error occured while fetching User Info for participant Id : " + participant.Id + " Error " + e.Message);
                    }
                }
                else if (participant?.Info?.Identity?.AdditionalData != null)
                {
                    if (participant.Info.Identity.AdditionalData.ContainsKey("guest") && participant.IsInLobby == false)
                    {
                        Identity guestDisplayName = (Identity)participant.Info.Identity.AdditionalData["guest"];
                        participantDetailsObject.Name = guestDisplayName.DisplayName + " (Guest)\n";
                    }
                    else if (participant.Info.Identity.AdditionalData.ContainsKey("phone") && participant.IsInLobby == false)
                    {
                        Identity phoneDisplayName = (Identity)participant.Info.Identity.AdditionalData["phone"];
                        participantDetailsObject.Name = phoneDisplayName.DisplayName + " (Phone)\n";
                    }
                }
                if (participantDetailsObject.Name != null)
                {
                    participantDetailsList.Add(participantDetailsObject);
                }
            }
            return participantDetailsList.ToList();
        }
    }
}
