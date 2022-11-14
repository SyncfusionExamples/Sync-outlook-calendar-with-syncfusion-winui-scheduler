using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.UI.Xaml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using static System.Formats.Asn1.AsnWriter;


// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace SyncOutlookCalendar
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        private GraphServiceClient Client = null;
        public MainWindow()
        {
            this.InitializeComponent();
            this.Authenticate();
        }

        private async void Authenticate()
        {
            AuthenticationResult tokenRequest;
            var accounts = await App.ClientApplication.GetAccountsAsync();
            if (accounts.Count() > 0)
            {
                tokenRequest = await App.ClientApplication.AcquireTokenSilent(App.Scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }
            else
            {

                tokenRequest = await App.ClientApplication.AcquireTokenInteractive(App.Scopes).ExecuteAsync();
            }

            Client = new GraphServiceClient("https://graph.microsoft.com/v1.0/",
                                new DelegateAuthenticationProvider(async (requestMessage) =>
                                {
                                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                                }));

            this.AddCalendar();
        }

        /// <summary>
        /// 
        /// </summary>
        private void AddCalendar()
        {
            if (Client != null)
            {
                var calEvent = new Event
                {
                    Subject = "Azure",

                    Start = new DateTimeTimeZone
                    {
                        DateTime = DateTime.Now.ToString(),
                        TimeZone = "GMT Standard Time"
                    },
                    End = new DateTimeTimeZone()
                    {
                        DateTime = DateTime.Now.AddHours(1).ToString(),
                        TimeZone = "GMT Standard Time"
                    },
                };

                Client.Me.Events.Request().AddAsync(calEvent);
            }
        }
    }

}
