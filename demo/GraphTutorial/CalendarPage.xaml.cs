// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using Microsoft.Toolkit.Graph.Providers;
using Microsoft.Toolkit.Uwp.UI.Controls;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=234238

namespace GraphTutorial
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class CalendarPage : Page
    {
        public CalendarPage()
        {
            this.InitializeComponent();
        }

        private void ShowNotification(string message)
        {
            // Get the main page that contains the InAppNotification
            var mainPage = (Window.Current.Content as Frame).Content as MainPage;

            // Get the notification control
            var notification = mainPage.FindName("Notification") as InAppNotification;

            notification.Show(message);
        }

        protected override async void OnNavigatedTo(NavigationEventArgs e)
        {
            // Get the Graph client from the provider
            var graphClient = ProviderManager.Instance.GlobalProvider.Graph;

            try
            {
                // Get the user's mailbox settings to determine
                // their time zone
                var user = await graphClient.Me.Request()
                    .Select(u => new { u.MailboxSettings })
                    .GetAsync();

                var startOfWeek = GetUtcStartOfWeekInTimeZone(DateTime.Today, user.MailboxSettings.TimeZone);
                var endOfWeek = startOfWeek.AddDays(7);

                var queryOptions = new List<QueryOption>
                {
                    new QueryOption("startDateTime", startOfWeek.ToString("o")),
                    new QueryOption("endDateTime", endOfWeek.ToString("o"))
                };

                // Get the events
                var events = await graphClient.Me.CalendarView.Request(queryOptions)
                    .Header("Prefer", $"outlook.timezone=\"{user.MailboxSettings.TimeZone}\"")
                    .Select(ev => new
                    {
                        ev.Subject,
                        ev.Organizer,
                        ev.Start,
                        ev.End
                    })
                    .OrderBy("start/dateTime")
                    .Top(50)
                    .GetAsync();

                EventList.ItemsSource = events.CurrentPage.ToList();
            }
            catch (ServiceException ex)
            {
                ShowNotification($"Exception getting events: {ex.Message}");
            }

            base.OnNavigatedTo(e);
        }

        private static DateTime GetUtcStartOfWeekInTimeZone(DateTime today, string timeZoneId)
        {
            TimeZoneInfo userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(timeZoneId);

            // Assumes Sunday as first day of week
            int diff = System.DayOfWeek.Sunday - today.DayOfWeek;

            // create date as unspecified kind
            var unspecifiedStart = DateTime.SpecifyKind(today.AddDays(diff), DateTimeKind.Unspecified);

            // convert to UTC
            return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, userTimeZone);
        }
    }
}
