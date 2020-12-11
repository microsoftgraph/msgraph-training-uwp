// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

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

// <UsingStatementsSnippet>
using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Graph.Extensions;
using Microsoft.Toolkit.Graph.Providers;
using Microsoft.Toolkit.Uwp.UI.Controls;
using System.Runtime.CompilerServices;
// </UsingStatementsSnippet>

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=234238

namespace GraphTutorial
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class NewEventPage : Page, INotifyPropertyChanged
    {
        // <NotifyPropertyChangedSnippet>
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            // Reevaluate if required fields are present and dates are valid
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("IsValid"));
        }
        // </NotifyPropertyChangedSnippet>

        // <PropertiesSnippet>
        private TimeZoneInfo _userTimeZone = null;

        // Value of the Subject text box
        private string _subject = "";
        public string Subject
        {
            get { return _subject; }
            set
            {
                _subject = value;
                OnPropertyChanged();
            }
        }

        // Value of the Start date picker
        private DateTimeOffset _startDate = DateTimeOffset.Now;
        public DateTimeOffset StartDate
        {
            get { return _startDate; }
            set
            {
                _startDate = value;
                OnPropertyChanged();
            }
        }

        // Value of the Start time picker
        private TimeSpan _startTime = TimeSpan.Zero;
        public TimeSpan StartTime
        {
            get { return _startTime; }
            set
            {
                _startTime = value;
                OnPropertyChanged();
            }
        }

        // Value of the End date picker
        private DateTimeOffset _endDate = DateTimeOffset.Now;
        public DateTimeOffset EndDate
        {
            get { return _endDate; }
            set
            {
                _endDate = value;
                OnPropertyChanged();
            }
        }

        // Value of the End time picker
        private TimeSpan _endTime = TimeSpan.Zero;
        public TimeSpan EndTime
        {
            get { return _endTime; }
            set
            {
                _endTime = value;
                OnPropertyChanged();
            }
        }

        // Value of the Body text box
        private string _body = "";
        public string Body
        {
            get { return _body; }
            set
            {
                _body = value;
                OnPropertyChanged();
            }
        }

        // Combine the date from date picker with time from time picker
        private DateTimeOffset CombineDateAndTime(DateTimeOffset date, TimeSpan time)
        {
            // Use the year, month, and day from the supplied DateTimeOffset
            // to create a new DateTime at midnight
            var dt = new DateTime(date.Year, date.Month, date.Day);

            // Add the TimeSpan, and use the user's timezone offset
            return new DateTimeOffset(dt + time, _userTimeZone.BaseUtcOffset);
        }

        // Combined value of Start date and time pickers
        public DateTimeOffset Start
        {
            get
            {
                return CombineDateAndTime(StartDate, StartTime);
            }
        }

        // Combined value of End date and time pickers
        public DateTimeOffset End
        {
            get
            {
                return CombineDateAndTime(EndDate, EndTime);
            }
        }

        public bool IsValid
        {
            get
            {
                // Subject is required, Start must be before
                // End
                return !string.IsNullOrWhiteSpace(Subject) &&
                       DateTimeOffset.Compare(Start, End) < 0;
            }
        }
        // </PropertiesSnippet>

        public NewEventPage()
        {
            InitializeComponent();
            DataContext = this;
        }

        // <LoadTimeZoneSnippet>
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

                _userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(user.MailboxSettings.TimeZone);
            }
            catch (ServiceException graphException)
            {
                ShowNotification($"Exception getting user's mailbox settings from Graph, defaulting to local time zone: {graphException.Message}");
                _userTimeZone = TimeZoneInfo.Local;
            }
            catch (Exception ex)
            {
                ShowNotification($"Exception loading time zone from system: {ex.Message}");
            }
        }

        private void ShowNotification(string message)
        {
            // Get the main page that contains the InAppNotification
            var mainPage = (Window.Current.Content as Frame).Content as MainPage;

            // Get the notification control
            var notification = mainPage.FindName("Notification") as InAppNotification;

            notification.Show(message);
        }
        // </LoadTimeZoneSnippet>

        // <CreateEventSnippet>
        private async void CreateEvent(object sender, RoutedEventArgs e)
        {
            CreateProgress.IsActive = true;
            CreateProgress.Visibility = Visibility.Visible;

            // Get the Graph client from the provider
            var graphClient = ProviderManager.Instance.GlobalProvider.Graph;

            // Initialize a new Event object with the required fields
            var newEvent = new Event
            {
                Subject = Subject,
                Start = Start.ToDateTimeTimeZone(_userTimeZone),
                End = End.ToDateTimeTimeZone(_userTimeZone)
            };

            // If there's a body, add it
            if (!string.IsNullOrEmpty(Body))
            {
                newEvent.Body = new ItemBody
                {
                    Content = Body,
                    ContentType = BodyType.Text
                };
            }

            // Add any attendees
            var peopleList = AttendeePicker.Items.ToList();
            var attendeeList = new List<Attendee>();

            foreach (object entry in peopleList)
            {
                // People Picker can contain Microsoft.Graph.Person objects
                // or text entries (for email addresses not found in the user's people list)
                if (entry is Person)
                {
                    var person = entry as Person;

                    attendeeList.Add(new Attendee
                    {
                        Type = AttendeeType.Required,
                        EmailAddress = new EmailAddress
                        {
                            Address = person.EmailAddresses.First().Address
                        }
                    });
                }
                else if (entry is ITokenStringContainer)
                {
                    var container = entry as ITokenStringContainer;

                    // Treat any unrecognized text as a list of email addresses
                    var emails = container.Text
                        .Split(new[] { ';', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var email in emails)
                    {
                        try
                        {
                            // Validate the email address
                            var addr = new System.Net.Mail.MailAddress(email);
                            if (addr.Address == email)
                            {
                                attendeeList.Add(new Attendee
                                {
                                    Type = AttendeeType.Required,
                                    EmailAddress = new EmailAddress
                                    {
                                        Address = email
                                    }
                                });
                            }
                        }
                        catch { /* Invalid, skip */ }
                    }
                }
            }

            if (attendeeList.Count > 0)
            {
                newEvent.Attendees = attendeeList;
            }

            try
            {
                await graphClient.Me.Events.Request().AddAsync(newEvent);
            }
            catch (ServiceException graphException)
            {
                ShowNotification($"Exception creating new event: ${graphException.Message}");
            }

            CreateProgress.IsActive = false;
            CreateProgress.Visibility = Visibility.Collapsed;

            ShowNotification("Event created");
        }
        // </CreateEventSnippet>
    }
}
