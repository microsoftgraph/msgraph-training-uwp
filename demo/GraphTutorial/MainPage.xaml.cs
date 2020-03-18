﻿using Microsoft.Toolkit.Graph.Providers;
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

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace GraphTutorial
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();

            // Load OAuth settings
            var oauthSettings = Windows.ApplicationModel.Resources.ResourceLoader.GetForCurrentView("OAuth");
            var appId = oauthSettings.GetString("AppId");
            var scopes = oauthSettings.GetString("Scopes");

            if (string.IsNullOrEmpty(appId) || string.IsNullOrEmpty(scopes))
            {
                Notification.Show("Could not load OAuth Settings from resource file.");
            }
            else
            {
                // Configure MSAL provider  
                MsalProvider.ClientId = appId;
                MsalProvider.Scopes = new ScopeSet(scopes.Split(' '));

                // Check signed-in state
                //var globalProvider = ProviderManager.Instance.GlobalProvider;
                //SetAuthState(globalProvider != null && globalProvider.State == ProviderState.SignedIn);

                // Handle auth state change
                ProviderManager.Instance.ProviderUpdated += ProviderUpdated;
                //globalProvider.StateChanged += AuthStateChanged;

                // Navigate to HomePage.xaml
                RootFrame.Navigate(typeof(HomePage));
            }
        }

        private void ProviderUpdated(object sender, ProviderUpdatedEventArgs e)
        {
            var globalProvider = ProviderManager.Instance.GlobalProvider;
            SetAuthState(globalProvider != null && globalProvider.State == ProviderState.SignedIn);
            RootFrame.Navigate(typeof(HomePage));
        }

        private void SetAuthState(bool isAuthenticated)
        {
            (Application.Current as App).IsAuthenticated = isAuthenticated;

            // Toggle controls that require auth
            Calendar.IsEnabled = isAuthenticated;
        }

        private void NavView_ItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
        {
            var invokedItem = args.InvokedItem as string;

            switch (invokedItem.ToLower())
            {
                case "calendar":
                    throw new NotImplementedException();
                    break;
                case "home":
                default:
                    RootFrame.Navigate(typeof(HomePage));
                    break;
            }
        }
    }
}