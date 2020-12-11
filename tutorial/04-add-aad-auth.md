<!-- markdownlint-disable MD002 MD041 -->

In this exercise you will extend the application from the previous exercise to support authentication with Azure AD. This is required to obtain the necessary OAuth access token to call the Microsoft Graph. In this step you will integrate the **LoginButton** control from the [Windows Graph Controls](https://github.com/windows-toolkit/Graph-Controls) into the application.

1. Right-click the **GraphTutorial** project in Solution Explorer and select **Add > New Item...**. Choose **Resources File (.resw)**, name the file `OAuth.resw` and select **Add**. When the new file opens in Visual Studio, create two resources as follows.

    - **Name:** `AppId`, **Value:** the app ID you generated in Application Registration Portal
    - **Name:** `Scopes`, **Value:** `User.Read User.ReadBasic.All People.Read MailboxSettings.Read Calendars.ReadWrite`

    ![A screenshot of the OAuth.resw file in the Visual Studio editor](./images/edit-resources-01.png)

    > [!IMPORTANT]
    > If you're using source control such as git, now would be a good time to exclude the `OAuth.resw` file from source control to avoid inadvertently leaking your app ID.

## Configure the LoginButton control

1. Open `MainPage.xaml.cs` and add the following `using` statement to the top of the file.

    ```csharp
    using Microsoft.Toolkit.Graph.Providers;
    ```

1. Replace the existing constructor with the following.

    :::code language="csharp" source="../demo/GraphTutorial/MainPage.xaml.cs" id="ConstructorSnippet":::

    This code loads the settings from `OAuth.resw` and initializes the MSAL provider with those values.

1. Now add an event handler for the `ProviderUpdated` event on the `ProviderManager`. Add the following function to the `MainPage` class.

    :::code language="csharp" source="../demo/GraphTutorial/MainPage.xaml.cs" id="ProviderUpdatedSnippet":::

    This event triggers when the provider changes, or when the provider state changes.

1. In Solution Explorer, expand **HomePage.xaml** and open `HomePage.xaml.cs`. Replace the existing constructor with the following.

    :::code language="csharp" source="../demo/GraphTutorial/HomePage.xaml.cs" id="ConstructorSnippet":::

1. Restart the app and click the **Sign In** control at the top of the app. Once you've signed in, the UI should change to indicate that you've successfully signed-in.

    ![A screenshot of the app after signing in](./images/add-aad-auth-01.png)

    > [!NOTE]
    > The `ButtonLogin` control implements the logic of storing and refreshing the access token for you. The tokens are stored in secure storage and refreshed as needed.
