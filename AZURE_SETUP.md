# Azure App Registration Setup for SSO

This document explains how to set up Azure App Registration to enable Single Sign-On (SSO) for the Outlook Auto Reply Add-in.

## Prerequisites

- Azure Active Directory (Azure AD) tenant
- Global Administrator or Application Administrator permissions in Azure AD
- The add-in manifest.xml file

## Step 1: Create Azure App Registration

1. Sign in to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in the registration details:
   - **Name**: `Outlook Auto Reply Add-in`
   - **Supported account types**: Select "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank for now
5. Click **Register**

## Step 2: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions**
5. Add the following permissions:
   - `User.Read` - Read user profile
   - `MailboxSettings.ReadWrite` - Read and write mailbox settings
   - `People.Read` - Read users' relevant people lists
6. Click **Add permissions**
7. Click **Grant admin consent** (requires admin privileges)

## Step 3: Configure Authentication

1. Go to **Authentication** in your app registration
2. Click **Add a platform**
3. Select **Single-page application (SPA)**
4. Add the following redirect URIs:
   - `https://ozridvan.github.io/outlook-auto-reply-addin-v2/taskpane.html`
   - `https://localhost:3000/taskpane.html` (for development)
5. Under **Implicit grant and hybrid flows**, check:
   - ✅ Access tokens
   - ✅ ID tokens
6. Click **Configure**

## Step 4: Expose an API

1. Go to **Expose an API**
2. Click **Set** next to Application ID URI
3. Accept the default URI: `api://{your-app-id}`
4. Click **Save**
5. Click **Add a scope**
6. Fill in the scope details:
   - **Scope name**: `access_as_user`
   - **Who can consent**: Admins and users
   - **Admin consent display name**: `Access the app as the user`
   - **Admin consent description**: `Allows Office to call the app's web APIs as the current user`
   - **User consent display name**: `Access the app as you`
   - **User consent description**: `Allows Office to call the app's web APIs as you`
   - **State**: Enabled
7. Click **Add scope**

## Step 5: Update Manifest Configuration

1. Copy your **Application (client) ID** from the app registration overview
2. Update the following files with your Application ID:

### manifest.xml
Replace `YOUR_AZURE_APP_ID_HERE` with your actual Application ID:

```xml
<WebApplicationInfo>
  <Id>YOUR_ACTUAL_APP_ID_HERE</Id>
  <Resource>api://YOUR_ACTUAL_APP_ID_HERE</Resource>
  <Scopes>
    <Scope>https://graph.microsoft.com/User.Read</Scope>
    <Scope>https://graph.microsoft.com/MailboxSettings.ReadWrite</Scope>
    <Scope>https://graph.microsoft.com/People.Read</Scope>
  </Scopes>
</WebApplicationInfo>
```

### taskpane.js
Update the AUTH_CONFIG object:

```javascript
const AUTH_CONFIG = {
    clientId: 'YOUR_ACTUAL_APP_ID_HERE',
    scopes: [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/MailboxSettings.ReadWrite',
        'https://graph.microsoft.com/People.Read'
    ]
};
```

## Step 6: Test the Configuration

1. Deploy your updated add-in
2. Install it in Outlook
3. The add-in should now prompt for consent when accessing Microsoft Graph APIs
4. Check the browser console for authentication success/failure messages

## Troubleshooting

### Common Error Codes

- **13000**: SSO is not supported on this platform
- **13001**: User is not signed in to Office
- **13002**: User consent is required
- **13003**: User consent was not granted
- **13006**: User is not in a supported Microsoft 365 subscription
- **13012**: Add-in is not configured for SSO

### Solutions

1. **Error 13000**: Use Outlook on the web or desktop version
2. **Error 13001**: Ensure user is signed in to Office with organizational account
3. **Error 13002/13003**: Grant admin consent in Azure portal
4. **Error 13012**: Verify manifest.xml WebApplicationInfo section is correct

## Security Considerations

- Never hardcode client secrets in client-side code
- Use the principle of least privilege for API permissions
- Regularly review and audit permissions
- Monitor authentication logs in Azure AD

## Additional Resources

- [Microsoft Graph permissions reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- [Office Add-ins SSO documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins)
- [Azure AD app registration guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
