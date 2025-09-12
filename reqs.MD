# SSO Implementation Requirements

## Azure App Registration Changes
1. Update Authentication settings:
   - Configure Platform settings for Single-page application (SPA)
   - Add redirect URI: https://outofoffice.oztiryakiler.com.tr/taskpane.html
   - Remove old static app URL
   - Configure Access tokens and ID tokens in implicit grant
   - Add redirect URI for Office client: https://www.office.com/sigmain

2. Update API permissions:
   - Microsoft Graph permissions needed:
     - ~~User.Read~~ (already present)
     - ~~MailboxSettings.ReadWrite~~ (already present)
     - ~~People.Read~~ (already present)
     - ~~openid~~ (already present)
     - ~~profile~~ (already present)

3. Update App Roles and Exposures:
   - Add scope "access_as_user" with following properties:
     - Who can consent: Admins and users
     - Admin consent display name: "Access Add-in"
     - Admin consent description: "Allows the add-in to access the user's mailbox settings."
     - User consent display name: "Access your mailbox settings"
     - User consent description: "Allows the add-in to access your mailbox settings."
     - State: Enabled
   - Add client application: d3590ed6-52b3-4102-aeff-aad2292ab01c (Microsoft Office)

## Office Add-in Manifest Changes
1. Update VersionOverrides:
   - Remove nested v1.1 VersionOverrides
   - Move WebApplicationInfo to main v1.1 VersionOverrides section
   - Add required SSO scopes

2. Update Permissions:
   - Verify ReadWriteMailbox permission
   - Add any additional required permissions

## Frontend Application Changes
1. Authentication Implementation:
   - Add Microsoft Authentication Library (@azure/msal-browser)
   - Implement getGraphToken() using Office.auth.getAccessToken()
   - Implement fallback authentication using MSAL for non-SSO scenarios
   - Handle consent prompts and token acquisition

2. API Integration:
   - Update API calls to use acquired tokens
   - Implement error handling for auth failures
   - Add token refresh logic

## Testing Requirements
1. Test SSO flow in different scenarios:
   - Test with Office on the web
   - Test with Office desktop client
   - Test SSO and fallback authentication
   - Test consent scenarios
   - Test error handling:
     - 13001 - No support for SSO
     - 13002 - Authentication in progress
     - 13003 - User not signed in
     - 13004 - Invalid resource

2. Verify permissions:
   - Test access to required Graph API endpoints
   - Verify mailbox settings access
   - Test user profile access

## Documentation Updates
1. Update deployment documentation
2. Add SSO configuration guide
3. Update troubleshooting guide

Please review these requirements before proceeding with the implementation.