# Microsoft OAuth Setup Guide for Arcadeus

## Prerequisites
- Microsoft Azure account (free tier is fine)
- Access to Azure Portal (https://portal.azure.com)

## Step 1: Create Microsoft App Registration

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **"New registration"**
4. Fill in the details:
   - **Name**: Arcadeus (or your preferred name)
   - **Supported account types**: "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts"
   - **Redirect URI**: 
     - Platform: Web
     - URI: `https://arcadeus-5641b.firebaseapp.com/__/auth/handler`
5. Click **"Register"**

## Step 2: Get Application Credentials

### Application ID:
1. After registration, you'll see the app overview page
2. Copy the **"Application (client) ID"**
3. This is what you'll paste in the Firebase "Application ID" field

### Application Secret:
1. In your app, go to **"Certificates & secrets"** (left sidebar)
2. Click **"New client secret"**
3. Add a description (e.g., "Firebase Auth")
4. Choose expiration (recommended: 24 months)
5. Click **"Add"**
6. **IMPORTANT**: Copy the secret value immediately (it won't be shown again!)
7. This is what you'll paste in the Firebase "Application secret" field

## Step 3: Configure Authentication Settings

1. Go to **"Authentication"** (left sidebar)
2. Under "Platform configurations", you should see your Web platform
3. Make sure the redirect URI is listed: `https://arcadeus-5641b.firebaseapp.com/__/auth/handler`
4. Under "Implicit grant and hybrid flows", check:
   - ✓ Access tokens
   - ✓ ID tokens
5. Under "Supported account types", ensure it's set to multitenant + personal accounts
6. Click **"Save"**

## Step 4: API Permissions (Optional but recommended)

1. Go to **"API permissions"** (left sidebar)
2. You should already have `User.Read` permission
3. This is sufficient for basic authentication

## Step 5: Complete Firebase Setup

1. Go back to your Firebase Console
2. Paste the **Application ID** in the first field
3. Paste the **Application secret** in the second field
4. Click **"Save"**

## Troubleshooting

### Common Issues:

1. **"Redirect URI mismatch"**: Make sure the URI in Azure exactly matches Firebase's
2. **"Invalid client secret"**: Secret may have expired or was copied incorrectly
3. **"Unauthorized client"**: Check that multitenant + personal accounts is enabled

### Testing:

After setup, test by:
1. Going to your login page
2. Clicking "Continue with Microsoft"
3. You should see Microsoft's login page
4. After login, you should be redirected back to your app

## Security Notes

- Never commit the client secret to version control
- Rotate secrets periodically (before expiration)
- Use environment variables in production
- Consider implementing refresh token rotation

## Additional Resources

- [Microsoft identity platform documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/)
- [Firebase Auth with Microsoft](https://firebase.google.com/docs/auth/web/microsoft-oauth)