# SES Bounce Notification Extractor

A Go application that extracts and processes AWS SES bounce notifications from Microsoft 365/Outlook mailboxes using the Microsoft Graph API.

## Description

This tool connects to your Microsoft 365 mailbox, searches for AWS SES bounce notification emails in a specified folder, parses the bounce details from the JSON payload, and outputs the results in a tab-separated format that can be easily pasted into Excel for analysis.

The application specifically looks for emails from AWS SNS with the subject "VerifiedSESBounceNotification" and extracts:
- Bounced recipient email addresses
- Bounce timestamps
- Bounce types (Permanent, Transient, etc.)
- Original email subjects
- Diagnostic codes explaining why the bounce occurred

## Prerequisites

### Development Requirements

- **Go**: Version 1.24.0 or later
- **Azure AD App Registration**: Required for Microsoft Graph API authentication
- **Microsoft 365 Account**: With access to the mailbox containing bounce notifications

### Azure AD App Setup

Before running the application, you need to create an Azure AD App Registration:

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** > **App registrations**
3. Click **New registration**
4. Configure:
   - **Name**: `EmailBounceExtractor` (or your preferred name)
   - **Supported account types**: Single tenant
   - **Redirect URI**: Select "Public client/native (mobile & desktop)" and enter `http://localhost`
5. Click **Register**
6. In the app's overview page, note the **Application (client) ID** and **Directory (tenant) ID**
7. Go to **API permissions**:
   - Click **Add a permission**
   - Select **Microsoft Graph**
   - Select **Delegated permissions**
   - Add `Mail.Read` permission
   - (Optional) Click **Grant admin consent** to avoid prompting users

## Setup

1. **Clone the repository** (or download the source code):
   ```bash
   git clone <repository-url>
   cd ses-bounce-notification-extractor
   ```

2. **Install dependencies**:
   ```bash
   go mod download
   ```

3. **Configure credentials**:

   Create a `.env` file in the project root directory (copy from `.env.example`):
   ```bash
   cp .env.example .env
   ```

   Edit `.env` and add your Azure AD credentials:
   ```env
   AZURE_CLIENT_ID=your-client-id-here
   AZURE_TENANT_ID=your-tenant-id-here
   MAIL_FOLDER_NAME=SES Bounces
   ```

   **Alternative**: Set environment variables directly instead of using a `.env` file:
   ```bash
   export AZURE_CLIENT_ID="your-client-id-here"
   export AZURE_TENANT_ID="your-tenant-id-here"
   export MAIL_FOLDER_NAME="SES Bounces"
   ```

4. **Create a mail folder** (if it doesn't exist):
   - In your Outlook mailbox, create a folder with the name specified in `MAIL_FOLDER_NAME` (default: `SES Bounces`)
   - Set up a mail rule to automatically move AWS SES bounce notifications to this folder

## Usage

### Running the Application

1. **Compile and run**:
   ```bash
   go run main.go
   ```

   Or build an executable:
   ```bash
   go build -o ses-bounce-notification-extractor
   ./ses-bounce-notification-extractor
   ```

2. **First-time authentication**:
   - On the first run, a browser window will open automatically
   - Sign in with your Microsoft 365 credentials
   - Grant the requested permissions (Mail.Read)
   - The authentication token will be cached locally in `.auth_cache.json`

3. **View the results**:
   - The application outputs tab-separated data to stdout (console)
   - Each line contains: Email Address | Date | Subject | Diagnostic Code
   - The output is formatted for easy copying into Excel or Google Sheets

### Configuration Options

#### Environment Variables

The following environment variables can be configured in your `.env` file or set directly in your shell:

- **AZURE_CLIENT_ID** (required): Your Azure AD Application (client) ID
- **AZURE_TENANT_ID** (required): Your Azure AD Directory (tenant) ID
- **MAIL_FOLDER_NAME** (optional): Name of the mail folder containing bounce notifications
  - Default: `SES Bounces`
  - Change this to match your mailbox folder structure

#### Code Modifications

You can modify the following settings in `main.go`:

- **Message filter**: Modify the filter in `findAllBouncedMessages()` to change which messages are processed
  - Currently filters for: unread messages from `no-reply@sns.amazonaws.com` with subject `VerifiedSESBounceNotification`
  - Set `isRead eq false` to `isRead eq true` to process already-read messages

### Output Format

The application outputs tab-separated values with the following columns:
