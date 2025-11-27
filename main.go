package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	"github.com/joho/godotenv"
	msgraphsdk "github.com/microsoftgraph/msgraph-sdk-go"
	msgraphcore "github.com/microsoftgraph/msgraph-sdk-go-core"
	"github.com/microsoftgraph/msgraph-sdk-go/models"
	"github.com/microsoftgraph/msgraph-sdk-go/users"
)

const authCacheFile = ".auth_cache.json"

type Config struct {
	ClientID       string
	TenantID       string
	MailFolderName string
}

type Message struct {
	Subject string
	Body    string
}

type BounceNotification struct {
	NotificationType string `json:"notificationType"`
	Bounce           struct {
		BounceType        string `json:"bounceType"`
		BounceSubType     string `json:"bounceSubType"`
		BouncedRecipients []struct {
			EmailAddress   string `json:"emailAddress"`
			Action         string `json:"action"`
			Status         string `json:"status"`
			DiagnosticCode string `json:"diagnosticCode"`
		} `json:"bouncedRecipients"`
		Timestamp string `json:"timestamp"`
	} `json:"bounce"`
	Mail struct {
		Source        string   `json:"source"`
		Destination   []string `json:"destination"`
		CommonHeaders struct {
			Subject string `json:"subject"`
		} `json:"commonHeaders"`
	} `json:"mail"`
}

func main() {
	// Load configuration from the.env file if it exists, then from environment variables
	config, err := loadConfig()
	if err != nil {
		panic(err)
	}

	client, err := setupMSAPIClient(config)
	if err != nil {
		panic(err)
	}

	messages, err := findAllBouncedMessages(config.MailFolderName, client)
	if err != nil {
		panic(err)
	}

	bounces, err := parseBounceNotifications(messages)
	if err != nil {
		panic(err)
	}

	for _, bounce := range bounces {
		var recipients []string
		var diagnosticCodes []string
		for _, recipient := range bounce.Bounce.BouncedRecipients {
			recipients = append(recipients, recipient.EmailAddress)
			diagnosticCodes = append(diagnosticCodes, recipient.DiagnosticCode)
		}
		var timestamp time.Time
		if bounce.Bounce.Timestamp != "" {
			timestamp, _ = time.Parse(time.RFC3339, bounce.Bounce.Timestamp)
		}
		fmt.Printf(
			"%s\t%s\t\t%s\t%s\n",
			strings.Join(recipients, ", "),
			timestamp.Format("02-01-2006"),
			bounce.Mail.CommonHeaders.Subject,
			strings.Replace(strings.Join(diagnosticCodes, ", "), "\n", " ", -1),
		)
	}
}

func setupMSAPIClient(config Config) (*msgraphsdk.GraphServiceClient, error) {
	// Try to load a cached authentication record
	authRecord, err := loadAuthRecord(authCacheFile)

	credOptions := &azidentity.InteractiveBrowserCredentialOptions{
		ClientID: config.ClientID,
		TenantID: config.TenantID,
	}

	// If we have a cached record, use it for silent authentication
	if err == nil && authRecord != nil {
		credOptions.AuthenticationRecord = *authRecord
		credOptions.LoginHint = authRecord.Username
	}

	cred, err := azidentity.NewInteractiveBrowserCredential(credOptions)
	if err != nil {
		return nil, err
	}

	// Use Mail.Read scope for a Graph client
	client, err := msgraphsdk.NewGraphServiceClientWithCredentials(cred, []string{"Mail.Read"})
	if err != nil {
		return nil, err
	}

	// After a successful first use, save the auth record for next time
	// Only save if we didn't load from cache (first run)
	if authRecord == nil {
		// Authenticate with the required scope to get the record
		ctx := context.Background()
		record, err := cred.Authenticate(ctx, &policy.TokenRequestOptions{
			Scopes: []string{"https://graph.microsoft.com/.default"},
		})
		if err == nil {
			_ = saveAuthRecord(authCacheFile, record)
		}
	}

	return client, nil
}

func loadConfig() (Config, error) {
	// Try to load the .env file (it's okay if it doesn't exist)
	_ = godotenv.Load()

	clientID := os.Getenv("AZURE_CLIENT_ID")
	if clientID == "" {
		return Config{}, fmt.Errorf("AZURE_CLIENT_ID environment variable is required")
	}

	tenantID := os.Getenv("AZURE_TENANT_ID")
	if tenantID == "" {
		return Config{}, fmt.Errorf("AZURE_TENANT_ID environment variable is required")
	}

	mailFolderName := os.Getenv("MAIL_FOLDER_NAME")
	if mailFolderName == "" {
		mailFolderName = "SES Bounces" // Default value
	}

	return Config{
		ClientID:       clientID,
		TenantID:       tenantID,
		MailFolderName: mailFolderName,
	}, nil
}

func loadAuthRecord(filePath string) (*azidentity.AuthenticationRecord, error) {
	data, err := os.ReadFile(filePath)
	if err != nil {
		if os.IsNotExist(err) {
			fmt.Println("No auth cache file found")
		}
		return nil, err
	}

	var record azidentity.AuthenticationRecord
	err = json.Unmarshal(data, &record)
	if err != nil {
		fmt.Printf("Failed to parse auth cache: %v\n", err)
		return nil, err
	}

	return &record, nil
}

func saveAuthRecord(filePath string, record azidentity.AuthenticationRecord) error {
	data, err := json.Marshal(record)
	if err != nil {
		return err
	}

	// Create a directory if it doesn't exist
	dir := filepath.Dir(filePath)
	if dir != "." && dir != "" {
		err = os.MkdirAll(dir, 0700)
		if err != nil {
			return err
		}
	}

	return os.WriteFile(filePath, data, 0600)
}

func findAllBouncedMessages(folderDisplayName string, client *msgraphsdk.GraphServiceClient) ([]Message, error) {
	var bounceMessages []Message

	// List mail folders to find the folder ID by display name
	folders, err := client.Me().MailFolders().Get(context.Background(), nil)
	if err != nil {
		return bounceMessages, fmt.Errorf("failed to get mail folders: %w", err)
	}

	var folderId string
	for _, f := range folders.GetValue() {
		if f.GetDisplayName() != nil && *f.GetDisplayName() == folderDisplayName {
			folderId = *f.GetId()
			break
		}
	}
	if folderId == "" {
		return bounceMessages, fmt.Errorf("folder '%s' not found", folderDisplayName)
	}

	// Query messages in the folder with the subject filter
	filter := "from/emailAddress/address eq 'no-reply@sns.amazonaws.com' and from/emailAddress/name eq 'VerifiedSESBounceNotification' and isRead eq false"
	messages, err := client.Me().MailFolders().ByMailFolderId(folderId).Messages().Get(context.Background(), &users.ItemMailFoldersItemMessagesRequestBuilderGetRequestConfiguration{
		QueryParameters: &users.ItemMailFoldersItemMessagesRequestBuilderGetQueryParameters{
			Filter: &filter,
		},
	})
	if err != nil {
		return bounceMessages, fmt.Errorf("failed to get messages: %w", err)
	}

	// Use PageIterator to handle pagination automatically
	iterator, err := msgraphcore.NewPageIterator[models.Messageable](messages, client.GetAdapter(), models.CreateMessageCollectionResponseFromDiscriminatorValue)
	if err != nil {
		return bounceMessages, fmt.Errorf("failed to create iterator: %w", err)
	}

	err = iterator.Iterate(context.Background(), func(msg models.Messageable) bool {
		// Extract other data as needed
		bounceMessages = append(bounceMessages, Message{Subject: *msg.GetSubject(), Body: *msg.GetBody().GetContent()})

		return true // Return true to continue to the next message/page
	})
	if err != nil {
		return bounceMessages, fmt.Errorf("failed to iterate messages: %w", err)
	}

	return bounceMessages, nil
}

func parseBounceNotifications(messages []Message) ([]BounceNotification, error) {
	var bounceNotifications []BounceNotification
	for _, message := range messages {
		startIndex := strings.Index(message.Body, "{")
		if startIndex != -1 {
			// Use a decoder to parse just the first JSON object from the stream
			// This avoids issues with trailing text (like the email footer)
			jsonPart := message.Body[startIndex:]
			decoder := json.NewDecoder(strings.NewReader(jsonPart))

			var bounceNotification BounceNotification
			if err := decoder.Decode(&bounceNotification); err != nil {
				return bounceNotifications, fmt.Errorf("error unmarshaling JSON: %v", err)
			} else {
				// Check if the bounce struct is actually populated
				if bounceNotification.Bounce.BounceType != "" {
					bounceNotifications = append(bounceNotifications, bounceNotification)
				}
			}
		}
	}

	return bounceNotifications, nil
}
