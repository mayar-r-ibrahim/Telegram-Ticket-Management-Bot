# Telegram-Ticket-Management-Bot
A Telegram bot built with Google Apps Script for managing travel tickets, user permissions, and analytics. The bot integrates with Google Sheets as a database and provides a complete ticket management system with real-time notifications and reporting features.


Features
🎫 Ticket Management
Add Tickets: Create new tickets via form submission

View Tickets: Browse open tickets by month with detailed information

Search System: Advanced search across traveler names, ticket IDs, purchase/sale information

Status Management: Open/close tickets and toggle status

Edit Tickets: Direct links to edit ticket information

👥 User Management
Multi-level Authorization: Admin and regular user permissions

User Management: Add, edit, and remove users with different permission levels

Broadcast Lists: Separate lists for notifications and announcements

📊 Analytics & Reporting
Dashboard: Comprehensive analytics with multiple time periods

Employee Performance: Track sales and operations team performance

Export Capabilities: Generate Excel reports with filtered data

Daily Reports: Automated ticket notifications

🔍 Advanced Search
Multi-field search (traveler name, ticket ID, purchase/sale information)

Real-time search results with inline buttons

Quick status toggling from search results

Setup Instructions
Prerequisites
Google Account with access to Google Sheets and Apps Script

Telegram account and BotFather access to create a bot

Step 1: Google Sheets Setup
Create a Google Spreadsheet with the following sheets:

Tickets: Main ticket database with required columns:

Timestamp, Email, Traveler Name, Departure, Arrival, Ticket Type

Departure Date, Return Date, Ticket ID, Employee Names

Purchase/Sale information, Passport, Edit URL, Status

Users1: User management with columns:

User ID, Permission Level, Name

Users2: Broadcast list for notifications

Suggestions: Optional suggestions database

MessageIDs: For tracking message history (auto-created)

Step 2: Bot Configuration
Create a bot via BotFather on Telegram

Copy the bot token and replace in the code:

javascript
var token = "YOUR_BOT_TOKEN_HERE";
Step 3: Apps Script Deployment
Create a new Google Apps Script project

Paste the entire code into the script editor

Set up triggers:

doPost for webhooks (needs to be deployed as web app)

onFormSubmit for form submissions

sendDailyTicketReport for daily notifications (time-based trigger)

Step 4: Webhook Setup
Deploy the script as a web app with execute permissions for "Anyone"

Set the webhook URL in Telegram:

text
https://api.telegram.org/bot<YOUR_BOT_TOKEN>/setWebhook?url=<YOUR_WEB_APP_URL>
Commands
👤 User Commands
/start - Welcome message and bot overview

/help - Detailed help with video guides

/add - Create new ticket (redirects to form)

/tickets - View open tickets by month

/search - Search tickets across multiple fields

🔐 Admin Commands
/analytics - Access analytics dashboard

/users - User management system

/suggestions - Manage suggestions (if implemented)

Key Systems
Authorization System
Two-tier permission system: Admin and Regular users

Session management for multi-step interactions

Automatic session cleanup

Search System
Intelligent text processing with Arabic/English support

Multi-field fallback search

Inline results with quick actions

Analytics System
Time period analysis (day, week, month, quarter, year)

Employee performance tracking

Financial reporting (sales, purchases, profit)

Export to Excel functionality

Notification System
Real-time ticket updates in groups

Edit tracking with change detection

Daily summary reports

Broadcast messaging to user lists

Database Structure
Tickets Sheet Columns
The system automatically detects columns based on headers. Expected columns include:

Basic Info: Timestamp, Email, Traveler Name

Travel Details: Departure, Arrival, Dates, Ticket Type

Business Info: Purchase/Sale details, Employees

Management: Ticket ID, Status, Edit Link

User Management
Users1: Active bot users with permissions

Users2: Notification/broadcast list

Permission levels: "مشرف" (Admin) or "مستخدم عادي" (Regular User)

Customization
Adding New Commands
Register in commandRegistry

Implement handler function

Add to help messages if needed

Modifying Search Fields
Update the TICKET_COLUMNS mapping and search logic in processSearchTermByField

Adding Analytics
Extend the analyzeTicketsForPeriod function and add new callback handlers

Security Features
User authorization checks

Admin-only functionality protection

Session-based interaction management

Input validation and error handling

Troubleshooting
Common Issues
Webhook not working: Check deployment permissions and webhook URL

Sheet access errors: Verify sheet names and sharing permissions

Search not finding results: Check column headers match expected patterns

Permission errors: Ensure user is added to Users1 sheet with correct permissions

Logging
Extensive logging using Logger.log

Error handling with user-friendly messages

Debug mode available in analytics functions

