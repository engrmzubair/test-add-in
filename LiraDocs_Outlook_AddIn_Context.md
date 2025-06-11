# LiraDocs Outlook Add-in Development Project Context

## Project Background & Purpose

I am building an Outlook add-in for **LiraDocs** - a legal case management system designed specifically for lawyers and law firms. The add-in's primary purpose is to enable lawyers to seamlessly transfer emails from Outlook to their LiraDocs case management system, similar to functionality I've already built for Gmail.

## Why Outlook Add-in (Not Browser Extension)?

- **Native Integration**: Outlook add-ins provide deeper, native integration with the email client
- **Cross-Platform Support**: Works across Outlook Web, Desktop (Windows/Mac), and Mobile
- **Enterprise Deployment**: Can be deployed through Microsoft 365 Admin Center for law firms
- **Professional Security**: Higher security standards required for legal/attorney-client privileged communications
- **API Access**: Full access to Office.js APIs for comprehensive email processing

## Existing Gmail Extension Context

I have already built a successful Chrome/Firefox extension for Gmail with these features:
- **Tech Stack**: Plasmo framework, React, TypeScript, Content Scripts
- **Functionality**: Email transfer to LiraDocs server, visual transfer status tracking
- **Storage**: Cloudflare KV for tracking transferred email status
- **APIs**: Gmail API for email content access
- **UI**: React components for transfer modal and status indicators
- **Authentication**: Login/logout flow with LiraDocs server
- **Badge System**: Visual badges showing transfer status with unlink capability

## Existing Cloudflare Infrastructure

My Chrome extension already uses a robust Cloudflare Worker + KV setup:
- **Cloudflare Worker**: Handles API requests and KV operations
- **KV Storage**: Tracks transferred email IDs per client/user
- **Worker Endpoints**:
  - `/api/store-email-mapping` - Store transferred email IDs
  - `/api/get-email-ids` - Get list of transferred emails for badge checking
  - `/api/remove-email` - Remove email mapping for unlink functionality
- **Data Structure**: Email IDs stored against client ID for status tracking

## Authentication & User Flow

### Login Process
- **User Authentication**: Users must login to LiraDocs server before transferring emails
- **Session Management**: Maintain authenticated session for API calls
- **Case Data Access**: Retrieve user's case matters/list from LiraDocs server after authentication
- **Logout Functionality**: Proper session cleanup and logout capability

### Email Transfer Flow
1. User logs into LiraDocs through add-in interface
2. Add-in fetches user's case matters from LiraDocs server
3. User selects email to transfer
4. User chooses target case from their case list
5. Email converted to .eml format and sent to LiraDocs server
6. Transfer status stored in Cloudflare KV and badge updated

## Badge System & Visual Indicators

### Transfer Status Badges
- **"LiraDocs" Badge**: Displayed on emails that have been successfully transferred
- **Clear/Unlink Icon**: Button to remove email from LiraDocs and clear the badge
- **Visual Feedback**: Immediate indication of transfer status in email list

### Status Tracking Implementation
- **Storage Method**: Reuse existing Cloudflare KV storage system
- **Data Structure**: Store email IDs against client ID (same as Chrome extension)
- **Badge Logic**: 
  - HTTP request to existing Cloudflare Worker to get transferred emails for current user
  - Check if current email ID exists in transferred list from KV
  - Show "LiraDocs" badge if transferred, no badge if not transferred
- **Unlink Functionality**: HTTP request to remove email ID from KV storage and clear badge when user clicks unlink

## Target Outlook Add-in Requirements

### Core Functionality
- **User Authentication**: Login/logout with LiraDocs server
- **Case Management**: Fetch and display user's case matters from LiraDocs
- **Email Transfer**: Convert emails to .eml format and transfer to LiraDocs server
- **Transfer Status Tracking**: Visual badges showing which emails have been transferred
- **Unlink Capability**: Remove emails from LiraDocs and clear badges
- **Duplicate Prevention**: Prevent accidental re-transfer of same emails
- **Professional UI**: Clean, lawyer-friendly interface matching legal workflows

### Technical Requirements
- **Framework**: Yeoman Generator for Office Add-ins (yo office) with React + TypeScript
- **APIs**: Office.js Mailbox APIs, EWS for detailed email access
- **Email Processing**: Convert email content to .eml format for server transfer
- **HTTP Requests**: 
  - POST login credentials to LiraDocs authentication endpoint
  - GET user's case matters from LiraDocs API
  - POST .eml email files to LiraDocs server
  - HTTP requests to existing Cloudflare Worker for KV operations
- **Status Storage**: Reuse existing Cloudflare KV infrastructure
- **Authentication**: OAuth/session management with LiraDocs system
- **Permissions**: ReadWriteMailbox level (required for full email access)

### User Experience Goals
- **Seamless Login**: Simple authentication flow similar to Gmail extension
- **Case Selection**: Easy dropdown/selection of user's cases
- **One-Click Transfer**: Simple transfer process after authentication
- **Clear Status Visibility**: Obvious visual feedback on transfer status with badges
- **Unlink Capability**: Easy way to remove emails from LiraDocs
- **Error Handling**: Graceful handling of authentication and transfer errors
- **Mobile Friendly**: Works on Outlook mobile apps

## Data Flow & Architecture

### Complete User Journey
1. **Authentication**: User logs into LiraDocs through add-in
2. **Case Loading**: Add-in fetches user's case matters from LiraDocs API
3. **Email Selection**: User selects email in Outlook
4. **Badge Check**: Add-in queries Cloudflare Worker to check if email is already transferred
5. **Transfer Interface**: If not transferred, show transfer options with case selection
6. **Email Processing**: Convert email to .eml format with all content/attachments
7. **Server Transfer**: POST .eml file to LiraDocs server with case association
8. **Status Update**: Store email ID in Cloudflare KV via Worker API for current user
9. **Badge Display**: Show "LiraDocs" badge with unlink option
10. **Unlink Option**: Allow user to remove email from LiraDocs and clear badge via Worker API

### Storage Architecture (Reuse Existing Cloudflare KV)
```
Outlook Add-in → HTTP Requests → Cloudflare Worker → Cloudflare KV
                                     ↓
Chrome Extension → HTTP Requests → (Same Worker) → (Same KV)
```

- **User Sessions**: Store authentication tokens/session data locally in add-in
- **Transferred Emails**: Use existing Cloudflare KV storage via Worker API
- **Case Data**: Cache user's case matters locally for quick access
- **Sync Logic**: Same KV storage serves both Chrome and Outlook users consistently

### API Integration Points
- **LiraDocs Authentication API**: Login/logout endpoints
- **LiraDocs Cases API**: Fetch user's case matters
- **LiraDocs Email Transfer API**: POST .eml files with case association
- **Cloudflare Worker API**: Existing endpoints for KV operations
  - `/api/store-email-mapping` - Store transferred email status
  - `/api/get-email-ids` - Check transfer status for badge display
  - `/api/remove-email` - Remove transfer status for unlink

## Architecture Decisions

### Development Framework
- **Chosen**: Yeoman Generator for Office Add-ins (yo office)
- **Reason**: Purpose-built for traditional Office add-ins, excellent for email processing, proven reliability, cross-platform compatibility
- **Why NOT Microsoft 365 Agents Toolkit**: Designed for AI-powered Copilot agents, not traditional email processing add-ins

### Storage Strategy (Cloudflare KV - Reuse Existing)
- **Solution**: Continue using existing Cloudflare Worker + KV infrastructure
- **Implementation**: Make HTTP requests from Outlook add-in to same Worker endpoints
- **Advantages**: 
  - Reuse proven, working infrastructure
  - Maintain data consistency between Chrome and Outlook users
  - Global distribution via Cloudflare edge network
  - No additional storage setup required

### Email Processing Strategy
- **Mailbox APIs**: For standard email operations and content extraction
- **EWS Integration**: For advanced email content and header access when needed
- **.eml Conversion**: Convert Outlook email format to standard .eml for server transfer
- **Async Processing**: Handle large emails and attachments efficiently


## Development Priorities

1. **Authentication Security**: Secure login/logout flow with LiraDocs server
2. **Badge System Reliability**: Accurate transfer status tracking and display using existing KV
3. **Email Conversion**: Proper .eml format generation with all content preserved
4. **User Experience**: Intuitive workflow matching Gmail extension experience
5. **Performance**: Handle large emails and attachments without issues
6. **Compliance**: Meet legal industry standards and requirements

## Questions to Help With

When assisting me, please consider this context for questions about:
- Office.js API usage for email content extraction and .eml conversion
- Authentication flow implementation in Outlook add-ins
- Badge system implementation and visual indicators in Outlook
- HTTP requests to existing Cloudflare Worker from Outlook add-ins
- Yeoman Generator setup and configuration for Outlook add-ins
- React component architecture for login/transfer interfaces
- Integration with existing Cloudflare KV infrastructure
- Email ID tracking and transfer status management
- Error handling for authentication and transfer failures
- Cross-platform compatibility (Web/Desktop/Mobile)
- Testing strategies for authentication and badge systems
- CORS configuration for Cloudflare Worker

## Current Development Phase

I am in the **planning and initial development** phase, having confirmed that Yeoman Generator for Office Add-ins is the correct tool for my traditional email processing needs. Key implementation focuses include:
- Replicating the authentication flow from Chrome extension
- Implementing badge system using existing Cloudflare KV infrastructure
- Converting Outlook emails to .eml format for server transfer
- Managing transfer status tracking via HTTP requests to existing Worker
- Ensuring CORS compatibility between Outlook add-in and Cloudflare Worker

## Framework Choice Analysis

### Microsoft 365 Agents Toolkit is NOT Suitable for LiraDocs
The Microsoft 365 Agents Toolkit is primarily designed for:
- **AI-powered Copilot agents** with natural language processing
- **Teams apps integration** and agent-based workflows  
- **Declarative agents** in Word, Excel, and PowerPoint
- **Add-in actions** that integrate with Copilot chat
- **Agent debugging experience** with AI capabilities

### Yeoman Generator (yo office) is PERFECT for LiraDocs
The Yeoman Generator for Office Add-ins is designed for:
- **Traditional email processing** and transfer functionality
- **Badge system** with visual indicators in Outlook
- **Direct HTTP requests** to existing Cloudflare infrastructure
- **Cross-platform compatibility** (Web, Desktop, Mobile) 
- **Enterprise deployment** capability
- **Standard Office.js APIs** without AI/agent requirements

## Chrome Extension Reference Points

My existing Chrome extension provides these features that need to be replicated:
- **Login/Logout Flow**: Seamless authentication with LiraDocs server
- **Cloudflare KV Integration**: Email ID tracking per client via Worker API
- **Badge System**: Visual "LiraDocs" badges with unlink capability
- **Case Integration**: Dropdown selection of user's cases
- **.eml Transfer**: Email conversion and server transfer functionality

## Technical Implementation Notes

### Cloudflare Worker Integration
- **Existing Endpoints**: Can be used as-is for Outlook add-in
- **CORS Headers**: May need to update Worker to allow Outlook add-in domains
- **Authentication**: Same user session management can be maintained
- **Data Consistency**: Both Chrome and Outlook users share same KV storage

### Office.js Considerations
- **Email Access**: Use `Office.context.mailbox.item` for email content
- **EWS Integration**: Use `makeEwsRequestAsync` for advanced email data when needed
- **HTTP Requests**: Use `fetch()` API to communicate with Cloudflare Worker
- **Error Handling**: Proper async/await patterns for reliable operation 