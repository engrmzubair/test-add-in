# LiraDocs Outlook Add-in

A professional Outlook add-in for transferring emails to the LiraDocs legal case management system with visual status tracking.

## 🚀 **2025 Setup - Latest Yeoman Generator**

This project was created using the **latest Yeoman Generator for Office Add-ins (v3.0.1)** with:
- **React + TypeScript** framework
- **Fluent UI** components for native Office look
- **Office.js Mailbox APIs** for email processing
- **Cloudflare Worker integration** for status tracking

## 📋 **Features**

- ✅ **Email Transfer**: Convert emails to .eml format and transfer to LiraDocs
- ✅ **Status Tracking**: Visual badges showing transfer status using Cloudflare KV
- ✅ **Badge System**: "LiraDocs" badges with unlink capability
- ✅ **Cross-Platform**: Works on Outlook Web, Desktop (Windows/Mac), and Mobile
- ✅ **Modern UI**: Fluent UI components for consistent Office experience

## 🛠 **Development Setup**

### Prerequisites
- Node.js 16+ (tested with v20.18.2)
- npm 7+ (tested with v10.8.2)
- Office 365 subscription or Outlook Web access

### Installation
```bash
# Clone and navigate to project
cd liradocs-outlook-addin

# Install dependencies
npm install

# Start development server
npm start
```

### Testing the Add-in

1. **Outlook Web**: The add-in will automatically open in Outlook Web
2. **Desktop Outlook**: Use sideloading via the manifest.xml file
3. **Development**: Access at https://localhost:3000

## 🔧 **Project Structure**

```
liradocs-outlook-addin/
├── src/
│   ├── services/
│   │   ├── cloudflareService.ts    # Cloudflare Worker API integration
│   │   └── officeService.ts        # Office.js email operations
│   ├── taskpane/
│   │   ├── components/
│   │   │   └── App.tsx            # Main LiraDocs UI component
│   │   ├── outlook.ts             # Outlook-specific functions
│   │   └── taskpane.html          # Task pane HTML
│   └── commands/
├── manifest.xml                   # Add-in manifest (LiraDocs branded)
└── package.json
```

## 🌐 **Cloudflare Worker Integration**

The add-in integrates with your existing Cloudflare Worker:
- **Endpoint**: `https://liradocs-email-transfer.imran-71e.workers.dev`
- **Operations**: 
  - `GET /emails/{clientId}` - Get transferred email IDs
  - `POST /emails` - Store email ID
  - `DELETE /emails/{clientId}/{emailId}` - Remove email ID

## 📧 **Email Processing**

### Supported Operations
- **Email Content Extraction**: Subject, body, from, to, cc, bcc, attachments
- **.eml Conversion**: Standard email format for server transfer
- **Unique ID Generation**: Using internetMessageId or itemId
- **Transfer Status**: Real-time checking via Cloudflare KV

### Office.js APIs Used
- `Office.context.mailbox.item` - Current email access
- `item.body.getAsync()` - Email body extraction
- `item.internetMessageId` - Unique email identification

## 🎨 **UI Components**

Built with **Fluent UI** for native Office experience:
- **Cards**: Email information display
- **Badges**: Transfer status indicators
- **Buttons**: Transfer/remove actions
- **MessageBar**: Status messages
- **Spinner**: Loading states

## 🚀 **Deployment Options**

### 1. Microsoft 365 Admin Center (Recommended)
- Upload manifest.xml
- Deploy to organization users
- Centralized management

### 2. Sideloading (Development)
- Load manifest.xml locally
- Test in Outlook Web/Desktop
- Development and testing

### 3. AppSource (Public Distribution)
- Submit for public availability
- Microsoft validation required
- Global distribution

## 🔐 **Security & Authentication**

- **HTTPS Required**: All communications encrypted
- **CORS Configured**: Proper headers for Cloudflare Worker
- **Domain Whitelist**: LiraDocs and Cloudflare domains approved
- **Future**: NAA (Nested App Authentication) ready for 2025

## 📝 **Development Commands**

```bash
# Development
npm start              # Start dev server with hot reload
npm run build:dev      # Build for development
npm run build          # Build for production

# Testing & Validation
npm run validate       # Validate manifest.xml
npm run lint           # Run ESLint
npm run test           # Run tests

# Debugging
npm run dev-server     # Start webpack dev server only
```

## 🔄 **Next Steps**

1. **Authentication**: Implement user login/logout flow
2. **Case Selection**: Add dropdown for user's cases
3. **LiraDocs API**: Integrate with LiraDocs server for .eml upload
4. **Badge Enhancement**: Add visual badges in email list
5. **Error Handling**: Improve error messages and retry logic

## 🐛 **Troubleshooting**

### Common Issues
- **CORS Errors**: Ensure Cloudflare Worker has proper CORS headers
- **Office.js Not Loading**: Check manifest.xml permissions
- **Build Errors**: Run `npm install` to update dependencies

### Development Tips
- Use F12 Developer Tools in Outlook Web
- Check browser console for JavaScript errors
- Validate manifest.xml before deployment

## 📚 **Resources**

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Fluent UI React Components](https://react.fluentui.dev/)
- [Office.js API Reference](https://docs.microsoft.com/en-us/javascript/api/office)
- [Cloudflare Workers Documentation](https://developers.cloudflare.com/workers/)

---

**Built with ❤️ for LiraDocs Legal Case Management**
