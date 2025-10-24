# Power BI MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to interact with Power BI workspaces, datasets, reports, and dashboards programmatically.

## 🚀 Features

- **Workspace Management**: List and manage Power BI workspaces
- **Report Operations**: Access, clone, export, and rebind reports
- **Dataset Management**: Execute DAX queries, refresh datasets, manage schedules
- **Dashboard Access**: List and interact with dashboards
- **Push Datasets**: Create and manage push datasets with real-time data
- **Authentication**: Secure OAuth2 authentication with Microsoft Entra ID

## 📋 Prerequisites

- Node.js (v18 or higher)
- npm or yarn
- Power BI Pro or Premium license
- Microsoft Entra ID app registration
- MCP-compatible AI assistant (Claude Desktop, etc.)

## 🔧 Installation

1. Clone the repository:
```bash
git clone https://github.com/michaelmckinleyconsulting/powerbi-mcp-server.git
cd powerbi-mcp-server
```

2. Install dependencies:
```bash
npm install
```

3. Build the project:
```bash
npm run build
```

## ⚙️ Configuration

### 1. Microsoft Entra ID Setup

1. Register an application in [Azure Portal](https://portal.azure.com)
2. Configure API permissions for Power BI Service
3. Note your:
   - Client ID
   - Tenant ID
   - Client Secret (if using app-only auth)

### 2. Environment Variables

Create a `.env` file in the root directory (see `.env.example` for template):

```env
PBI_PUBLIC_CLIENT_ID=your_client_id
PBI_TENANT_ID=your_tenant_id  # Optional, defaults to 'common'
PBI_SCOPES=https://analysis.windows.net/powerbi/api/.default  # Optional, or specify custom scopes
```

**Note**: This server uses Authorization Code flow with PKCE (no client secret required). Configure your Azure app with redirect URIs:
- `http://localhost` 
- `http://127.0.0.1`

### 3. MCP Configuration

Add to your MCP client configuration:

#### Claude Desktop
```json
{
  "mcpServers": {
    "powerbi": {
      "command": "node",
      "args": ["path/to/powerbi-mcp-server/dist/index.js"],
      "env": {
        "PBI_PUBLIC_CLIENT_ID": "your_client_id",
        "PBI_TENANT_ID": "your_tenant_id"
      }
    }
  }
}
```

#### VS Code MCP Extension
```json
{
  "mcp.servers": {
    "powerbi": {
      "command": "node",
      "args": ["path/to/powerbi-mcp-server/dist/index.js"],
      "transport": "stdio",
      "env": {
        "PBI_PUBLIC_CLIENT_ID": "your_client_id"
      }
    }
  }
}
```

## 📖 Usage

### Basic Operations

Once configured, your AI assistant can:

- **List workspaces**: "Show me all my Power BI workspaces"
- **Execute DAX**: "Run a DAX query to get sales by region"
- **Export reports**: "Export the quarterly report to PDF"
- **Refresh datasets**: "Refresh the sales dataset"
- **Manage access**: "Add user@example.com to the Analytics workspace"

### Example Prompts

```
"List all reports in my Sales workspace"
"Execute this DAX query against the Finance dataset: EVALUATE SUMMARIZE(...)"
"Export the Monthly KPI report to PowerPoint"
"Show me the refresh history for the Customer dataset"
"Create a push dataset for real-time monitoring"
```

## 🔐 Authentication

The server supports two authentication methods:

1. **Interactive (Recommended)**: User signs in via browser
2. **App-Only**: Uses client credentials (requires admin consent)

On first run, the server will prompt for authentication via your default browser.

## 📁 Project Structure

```
powerbi-mcp-server/
├── src/
│   ├── index.ts           # Main server entry point
│   ├── auth/              # Authentication logic
│   ├── handlers/          # Request handlers
│   ├── types/             # TypeScript definitions
│   └── utils/             # Utility functions
├── dist/                  # Compiled JavaScript (git-ignored)
├── package.json
├── tsconfig.json
└── README.md
```

## 🛠️ Development

### Running in Development Mode

```bash
npm run dev
```

### Running Tests

```bash
npm test
```

### Building for Production

```bash
npm run build
```

## 📊 Supported Power BI Operations

### Workspaces (Groups)
- List workspaces
- Get workspace users
- Add/remove users

### Reports
- List reports
- Get report metadata
- Clone reports
- Export to various formats (PDF, PPTX, PNG)
- Rebind to different datasets

### Datasets
- List datasets
- Execute DAX queries
- Trigger refreshes
- Manage refresh schedules
- Get refresh history

### Dashboards
- List dashboards
- Get dashboard tiles

### Push Datasets
- Create push datasets
- Add rows to tables
- Clear table data

## 🔧 Troubleshooting

### Common Issues

1. **Authentication fails**: Ensure your Azure app registration has proper Power BI permissions
2. **No workspaces found**: Verify your account has access to Power BI workspaces
3. **DAX queries fail**: Check dataset permissions and query syntax
4. **Export timeouts**: Large reports may take time; the server handles async polling

### Debug Mode

Enable verbose logging:

```bash
DEBUG=powerbi:* node dist/index.js
```

## 📄 License

This project is **source-available** under a Modified MIT License with Commons Clause - see the [LICENSE](LICENSE) file for details.

**Non-Commercial Use**: Free for personal, educational, and internal company use.

**Commercial Use**: Requires a separate commercial license. Contact [michael@mckinley.consulting] for commercial licensing.

This is not OSI-approved "open source" due to the commercial restrictions, but allows full community use for non-commercial purposes.

## 🤝 Contributing

We welcome contributions! Please note:

1. By contributing, you agree that your contributions will be licensed under the same license
2. For substantial changes, please open an issue first to discuss
3. Ensure all tests pass and add tests for new features

## 💬 Support

- **Issues**: [GitHub Issues](https://github.com/michaelmckinleyconsulting/powerbi-mcp-server/issues)
- **Commercial Support**: [michael@mckinley.consulting]
- **Documentation**: [Wiki](https://github.com/michaelmckinleyconsulting/powerbi-mcp-server/wiki)

## 🙏 Acknowledgments

- Built on the [Model Context Protocol](https://modelcontextprotocol.io) specification
- Uses Microsoft Power BI REST APIs
- Inspired by the MCP community

---

**Note**: This is not an official Microsoft product. Power BI is a trademark of Microsoft Corporation.
