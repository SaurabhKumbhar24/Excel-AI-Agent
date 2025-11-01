# Excel AI Agent 🤖📊

An intelligent Excel add-in that brings the power of AI to your spreadsheets. Interact with Excel using natural language and let AI handle complex formulas, pivot tables, charts, and data manipulation tasks.

![Excel AI Agent Demo](./Excel%20AI%20Agent%20-%20Made%20with%20Clipchamp.mp4)

## 🌟 Features

- **Natural Language Processing**: Ask questions in plain English like "create a formula to sum A1 to A10" or "make a pivot table showing sales by region"
- **Smart Formula Generation**: Automatically generates Excel formulas based on your description
- **Pivot Table Creation**: Create complex pivot tables with simple natural language commands
- **Chart Generation**: Generate various chart types (bar, line, pie, etc.) from your data
- **Data Filtering & Sorting**: Filter and sort data using conversational commands
- **Context-Aware**: Analyzes your selected ranges, sheet data, and headers to provide relevant responses
- **Auto-Execution**: Generates and executes Office.js code to perform actions directly in Excel

## 🏗️ Architecture

### Frontend
- **Technology**: React + TypeScript
- **Framework**: Office.js Add-in
- **UI Library**: Fluent UI React Components
- **Build Tool**: Webpack

### Backend
- **Framework**: FastAPI (Python)
- **AI Model**: Google Gemini AI
- **API**: RESTful API with CORS support

### Communication Flow
```
User Input → Excel Add-in (React) → FastAPI Backend → Gemini AI → 
Excel Interpreter → Office.js Code → Excel Workbook
```

## 📋 Prerequisites

- **Node.js** (v14 or higher)
- **Python** (v3.8 or higher)
- **Microsoft Excel** (Desktop or Online)
- **Google Gemini API Key** ([Get one here](https://ai.google.dev/))

## 🚀 Installation

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/excel-ai-agent.git
cd excel-ai-agent
```

### 2. Backend Setup

Navigate to the backend directory and install dependencies:

```bash
cd backend
python -m venv venv

# On Windows
venv\Scripts\activate

# On macOS/Linux
source venv/bin/activate

pip install -r requirements.txt
```

Create a `.env` file in the `backend` directory:

```env
GEMINI_API_KEY=your_gemini_api_key_here
```

Start the backend server:

```bash
uvicorn app.main:app --reload --port 8000
```

The API will be available at `http://localhost:8000`

### 3. Frontend Setup

Navigate to the Excel add-in directory:

```bash
cd excel-ai-agent
npm install
```

Install development certificates (required for HTTPS):

```bash
npx office-addin-dev-certs install
```

Start the development server:

```bash
npm run dev-server
```

The add-in will be available at `https://localhost:3000`

### 4. Sideload the Add-in

**For Excel Desktop:**
```bash
npm start
```

This will automatically open Excel and sideload the add-in.

**For Manual Sideloading:**
1. Open Excel
2. Go to Insert → Add-ins → My Add-ins
3. Click "Upload My Add-in"
4. Select the `manifest.xml` file from the `excel-ai-agent` directory

## 💡 Usage

1. **Open the Task Pane**: Click the "Show Task Pane" button in the Excel ribbon
2. **Select Data** (Optional): Select a range in your Excel sheet for context
3. **Ask a Question**: Type your request in natural language, for example:
   - "Create a SUM formula for column A"
   - "Make a pivot table showing total sales by region"
   - "Create a bar chart from this data"
   - "Sort this data by column B in descending order"
   - "Filter rows where sales are greater than 1000"
4. **Execute**: The AI will interpret your request and automatically perform the action in Excel

## 🎯 Example Queries

### Formulas
- "Create a formula to calculate the average of B2 to B10"
- "Add a VLOOKUP formula to find values from another sheet"
- "Generate an IF formula to check if values are greater than 100"

### Pivot Tables
- "Create a pivot table with products as rows and sum of sales"
- "Make a pivot table showing average revenue by month and category"

### Charts
- "Create a bar chart showing sales by region"
- "Generate a line chart for monthly trends"
- "Make a pie chart of market share"

### Data Manipulation
- "Sort this data by date in ascending order"
- "Filter rows where status is 'Active'"
- "Remove duplicates from column A"

## 🛠️ Development

### Project Structure

```
excel-ai-agent/
├── backend/
│   ├── app/
│   │   ├── main.py              # FastAPI application
│   │   ├── routers/
│   │   │   └── ai_routers.py    # API endpoints
│   │   └── services/
│   │       ├── ai_service.py    # Gemini AI integration
│   │       ├── excel_interpreter.py  # Converts AI to Office.js
│   │       └── formula_generator.py  # Formula generation logic
│   └── requirements.txt
│
├── excel-ai-agent/
│   ├── src/
│   │   ├── taskpane/
│   │   │   ├── components/
│   │   │   │   └── App.tsx      # Main React component
│   │   │   ├── taskpane.html
│   │   │   └── taskpane.ts
│   │   └── commands/
│   ├── manifest.xml             # Add-in manifest
│   ├── package.json
│   └── webpack.config.js
│
└── README.md
```

### Available Scripts

**Frontend:**
- `npm run dev-server` - Start development server
- `npm run build` - Build for production
- `npm start` - Start and sideload add-in in Excel
- `npm run validate` - Validate manifest.xml

**Backend:**
- `uvicorn app.main:app --reload` - Start development server
- `uvicorn app.main:app --reload --port 8000` - Start on specific port

## 🔧 Configuration

### Backend Configuration
Edit `backend/app/main.py` to configure CORS, API settings, etc.

### Frontend Configuration
Edit `excel-ai-agent/manifest.xml` to customize:
- Add-in name and description
- Icons and branding
- Permissions
- Supported Office hosts

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Built with [Office.js](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- Powered by [Google Gemini AI](https://ai.google.dev/)
- UI components from [Fluent UI](https://react.fluentui.dev/)
- Backend framework: [FastAPI](https://fastapi.tiangolo.com/)

## 📧 Support

For issues, questions, or suggestions, please open an issue on GitHub.

## 🔮 Future Enhancements

- [ ] Support for more complex Excel operations
- [ ] Multi-language support
- [ ] Custom AI model training
- [ ] Batch operations
- [ ] Export/Import AI command templates
- [ ] Integration with other Office applications (Word, PowerPoint)

---

Made with ❤️ by [Your Name]

