# ExcelSQLApp Pro - SQL Editor & Query Tool

![Python](https://img.shields.io/badge/python-3.7+-blue.svg)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)

![App Screenshot](screenshot.png) <!-- Add actual screenshot later -->

A powerful SQL editor tool with Excel integration featuring:
- SQL query execution with syntax highlighting
- Multi-database Excel file management
- Advanced spooling capabilities
- Enhanced query history and export functionality

## Features ✨

### Core Functionality
✔️ **SQL Query Editor** with syntax highlighting  
✔️ **Multi-file Excel integration** (load multiple workbooks simultaneously)  
✔️ **Smart query execution** (run selected text or full queries)  
✔️ **Query history** with quick recall functionality  

### Enhanced Features
🎯 **Case-insensitive SQL validation** (ignores keywords in comments)  
🚀 **Horizontal scrolling** for wide result sets  
📋 **Right-click context menus** (copy cells/columns)  

### Spooling System
📁 **Output to CSV/TXT** with timestamps  
🔍 **Multi-query support** (separate queries with semicolons)  
⚡ **Live spooling toggle** (Start/Stop anytime)  

## Installation 🛠️

1. **Prerequisites**:
   - Python 3.7+
   - Required packages:
     ```bash
     pip install pandas tkinter sqlite3

Usage Guide 📖
Basic Operation
Load Excel files via Browse Excel Files button
Write or paste SQL queries in the editor
Execute with:
▶ Execute button
Ctrl+Enter (for selected text)
Spooling Features
Click 🔴 Start Spooling to begin recording
Select output file location
Execute queries (results auto-saved)
Click ✅ Stop Spooling when done
Advanced Tips
Use ; to separate multiple queries in one execution
Right-click result grid for quick copy options
Export final results via 💾 Export to Excel
Development 🧑💻
Project Structure

Run
Copy code
ExcelSQLApp-Pro/
├── ExcelSQLApp.py      # Main application
├── LICENSE
├── README.md
└── requirements.txt
Contributing
Fork the project
Create your feature branch (git checkout -b feature/AmazingFeature)
Commit changes (git commit -m 'Add AmazingFeature')
Push to branch (git push origin feature/AmazingFeature)
Open a Pull Request
     ```

2. **Clone repository**:
   ```bash
   git clone https://github.com/yourusername/ExcelSQLApp-Pro.git
   cd ExcelSQLApp-Pro
