# 🚀 Google Sheets Task Tracker & JSON Converter

Advanced Google Sheets automation tool that combines JSON data conversion with comprehensive progress tracking and analytics for development teams.

## ✨ Features

### 🔄 JSON Conversion (Add to Sheet)
- **Active Row Conversion**: Convert currently selected row to JSON format
- **Batch Processing**: Convert selected rows or all data rows simultaneously
- **Smart Column Handling**: Automatically excludes `worker_id`, `status`, and `complete_json` columns
- **Automatic Placement**: Places JSON in designated `complete_json` column
- **Tracker Integration**: Updates progress tracker after each conversion

### 💾 JSON Export (Download Files)
- **Direct File Downloads**: Create actual downloadable .json files (not sheet tabs)
- **Multiple Export Options**: Export active row, selected rows, or all data
- **Google Drive Integration**: Files saved to Drive with direct download links
- **Custom File Naming**: Set custom names for exported files
- **Metadata Inclusion**: Adds `_row_number` and `_exported_at` timestamps
- **Interactive Export Wizard**: Guided export process with prompts
- **Quick Export**: One-click export of all data rows

### 📊 Progress Tracking
- **Real-time Analytics**: Track completion rates, success rates, and daily averages
- **Project-Based Metrics**: Configure project start date for accurate daily averages
- **Daily Performance**: Monitor completions for the last 7 days by worker
- **Weekly Statistics**: Track progress over the last 4 weeks
- **Summary Dashboard**: Overview with top performers and key metrics
- **Visual Indicators**: Color-coded performance levels with conditional formatting
- **Multiple Status Recognition**: Supports various status formats (complete/done/finished/success)

### ⚙️ Project Configuration
- **Project Start Date**: Set and manage project timeline for accurate metrics
- **Smart Daily Averages**: Formula: `(Completed + Pending Tasks) ÷ Days Since Start`
- **Capacity Planning**: Includes both finished work and assigned workload
- **Project Info Display**: Shows calculation methods and project duration
- **Settings Management**: View, update, or reset project configuration

### 📈 Advanced Analytics
- **Performance Reports**: Comprehensive analysis with trends and insights
- **Productivity Insights**: Automated recommendations based on data patterns
- **Top Performers Ranking**: Identify high-performing team members
- **Completion Trend Analysis**: 30-day historical completion patterns
- **Worker Statistics**: Individual performance breakdowns
- **Export Capabilities**: Export tracker data to separate spreadsheets

### ⏰ Automation & Scheduling
- **Daily Auto-Updates**: Scheduled tracker refresh at 9:00 AM
- **Email Notifications**: Optional stakeholder alerts with summary statistics
- **Trigger Management**: Easy setup and removal of automated processes
- **Background Processing**: Automatic updates without manual intervention

### 🔧 Debug & Troubleshooting
- **Timestamp Detection**: Automatic detection of date/time columns
- **Column Analysis**: Debug tool to verify timestamp column recognition
- **Data Validation**: Comprehensive error handling and user feedback
- **Status Recognition**: Flexible status categorization system

## 📋 Requirements

### Required Columns
Your Google Sheet must contain these columns:
- `worker_id` - Unique identifier for each team member
- `status` - Task status (completed/pending/in progress/failed)
- `complete_json` - Target column for JSON output (will be created if missing)

### Optional Columns (for Enhanced Features)
- **Timestamp Column**: Any column with names like:
  - `timestamp`, `date`, `created_at`, `updated_at`, `completed_at`
  - `last_updated`, `time`, `datetime`, `completion_date`, `task_date`
  - Required for daily/weekly statistics and trend analysis

### Supported Status Values
- **Completed**: `complete`, `done`, `finished`, `success`
- **In Progress**: `progress`, `working`, `active`, `ongoing`  
- **Failed**: `fail`, `error`, `reject`, `cancel`
- **Pending**: Everything else (default category)

## 🚀 Installation

1. **Open Google Sheets** and create or open your project spreadsheet
2. **Go to Extensions** → **Apps Script**
3. **Delete default code** and paste the complete script
4. **Save the project** (Ctrl+S)
5. **Refresh your Google Sheet** - the menu will appear automatically

## 📖 Usage Guide

### Initial Setup
1. **Set Project Start Date**: 
   - Go to `⚙️ Project Settings` → `📅 Set Project Start Date`
   - Enter your project start date for accurate daily averages
   - Format: `MM/DD/YYYY` or `YYYY-MM-DD`

2. **Create Progress Tracker**:
   - Go to `📊 Progress Tracker` → `Create Tracker Tab`
   - This creates a comprehensive tracking dashboard

### JSON Operations

#### Adding JSON to Sheet
- **Single Row**: Select a row → `📝 JSON Conversion` → `Convert Active Row to JSON`
- **Multiple Rows**: Select rows → `Convert Selected Rows to JSON`
- **All Data**: `Convert All Rows to JSON`

#### Downloading JSON Files
- **Quick Export**: `💾 JSON Export` → `🔥 Quick Export All Rows`
- **Custom Export**: Use `📝 Create JSON File (Interactive)` for guided process
- **Custom Naming**: `🏷️ Export with Custom Name` for personalized file names

### Tracking & Analytics

#### Daily Monitoring
- **Refresh Tracker**: `📊 Progress Tracker` → `Refresh Tracker`
- **View Daily Stats**: Check the DAILY PERFORMANCE section
- **Check Weekly Trends**: Review WEEKLY PERFORMANCE data

#### Performance Analysis
- **Generate Reports**: `📈 Analytics & Reports` → `Generate Performance Report`
- **Export Data**: Create standalone tracker exports for stakeholders
- **View Project Settings**: Monitor project duration and calculation methods

### Automation Setup
1. **Enable Auto-Updates**: `⏰ Automation` → `Setup Daily Auto-Update`
2. **Configure Notifications**: Edit the `automaticTrackerUpdate()` function to add email recipients
3. **Manage Triggers**: Use `Remove Auto-Update` to disable automation

## 🎯 Key Metrics Explained

### Daily Average Calculation
**Formula**: `(Completed + Pending Tasks) ÷ Days Since Project Start`

**Includes**:
- ✅ **Completed Tasks**: Finished work showing productivity
- ⏳ **Pending Tasks**: Assigned/queued work showing capacity

**Excludes**:
- 🔄 **In Progress**: Temporary state (currently being worked)
- ❌ **Failed/Rejected**: Unsuccessful attempts

**Example**: 
- 50 completed + 10 pending = 60 productive tasks
- Project running 150 days
- Daily Average: 60 ÷ 150 = 0.40 tasks/day

### Success Rate vs Completion Rate
- **Completion Rate**: `Completed Tasks ÷ Total Tasks`
- **Success Rate**: `Completed Tasks ÷ (Total Tasks - Pending Tasks)`

## 🔧 Troubleshooting

### Common Issues

**Daily/Weekly Stats Not Working**:
1. Use `🔧 Debug Tools` → `Check Timestamp Columns`
2. Ensure you have a date/time column with proper naming
3. Verify date formats are recognizable (YYYY-MM-DD, MM/DD/YYYY, etc.)

**JSON Export Not Working**:
- Check Google Drive permissions
- Ensure `complete_json` column exists
- Verify you have data rows (not just headers)

**Tracker Not Updating**:
- Confirm `worker_id` and `status` columns exist
- Check that status values match supported formats
- Use `Refresh Tracker` to force update

**Daily Average Showing Zero**:
1. Set project start date: `⚙️ Project Settings` → `📅 Set Project Start Date`
2. Ensure you have completed or pending tasks
3. Check `👁️ View Project Settings` for configuration details

### Debug Features
- **Check Timestamp Columns**: Analyzes your sheet's date columns
- **View Project Settings**: Shows current configuration and calculations
- **Manual Refresh**: Force update tracker data
