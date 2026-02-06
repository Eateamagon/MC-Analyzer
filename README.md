# Mastery Connect Analyzer Web App v1.0

A comprehensive web application for analyzing Mastery Connect benchmark assessment data for Waynesboro Public Schools.

## 📊 Features

### Core Analysis Features
- **CSV Upload** - Drag-and-drop file upload from Mastery Connect exports
- **Class Period Entry** - Assign class periods to students after upload
- **Item Analysis** - Question difficulty, distractor reports, high-performer alerts
- **Small Group Interventions** - Automated grouping by class period with shared weak SOLs
- **Heatmaps** - Visual student & SOL performance matrices
- **Teacher Summaries** - SOL categorization (Growth/Monitor/Strength)

### Admin Features
- **District Summary** - Overview of all teachers' performance
- **Separate Teacher Files** - Create individual Google Sheets for each teacher
- **User Management** - View and manage division users
- **Admin List Management** - Configure division and school administrators

### Export Options
- **Excel Download** - XLSX format with all analysis results
- **PDF Reports** - Printable PDF format
- **Google Drive Integration** - Organized folder structure for teacher files
- **Email Notifications** - Optional email when analysis completes

## 🔐 User Roles

| Role | Access |
|------|--------|
| **Teacher** | Upload own data, run analysis, view own results, download reports |
| **School Admin** | Everything teachers can do + view school-wide data |
| **Division Admin** | Everything + view all schools, manage admins, create teacher files |

## 🏫 Schools
- Berkeley Glenn Elementary School
- Wenonah Elementary School
- Westwood Hills Elementary School
- William Perry Elementary School
- Kate Collins Middle School
- Waynesboro High School

## 📋 Deployment Instructions

### Step 1: Create the Google Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click **New Project**
3. Name the project "Mastery Connect Analyzer"

### Step 2: Add Script Files

Create the following `.gs` files and copy the content:

1. **Code.gs** - Main application code
2. **AnalysisFunctions.gs** - Analysis functions
3. **ExportFunctions.gs** - Export and email functions
4. **AdminFunctions.gs** - Admin-specific functions

### Step 3: Add HTML Files

Create the following `.html` files:

1. **Dashboard.html** - Main dashboard page
2. **Upload.html** - CSV upload page
3. **Analysis.html** - Analysis results page
4. **History.html** - Assessment history page
5. **Admin.html** - Admin panel page
6. **Login.html** - Login/registration page
7. **Styles.html** - CSS styles (include file)
8. **Scripts.html** - JavaScript utilities (include file)

### Step 4: Configure OAuth Scopes

In the script editor, go to **Project Settings** and ensure the following scopes are authorized:

```
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/script.send_mail
https://www.googleapis.com/auth/userinfo.email
```

### Step 5: Deploy as Web App

1. Click **Deploy** > **New deployment**
2. Select type: **Web app**
3. Configure:
   - **Description**: "Mastery Connect Analyzer v1.0"
   - **Execute as**: "User accessing the web app"
   - **Who has access**: "Anyone with Waynesboro Google account" (or your domain)
4. Click **Deploy**
5. Copy the Web app URL

### Step 6: Initial Setup

1. Open the web app URL
2. The first user will be prompted to register
3. Add the first division admin email to the Settings sheet in the auto-created database
4. The database spreadsheet will be created automatically in the user's Drive

## ⚙️ Configuration

### Adding Administrators

1. Open the database spreadsheet
2. Go to the "Settings" sheet
3. Add rows with format:
   - Column A: `division_admin` or `school_admin`
   - Column B: Email address
   - Column C: School name (for school admins only)

### Default Thresholds

The default analysis thresholds can be modified in the Admin panel:
- **Growth**: 50% (scores below this need growth)
- **Strength**: 75% (scores at or above are strengths)
- **Small Group Weakness**: 70%
- **Max Group Size**: 5

## 📁 File Structure

```
mastery-connect-app/
├── Code.gs                # Main web app entry points & user management
├── AnalysisFunctions.gs   # All analysis functions
├── ExportFunctions.gs     # Export to Excel/PDF/Google Sheets
├── AdminFunctions.gs      # Admin-specific functions
├── Dashboard.html         # Main dashboard
├── Upload.html            # CSV upload wizard
├── Analysis.html          # Analysis results viewer
├── History.html           # Assessment history
├── Admin.html             # Admin panel
├── Login.html             # Login/registration
├── Styles.html            # CSS styles
└── Scripts.html           # JavaScript utilities
```

## 📊 CSV Format

The app expects Mastery Connect CSV exports with:

**Row 1**: SOL standards (in question columns)
**Row 2**: Headers including:
- district, school, teacher, tracker, school_year
- assessment_name, student_id, state_number
- last_name, first_name, student_status, created_at
- points_possible, score, percentage
- Followed by alternating answer/score columns (answer1, score1, answer2, score2, ...)

## 🎨 Theme

- **Primary Color**: Mastery Connect Green (#00a14b)
- **Background**: Dark theme (#1a1a2e)
- **Style**: Clean, card-based layout with hover animations

## 📈 Future Enhancements (V2)

The following features are planned for future releases:
- SOL trend analysis over time (Q1→Q3 benchmark comparisons)
- Individual student growth tracking
- Teacher performance vs. division average
- Advanced filtering and search
- Bulk period assignment improvements

## 🆘 Support

For questions or issues, contact the Division ITRT.

---

**Version**: 1.0  
**Last Updated**: January 2026  
**Author**: Waynesboro Public Schools IT Department
