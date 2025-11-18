# PowerPoint Automation Tool

## Overview
This project automates the generation of PowerPoint reports by:
1. Inserting images into PowerPoint templates based on shape names
2. Fetching data from Grafana dashboards and inserting statistics into presentations
3. **Multi-user support** with authentication and role-based access control
4. **Work history tracking** for audit and report retrieval

## Project Structure
```
.
├── app.py                       # Streamlit web GUI with authentication (권장)
├── config.py                    # Configuration settings
├── main.py                      # CLI interface
├── insert_images.py             # Step 1: Image insertion script
├── numinsert3.py                # Step 2: Grafana statistics insertion script
├── db_models.py                 # Database models (SQLAlchemy)
├── db_utils.py                  # Database utility functions
├── create_admin.py              # Admin user creation script
├── requirements.txt             # Python dependencies
├── .streamlit/
│   ├── config.toml              # Streamlit configuration
│   └── secrets.toml             # Grafana API configuration (git-ignored)
├── Report/
│   ├── template/                # Place your PowerPoint templates here
│   ├── completed_with_images/   # Intermediate output (Step 1)
│   └── completed_final/         # Final output (Step 2)
└── replit.md                    # This file
```

## Setup Instructions

### 1. Directory Structure
The tool requires the following directories (created automatically on first run):
- `Report/template/` - Place your PowerPoint template files (.pptx) here
- `Report/` - Place image files (.png, .jpg, .jpeg, .gif) here (organized by customer folders)
- `Report/completed_with_images/` - Intermediate files after image insertion
- `Report/completed_final/` - Final reports with statistics

### 2. Grafana Configuration
If you want to use the Grafana statistics feature, set these environment variables:
- `GRAFANA_URL` - Your Grafana server URL (default: http://localhost:3000)
- `GRAFANA_API_KEY` - Your Grafana API key
- `GRAFANA_VERIFY_SSL` - Whether to verify SSL certificates (default: true)
  - Set to "false" only if using self-signed certificates or testing environments
  - For production, always keep this as "true" for security

To set environment variables in Replit:
1. Click on "Tools" in the left sidebar
2. Select "Secrets"
3. Add your secrets with the key names above

### 3. Customer Dashboard Mapping
Edit `config.py` to update the `DASHBOARD_MAP` dictionary with your customer names and Grafana dashboard UIDs.

## How to Use

### Using the Web GUI (권장)
웹 브라우저에서 사용할 수 있는 GUI가 제공됩니다:
```bash
streamlit run app.py
```

Replit에서는 자동으로 실행되며, 브라우저에서 바로 사용하실 수 있습니다.

**First-time setup:**
1. Run the following command to create an admin account:
   ```bash
   python create_admin.py
   ```
2. Enter your desired username and secure password when prompted
3. ⚠️ **Never use default or weak passwords in production!**

GUI에서는 다음 기능을 제공합니다:
- 🔐 로그인/로그아웃: 사용자 인증
- 📊 홈 대시보드: 전체 상태 확인
- 🖼️ 이미지 삽입: 템플릿에 이미지 삽입
- 📈 통계 삽입: Grafana 통계 데이터 삽입
- 📂 작업 이력: 과거 생성된 보고서 조회 및 다운로드
- 👥 사용자 관리 (관리자 전용): 사용자 계정 생성/관리
- ⚙️ 설정: 환경 설정 확인 및 관리

### Using the CLI
터미널에서 대화형 메뉴를 사용하려면:
```bash
python main.py
```

The menu provides the following options:
1. **Image Insertion (Step 1)** - Process templates and insert images
2. **Grafana Statistics Insertion (Step 2)** - Add Grafana data to presentations
3. **Full Process** - Run both steps sequentially
4. **Create Directories** - Set up the directory structure
5. **View Configuration** - Display current settings
6. **Exit** - Close the program

### Running Scripts Directly

#### Image Insertion Only
```bash
python insert_images.py
```

#### Statistics Insertion Only
```bash
# Process all customers
python numinsert3.py

# Process specific customer
python numinsert3.py customer_name
```

## Workflow

### Step 1: Image Insertion
1. Places PowerPoint templates in `Report/template/`
2. Organizes images in customer-specific folders
3. Script matches shape names to image filenames
4. Outputs to `Report/completed_with_images/`

### Step 2: Statistics Insertion
1. Takes files from `Report/completed_with_images/`
2. Queries Grafana dashboards for metrics
3. Replaces placeholders like `{{panel-name_A}}` with statistics
4. Replaces date placeholders ({{START_DATE}}, {{END_DATE}}, etc.)
5. Outputs final reports to `Report/completed_final/`

## Placeholder Format

### Date Placeholders
- `{{START_DATE}}` - Start date in Korean format
- `{{END_DATE}}` - End date in Korean format
- `{{MONTH}}` - Month number
- `{{DATE_RANGE}}` - Full date range in Korean format
- `{{DATE_RANGE_HYPHEN}}` - Full date range with hyphens

### Grafana Statistics Placeholders
Format: `{{panel-name_QueryLetter}}`
Example: `{{CPU-Usage_A}}`

The script will:
1. Find the panel with matching title
2. Query the specified query letter (A, B, C, etc.)
3. Calculate max and mean values
4. Replace with: "사용량 최대 X%, 평균 Y% 입니다."

## Configuration

### config.py
Key settings:
- `BASE_TEMPLATE_DIR` - Template directory
- `OUTPUT_DIR_WITH_IMAGES` - Intermediate output
- `OUTPUT_DIR` - Final output
- `GRAFANA_URL` - Grafana server URL
- `API_KEY` - Grafana API key
- `DASHBOARD_MAP` - Customer to dashboard UID mapping
- `SENTENCE_TEMPLATE` - Output format for statistics

## Dependencies
- python-pptx==0.6.23 - PowerPoint manipulation
- python-dateutil==2.8.2 - Date calculations
- requests==2.31.0 - HTTP requests for Grafana API
- urllib3==2.1.0 - HTTP client
- streamlit - Web GUI framework
- streamlit-authenticator - User authentication
- sqlalchemy - Database ORM
- psycopg2-binary - PostgreSQL adapter
- bcrypt - Password hashing
- pyyaml - Configuration file parsing

## Troubleshooting

### No templates found
- Ensure .pptx files are in `Report/template/`
- Check file permissions
- Run option 4 to create directories

### Images not inserting
- Verify image filenames match shape names (case-insensitive, special characters ignored)
- Ensure images are in the correct customer folder
- Check image file extensions (.png, .jpg, .jpeg, .gif)

### Grafana queries failing
- Verify GRAFANA_URL is correct
- Check GRAFANA_API_KEY is valid
- Ensure dashboard UIDs in DASHBOARD_MAP are correct
- Verify panel titles match placeholder names

### Environment Variables
Set these in Replit Secrets if needed:
- `GRAFANA_URL` - Grafana server URL
- `GRAFANA_API_KEY` - Grafana API token
- `GRAFANA_VERIFY_SSL` - SSL verification (default: true, set to "false" only for testing)

## Database Schema

### Users Table
- **id**: Primary key
- **username**: Unique username
- **email**: Email address
- **password_hash**: Bcrypt hashed password
- **full_name**: User's full name
- **role**: 'admin' or 'user'
- **is_active**: Account status
- **created_at**, **last_login**: Timestamps

### Report Runs Table
- **id**: Primary key
- **user_id**: Foreign key to users
- **customer_name**: Customer/project name
- **report_type**: 'images', 'stats', or 'full'
- **template_name**: Template file name
- **status**: 'success' or 'failed'
- **created_at**, **duration_seconds**: Execution metadata
- **log_data**: Execution logs (JSON)

### Report Files Table
- **id**: Primary key
- **run_id**: Foreign key to report_runs
- **filename**: Generated file name
- **file_path**: Full file path
- **file_size**: Size in bytes
- **step**: 'step1' or 'step2'
- **created_at**: Timestamp

## User Management

### Creating New Users
Administrators can create new users through the "사용자 관리" tab in the web GUI, or use the command-line tool:
```bash
python create_admin.py
```

### User Roles
- **Admin**: Full access including user management and all user reports
- **User**: Access to own reports and standard features

## Recent Changes
- 2025-11-17: Initial setup for Replit environment
  - Converted hardcoded /root paths to configurable paths
  - Created CLI interface (main.py) with interactive menu and non-interactive flags
  - Added Streamlit web GUI (app.py) for easy browser-based usage
  - Added config.py for centralized configuration
  - Set up proper directory structure
  - Added comprehensive documentation
  - Made SSL verification configurable (defaults to secure)
  - Created workflow that runs Streamlit on port 5000

- 2025-11-17: Multi-user authentication system
  - **PostgreSQL database integration** for user management
  - **Session-based authentication** with bcrypt password hashing
  - **Role-based access control** (admin/user roles)
  - **User management page** for administrators
  - **Work history tracking** to log all report generation activities
  - **Report file archiving** with download capability
  - Created database models (users, report_runs, report_files)
  - Added admin account creation tool (create_admin.py)

- 2025-11-18: Enhanced statistics insertion workflow
  - **Optimized login performance** with SQLAlchemy connection pooling (reduced overhead by 200-600ms)
  - **Grafana configuration editor** in Settings tab with connection testing
  - **Improved statistics insertion UI/UX**:
    - Dual-button upload: "저장만 하기" and "저장 후 바로 통계 삽입" for flexible workflows
    - Simplified execution: Default to "전체 고객사 처리", optional specific customer selection
    - Enhanced file download section with full sub-folder structure (GIT/GIT2/GIT3/GIT4 displayed separately)
    - Structured log display: Shows processed file count, failed placeholders, and generated file paths
  - **Sub-customer folder support**: Properly handles nested structures like GIT/GIT2/GIT3/GIT4
  - **Smart Grafana API calls**: Only queries panels that have placeholders in the template (no unnecessary API calls)
