# PowerPoint Automation Tool

## Overview
This project automates the generation of PowerPoint reports by:
1. Inserting images into PowerPoint templates based on shape names
2. Fetching data from Grafana dashboards and inserting statistics into presentations

## Project Structure
```
.
├── app.py                       # Streamlit web GUI (권장)
├── config.py                    # Configuration settings
├── main.py                      # CLI interface
├── insert_images.py             # Step 1: Image insertion script
├── numinsert3.py                # Step 2: Grafana statistics insertion script
├── requirements.txt             # Python dependencies
├── .streamlit/
│   └── config.toml              # Streamlit configuration
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

GUI에서는 다음 기능을 제공합니다:
- 📊 홈 대시보드: 전체 상태 확인
- 🖼️ 이미지 삽입: 템플릿에 이미지 삽입
- 📈 통계 삽입: Grafana 통계 데이터 삽입
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

## Recent Changes
- 2025-11-17: Initial setup for Replit environment
  - Converted hardcoded /root paths to configurable paths
  - Created CLI interface (main.py) with interactive menu and non-interactive flags
  - **Added Streamlit web GUI (app.py) for easy browser-based usage**
  - Added config.py for centralized configuration
  - Set up proper directory structure
  - Added comprehensive documentation
  - Made SSL verification configurable (defaults to secure)
  - Created workflow that runs Streamlit on port 5000
