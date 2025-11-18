# PowerPoint Automation Backend API

This Flask API runs on the Zabbix server and handles all PowerPoint processing tasks.

## Setup on Zabbix Server

1. Install Python dependencies:
```bash
cd backend_api
pip install -r requirements.txt
```

2. Set environment variables:
```bash
export GRAFANA_URL="http://zabbix.vlan24.co.kr:3000"
export GRAFANA_API_KEY="your_api_key_here"
export GRAFANA_VERIFY_SSL="false"  # Only for testing with self-signed certs
```

3. Run the server:
```bash
python app.py
```

The API will start on port 5001 by default.

## Directory Structure

The backend expects the following directory structure:
```
backend_api/
├── Report/
│   ├── template/          # PowerPoint templates
│   ├── [customer]/        # Customer image folders
│   ├── completed_with_images/  # Step 1 output
│   └── completed_final/   # Final output
├── app.py
├── config.py
└── requirements.txt
```

## API Endpoints

- `GET /health` - Health check and configuration status
- `GET /api/config` - Get current configuration
- `GET /api/customers` - List available customers

## Upcoming Endpoints

- `POST /api/process/images` - Process image insertion
- `POST /api/process/statistics` - Process Grafana statistics insertion
- `POST /api/process/all` - Run full pipeline for all customers
- `POST /api/template/generate` - Auto-generate template from Grafana panels
- `POST /api/support/add` - Add support history to reports
- `POST /api/recommendation/auto` - Auto-generate recommendations
- `POST /api/upload/template` - Upload template files
- `POST /api/upload/images` - Upload image files
- `GET /api/download/<path>` - Download generated reports
