# PowerPoint Automation Tool

Automate PowerPoint report generation with image insertion and Grafana statistics integration.

## Features

- **Image Insertion**: Automatically insert images into PowerPoint templates based on shape names
- **Grafana Integration**: Fetch metrics from Grafana dashboards and insert statistics
- **Date Automation**: Automatically populate date placeholders for monthly reports
- **Interactive CLI**: User-friendly command-line interface
- **Batch Processing**: Process multiple files and customers simultaneously

## Quick Start

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. **Run the Web GUI** (Recommended):
```bash
streamlit run app.py
```

   Or use the interactive CLI:
```bash
python main.py
```

3. Place your PowerPoint templates in `Report/template/`
4. Place images in customer-specific folders under `Report/`
5. Configure Grafana settings in `config.py` or use environment variables

The web GUI provides an easy-to-use interface with:
- Real-time status monitoring
- One-click execution
- Detailed logging
- Configuration management

## Documentation

For detailed documentation, see [replit.md](replit.md)

## Workflow

1. **Step 1**: Insert images into templates → output to `Report/completed_with_images/`
2. **Step 2**: Insert Grafana statistics and dates → output to `Report/completed_final/`

## Configuration

Set these environment variables for Grafana integration:
- `GRAFANA_URL` - Your Grafana server URL
- `GRAFANA_API_KEY` - Your Grafana API key
- `GRAFANA_VERIFY_SSL` - SSL verification (default: true, disable only for testing)

Edit `config.py` to configure:
- Directory paths
- Customer dashboard mappings
- Output format templates

## Requirements

- Python 3.11+
- python-pptx
- python-dateutil
- requests
- urllib3

## License

Open source - feel free to use and modify for your needs.
