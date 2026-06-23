# WRLD Document Generator Microservice

Flask microservice that fills Word templates (.docx) with content from JSON and returns generated documents.

## Features

- Download templates from Google Drive
- Replace placeholders `[like this]` with JSON values
- Return binary .docx files
- List available templates
- Full error handling and logging
- Deploy to Railway in 5 minutes

## Local Setup

### Prerequisites
- Python 3.9+
- pip

### Installation

1. Clone repo and navigate to directory:
```bash
cd wrld-docx-generator
```

2. Create virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Set up environment variables:
```bash
cp .env.example .env
# Edit .env with your credentials
```

### Running Locally

```bash
python app.py
```

Server runs on `http://localhost:5000`

Test endpoint:
```bash
curl http://localhost:5000
```

## API Endpoints

### POST /fill-template

Fill a template with content and return .docx file.

**Request:**
```json
{
  "template_name": "WRLD-Template-Letter (1).docx",
  "content": {
    "date": "22 June 2026",
    "recipient_full_name": "John Smith",
    "recipient_title": "Hotel Director",
    "organisation": "Luxury Hotels XYZ",
    "address_line_1": "123 Oxford Street",
    "city_postcode": "London, W1D 2AF",
    "country": "United Kingdom",
    "subject_reference": "Uniform Supply Proposal",
    "recipient_name": "John",
    "opening_paragraph": "We are pleased to present...",
    "body_paragraph_1": "Our team specializes in...",
    "body_paragraph_2": "We propose the following terms...",
    "closing_paragraph": "We welcome your response...",
    "signatory_full_name": "Sebastian Dawber",
    "signatory_title": "Director",
    "company_name": "AFC",
    "email": "sebastian@afc.com",
    "phone": "+44 20 7946 0958"
  },
  "output_filename": "Letter_AFC_2026.docx"
}
```

**Response:** Binary .docx file (application/vnd.openxmlformats-officedocument.wordprocessingml.document)

**Example with curl:**
```bash
curl -X POST http://localhost:5000/fill-template \
  -H "Content-Type: application/json" \
  -d @request.json \
  --output output.docx
```

### GET /list-templates

List available templates in Drive folder.

**Response:**
```json
{
  "status": "ok",
  "count": 5,
  "templates": [
    "WRLD-Template-Letter (1).docx",
    "WRLD-Template-Brief (1).docx",
    "WRLD-Template-Proposal.docx",
    "WRLD-Template-Research.docx",
    "WRLD-Template-Concept.docx"
  ]
}
```

### GET /health

Health check.

**Response:**
```json
{
  "status": "ok",
  "service": "WRLD Document Generator"
}
```

## Google Drive Setup

The service uses a Google Service Account to access Drive. You need:

1. A Google Cloud project
2. A service account with Drive API access
3. Share the templates folder with the service account email

### Getting Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or use existing)
3. Enable Google Drive API
4. Create a Service Account:
   - Go to IAM & Admin → Service Accounts
   - Click "Create Service Account"
   - Fill in name and description
   - Grant Editor role (or Drive Editor)
5. Create a key:
   - Click the service account
   - Go to "Keys" tab
   - Add JSON key
   - Download the JSON file
6. Share the Drive folder with the service account:
   - Get the service account email from the JSON file
   - Share the templates folder with that email

### Environment Variables

Set these in Railway:

- `GOOGLE_CREDENTIALS_JSON`: Entire JSON service account key (as string)
- `GOOGLE_DRIVE_TEMPLATES_FOLDER`: Template folder ID (get from Drive URL: `drive.google.com/drive/folders/[FOLDER_ID]`)

## Deployment to Railway

### 1. Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/wrld-docx-generator.git
git push -u origin main
```

### 2. Connect Railway

1. Go to [Railway.app](https://railway.app)
2. Click "New Project"
3. Select "Deploy from GitHub"
4. Authorize and select your repo
5. Click "Deploy"

### 3. Set Environment Variables in Railway

1. Go to project settings
2. Click "Variables"
3. Add:
   - `GOOGLE_CREDENTIALS_JSON`: (paste entire JSON)
   - `GOOGLE_DRIVE_TEMPLATES_FOLDER`: `1R0F8BUkbC8hKJBv3Z8kJp2_jBS3hj_gf`

### 4. Restart Service

The service auto-deploys on push. Check logs in Railway dashboard.

## Template Placeholders

Templates use `[Placeholder]` format in square brackets.

### Letter Template
- `[date]`
- `[recipient_full_name]`
- `[recipient_title]`
- `[organisation]`
- `[address_line_1]`
- `[city_postcode]`
- `[country]`
- `[subject_reference]`
- `[recipient_name]`
- `[opening_paragraph]`
- `[body_paragraph_1]`
- `[body_paragraph_2]`
- `[closing_paragraph]`
- `[signatory_full_name]`
- `[signatory_title]`
- `[company_name]`
- `[email]`
- `[phone]`

### Other Templates
Each template has its own set of placeholders. Use `/list-templates` to see available templates, then check the Drive folder for their structure.

## Error Handling

The service returns meaningful error messages:

```json
{
  "error": "Template 'INVALID.docx' not found in Drive folder",
  "status": "template_not_found"
}
```

Status codes:
- `200`: Success
- `400`: Validation error (bad request)
- `404`: Template not found
- `500`: Server error

## Logging

All operations are logged. View logs in:
- **Local:** Console output
- **Railway:** Railway dashboard → Logs tab

## Development

### Adding New Templates

1. Create .docx file with placeholders in format `[like_this]`
2. Upload to WRLD Templates folder in Drive
3. Use in requests - service discovers automatically

### Testing

```python
# Create test_request.py
import requests
import json

payload = {
    "template_name": "WRLD-Template-Letter (1).docx",
    "content": {
        "date": "22 June 2026",
        "recipient_full_name": "Test User"
    },
    "output_filename": "test.docx"
}

response = requests.post('http://localhost:5000/fill-template', json=payload)
with open('output.docx', 'wb') as f:
    f.write(response.content)
```

## Troubleshooting

### "GOOGLE_CREDENTIALS_JSON not set"
- Check Railway Variables are set
- Credentials must be JSON string (not file)

### "Template not found"
- Check template name matches exactly (case-sensitive)
- Verify folder ID is correct
- Check service account has access to folder

### Large placeholders not replaced
- python-docx has limitations with complex formatting
- If placeholder spans multiple runs, document may need manual fix
- Open in Word and resave if needed

## Production Checklist

- [ ] Google service account created and verified
- [ ] Drive templates folder shared with service account
- [ ] Railway project created
- [ ] Environment variables set in Railway
- [ ] Tested /fill-template endpoint
- [ ] Tested /list-templates endpoint
- [ ] Monitored logs for errors
- [ ] Set up error notifications (optional)

## Support

Contact: Sebastian Dawber (sebastian@afc.com)
