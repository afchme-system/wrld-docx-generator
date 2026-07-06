import os
import io
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

logger = logging.getLogger(__name__)

class DriveService:
    """Google Drive API wrapper for template and document operations"""
    
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
    def __init__(self):
        """Initialize Google Drive service using service account credentials"""
        try:
            # Check for service account JSON in environment or files
            # Accepts either GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_CREDENTIALS_JSON
            # (both names are in use across our services, so we check both)
            sa_json = (
                os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
                or os.getenv('GOOGLE_CREDENTIALS_JSON')
            )
            sa_file = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', '/etc/secrets/service-account.json')
            
            if sa_json:
                # JSON credentials in env variable
                import json
                sa_dict = json.loads(sa_json)
                credentials = service_account.Credentials.from_service_account_info(
                    sa_dict, scopes=self.SCOPES
                )
            elif os.path.exists(sa_file):
                # Credentials file on disk
                credentials = service_account.Credentials.from_service_account_file(
                    sa_file, scopes=self.SCOPES
                )
            else:
                raise ValueError("No Google service account credentials found. Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_CREDENTIALS_JSON")
            
            self.service = build('drive', 'v3', credentials=credentials)
            logger.info("Google Drive service initialized")
        except Exception as e:
            logger.error(f"Failed to initialize Drive service: {e}")
            raise
    
    def download_file(self, file_id):
        """Download a file from Google Drive by file_id, return bytes"""
        try:
            request = self.service.files().get_media(fileId=file_id)
            file_bytes = io.BytesIO()
            downloader = MediaIoBaseDownload(file_bytes, request)
            
            done = False
            while not done:
                _, done = downloader.next_chunk()
            
            file_bytes.seek(0)
            logger.info(f"Downloaded file: {file_id}")
            return file_bytes.getvalue()
        except Exception as e:
            logger.error(f"Failed to download file {file_id}: {e}")
            raise
    
    def upload_file(self, file_bytes, filename, folder_id, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'):
        """Upload a file to Google Drive, return file_id and link"""
        try:
            file_metadata = {
                'name': filename,
                'parents': [folder_id] if folder_id else []
            }
            
            media = MediaFileUpload(
                io.BytesIO(file_bytes),
                mimetype=mimetype,
                resumable=True
            )
            
            # Note: MediaFileUpload doesn't accept BytesIO directly for non-resumable uploads
            # We need to save to temp file first
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False) as tmp:
                tmp.write(file_bytes)
                tmp_path = tmp.name
            
            try:
                media = MediaFileUpload(tmp_path, mimetype=mimetype, resumable=False)
                file = self.service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id, webViewLink'
                ).execute()
                
                file_id = file.get('id')
                drive_link = file.get('webViewLink')
                logger.info(f"Uploaded file: {filename} ({file_id})")
                return file_id, drive_link
            finally:
                os.unlink(tmp_path)
        except Exception as e:
            logger.error(f"Failed to upload file {filename}: {e}")
            raise
    
    def get_file_metadata(self, file_id):
        """Get file metadata (name, size, etc.)"""
        try:
            file = self.service.files().get(
                fileId=file_id,
                fields='id, name, mimeType, size'
            ).execute()
            logger.info(f"Retrieved metadata for {file_id}")
            return file
        except Exception as e:
            logger.error(f"Failed to get metadata for {file_id}: {e}")
            raise
    
    def folder_exists(self, folder_id):
        """Check if a folder exists and is accessible"""
        try:
            result = self.service.files().get(
                fileId=folder_id,
                fields='id, mimeType'
            ).execute()
            is_folder = result.get('mimeType') == 'application/vnd.google-apps.folder'
            logger.info(f"Folder check: {folder_id} exists={bool(result)}, is_folder={is_folder}")
            return is_folder
        except Exception as e:
            logger.warning(f"Folder {folder_id} not accessible: {e}")
            return False
