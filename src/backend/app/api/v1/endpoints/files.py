# File Management Endpoints
# Handles Excel file uploads and management
# - POST /files/upload - Upload Excel files with validation
# - GET /files - List user's uploaded files with pagination
# - GET /files/{file_id} - Get specific file metadata
# - DELETE /files/{file_id} - Delete uploaded file
# - GET /files/{file_id}/download - Download original Excel file
# - GET /files/{file_id}/preview - Preview file content and structure
# - PUT /files/{file_id}/metadata - Update file metadata (name, description)