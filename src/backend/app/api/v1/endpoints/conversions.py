# Excel to PowerPoint Conversion Endpoints
# Handles the core conversion functionality
# - POST /conversions/start - Start Excel to PPT conversion job
# - GET /conversions/{job_id}/status - Check conversion job status
# - GET /conversions/{job_id}/result - Download converted PowerPoint file
# - GET /conversions - List user's conversion history
# - DELETE /conversions/{job_id} - Cancel or delete conversion job
# - POST /conversions/preview - Preview conversion without saving
# - GET /conversions/templates - Get available PowerPoint templates