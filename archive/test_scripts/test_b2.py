# test_b2.py
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def format_bytes(bytes):
    """Convert bytes to human readable format"""
    if bytes == 0:
        return "0 B"
    
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if bytes < 1024.0:
            return f"{bytes:.2f} {unit}"
        bytes /= 1024.0
    return f"{bytes:.2f} TB"

def get_storage_usage(bucket):
    """Calculate total storage usage in the bucket"""
    total_size = 0
    file_count = 0
    
    print("\n📊 Analyzing storage usage...")
    
    try:
        # List all files in bucket
        for file_version in bucket.ls(recursive=True):
            total_size += file_version.size
            file_count += 1
            
        return total_size, file_count
    except Exception as e:
        print(f"❌ Error calculating storage: {e}")
        return 0, 0

def get_account_info(b2_api):
    """Get account storage limits and usage"""
    try:
        # Get account info
        account_info = b2_api.account_info
        
        # Note: B2 doesn't provide storage limits via API for free accounts
        # We'll use the known free tier limits
        FREE_TIER_BYTES = 10 * 1024 * 1024 * 1024  # 10 GB
        
        return FREE_TIER_BYTES
    except Exception as e:
        print(f"⚠️ Could not get account limits: {e}")
        return 10 * 1024 * 1024 * 1024  # Default to 10GB

try:
    from b2sdk.v2 import InMemoryAccountInfo, B2Api
    
    # Get credentials from environment
    key_id = os.getenv('B2_APPLICATION_KEY_ID')
    key = os.getenv('B2_APPLICATION_KEY')
    bucket_name = os.getenv('B2_BUCKET_NAME', 'findeck-files')
    
    print(f"🔥 Testing Backblaze B2 Connection")
    print(f"=" * 50)
    print(f"Key ID: {key_id}")
    print(f"Bucket: {bucket_name}")
    
    # Initialize B2
    info = InMemoryAccountInfo()
    b2_api = B2Api(info)
    b2_api.authorize_account("production", key_id, key)
    
    print("✅ B2 Authorization successful!")
    
    # Test bucket access
    try:
        bucket = b2_api.get_bucket_by_name(bucket_name)
        print(f"✅ Bucket '{bucket_name}' found!")
        
        # Get storage usage
        used_bytes, file_count = get_storage_usage(bucket)
        limit_bytes = get_account_info(b2_api)
        remaining_bytes = limit_bytes - used_bytes
        
        # Calculate percentages
        usage_percentage = (used_bytes / limit_bytes * 100) if limit_bytes > 0 else 0
        
        print(f"\n💾 STORAGE ANALYSIS")
        print(f"=" * 50)
        print(f"📁 Total Files: {file_count}")
        print(f"📊 Storage Used: {format_bytes(used_bytes)}")
        print(f"🎯 Storage Limit: {format_bytes(limit_bytes)} (Free Tier)")
        print(f"📈 Usage: {usage_percentage:.1f}%")
        print(f"🆓 Remaining: {format_bytes(remaining_bytes)}")
        
        # Storage status
        if usage_percentage > 90:
            print(f"🚨 WARNING: Storage almost full!")
        elif usage_percentage > 70:
            print(f"⚠️  CAUTION: Storage 70%+ used")
        elif usage_percentage > 50:
            print(f"💡 INFO: Storage over halfway used")
        else:
            print(f"✅ GOOD: Plenty of storage available")
            
        # File breakdown by type
        if file_count > 0:
            print(f"\n📋 FILE BREAKDOWN")
            print(f"=" * 30)
            
            file_types = {
                'uploads': {'count': 0, 'size': 0},
                'analysis': {'count': 0, 'size': 0}, 
                'outputs': {'count': 0, 'size': 0},
                'other': {'count': 0, 'size': 0}
            }
            
            for file_version in bucket.ls(recursive=True):
                path = file_version.file_name
                size = file_version.size
                
                if path.startswith('uploads/'):
                    file_types['uploads']['count'] += 1
                    file_types['uploads']['size'] += size
                elif path.startswith('analysis/'):
                    file_types['analysis']['count'] += 1
                    file_types['analysis']['size'] += size
                elif path.startswith('outputs/'):
                    file_types['outputs']['count'] += 1
                    file_types['outputs']['size'] += size
                else:
                    file_types['other']['count'] += 1
                    file_types['other']['size'] += size
            
            for category, data in file_types.items():
                if data['count'] > 0:
                    print(f"{category.capitalize():10}: {data['count']:3} files | {format_bytes(data['size']):>10}")
        
        # Recommendations
        print(f"\n💡 RECOMMENDATIONS")
        print(f"=" * 30)
        if usage_percentage > 80:
            print("🔧 Consider implementing file cleanup job")
            print("🗑️  Delete files older than 24 hours")
            print("📦 Compress analysis data before storage")
        elif usage_percentage > 50:
            print("📊 Monitor usage regularly")
            print("🔄 Implement automatic cleanup soon")
        else:
            print("✅ Storage usage is healthy")
            print("📈 Monitor as user base grows")
            
    except Exception as e:
        print(f"❌ Bucket error: {e}")
        print("💡 Create the bucket in B2 console first")
    
except ImportError:
    print("❌ B2SDK not installed. Run: pip install b2sdk")
except Exception as e:
    print(f"❌ B2 connection failed: {e}")