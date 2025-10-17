# Automated File Cleanup Service
# Runs scheduled cleanup jobs to delete files after 24 hours
# - Scheduled background task for automatic cleanup
# - Manual cleanup trigger for admin use
# - Cleanup statistics and reporting
# - Configurable cleanup intervals and retention periods

import asyncio
import logging
from datetime import datetime, timedelta
from typing import Optional,Dict
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.interval import IntervalTrigger

from .b2_storage import get_storage_client
from ..core.config import settings

logger=logging.getLogger(__name__)

class FileCleanupService:
    """Sutomated file cleanup service for Findeck"""

    def __init__(self):
        self.scheduler=AsyncIOScheduler()
        self.storage_client=None
        self.cleanup_stats={
            "last_cleanup":None,
            "total_cleanup":0,
            "total_files_deleted":0,
            "total_mb_freed":0
        }

    async def initialize(self):
        """Initialize the cleanup service"""

        try:
            # Initialize storage client
            self.storage_client=await get_storage_client(
                use_b2=settings.use_b2_storage,
                application_key_id=settings.b2_application_key_id,
                application_key=settings.b2_application_key,
                bucket_name=settings.b2_bucket_name
            )

            # Schedule periodic cleanup every 6 hours
            self.scheduler.add_job(
                func=self.run_cleanup,
                trigger=IntervalTrigger(hours=6),
                id="file_cleanup_job",
                name="24-Hour File Cleanup Job",
                replace_existing=True
            )

            # start the scheduler
            self.scheduler.start()

            logger.info("File cleanup service initialized and scheduler started")

        except Exception as e:
            logger.error(f"Error initializing FileCleanupService: {e}")
            raise e
        
    async def run_cleanup(self,retention_hours:int=24):
        """Run the cleanup job"""
        try:
            logger.info(f"Starting cleanup job with retention {retention_hours} hours")

            # Run cleanup
            deleted_counts=await self.storage_client.cleanup_expiried_files(retention_hours=retention_hours)

            # Update stats
            self.cleanup_stats["last_cleanup"]=datetime.utcnow().isoformat()
            self.cleanup_stats["total_cleanup"]+=1

            total_deleted=sum(deleted_counts[k] for k in['uploads','analysis','outputs','others'])
            self.cleanup_stats["total_files_deleted"]+=total_deleted
            self.cleanup_stats["total_mb_freed"]+=deleted_counts.get("total_mb",0)

            # Log results
            logger.info(f"Cleanup completed: {total_deleted}," f"{deleted_counts['total_mb']:.2f}MB freed")

            return {
                'cleanup_time': self.cleanup_stats["last_cleanup"],
                'files_deleted': total_deleted,
                'success': True
            }
        except Exception as e:
            logger.error(f"Error during cleanup job: {e}")
            return {
                'cleanup_time': datetime.utcnow().isoformat(),
                'files_deleted': 0,
                'success': False,
                'error': str(e)
            }
        

    async def manual_cleanup(self,retention_hours:int=24)->Dict:
        """Manually trigger a cleanup job"""
        return await self.run_cleanup(retention_hours=retention_hours)
    

    def get_cleanup_stats(self)->Dict:
        """Get current cleanup statistics"""
        return{
            **self.cleanup_stats,
            'scheduler_running': self.scheduler.running,
            'next_cleanup': self.scheduler.get_job("file_cleanup_job").next_run_time.isoformat() if self.scheduler.get_job("file_cleanup_job") else None
        }
    

    async def shutdown(self):
        """Shutdown the cleanup service"""
        if self.scheduler.running:
            self.scheduler.shutdown()
            logger.info("File cleanup service scheduler shut down")


## Global cleanup service instance
cleanup_service=FileCleanupService()