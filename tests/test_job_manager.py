import asyncio
import pytest

try:
    from app.application.job_manager import JobManager
except ImportError:
    JobManager = None

def test_job_manager_singleton():
    if JobManager is None:
        pytest.fail("JobManager not implemented")
    
    manager1 = JobManager()
    manager2 = JobManager()
    assert manager1 is manager2

def test_job_manager_concurrency_lock():
    if JobManager is None:
        pytest.fail("JobManager not implemented")
        
    async def run_test():
        manager = JobManager()
        manager._validation_jobs = {}
        manager._generate_active = False
        manager._finalize_active = False
        
        success1 = manager.try_start_generate()
        assert success1 is True
        
        success2 = manager.try_start_generate()
        assert success2 is False
        
        manager.finish_generate()
        
        success3 = manager.try_start_generate()
        assert success3 is True
        manager.finish_generate()
        
    asyncio.run(run_test())

def test_job_manager_status_updates():
    if JobManager is None:
        pytest.fail("JobManager not implemented")
        
    async def run_test():
        manager = JobManager()
        job_id = "test_job_123"
        
        await manager.set_job(job_id, {"status": "running"})
        job = manager.get_job(job_id)
        assert job is not None
        assert job["status"] == "running"
        
    asyncio.run(run_test())
