import asyncio
from typing import Any

class JobManager:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(JobManager, cls).__new__(cls)
            cls._instance._init_state()
        return cls._instance
        
    def _init_state(self):
        self._job_lock = asyncio.Lock()
        self._validation_jobs: dict[str, dict[str, Any]] = {}
        # Concurrency flags for human-in-the-loop workflows
        self._generate_active = False
        self._finalize_active = False
        
    async def set_job(self, job_id: str, updates: dict[str, Any]) -> None:
        async with self._job_lock:
            current = self._validation_jobs.get(job_id, {})
            current.update(updates)
            self._validation_jobs[job_id] = current
            
    def get_job(self, job_id: str) -> dict[str, Any] | None:
        return self._validation_jobs.get(job_id)

    def try_start_generate(self) -> bool:
        """Intenta iniciar un flujo generate. Retorna True si tiene exito."""
        if self._generate_active or self._finalize_active:
            return False
        self._generate_active = True
        return True
        
    def finish_generate(self) -> None:
        self._generate_active = False

    def try_start_finalize(self) -> bool:
        """Intenta iniciar un flujo finalize. Retorna True si tiene exito."""
        if self._generate_active or self._finalize_active:
            return False
        self._finalize_active = True
        return True
        
    def finish_finalize(self) -> None:
        self._finalize_active = False
