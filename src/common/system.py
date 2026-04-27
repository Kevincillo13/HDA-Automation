import subprocess
import logging

logger = logging.getLogger(__name__)

def kill_processes(process_names: list[str]) -> None:
    """
    Kills a list of Windows processes by name.
    Does not raise error if process is not found.
    """
    for name in process_names:
        try:
            # /F = force, /IM = image name, /T = terminate child processes
            subprocess.run(
                ["taskkill", "/F", "/IM", name, "/T"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False
            )
            logger.info("Cleanup: tried to kill process %s", name)
        except Exception as e:
            logger.warning("Could not kill process %s: %s", name, e)
