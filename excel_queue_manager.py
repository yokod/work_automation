import os
import json
import fcntl
import time
import datetime

QUEUE_FILE = r"d:\yoel\projects\auto\logs\excel_tasks_queue.json"

def _lock_and_run(func):
    """File locking wrapper to safely read/write the JSON queue across multiple scripts."""
    # Since this is Windows, we use a simple lock file mechanism
    LOCK_FILE = QUEUE_FILE + ".lock"
    
    def wrapper(*args, **kwargs):
        for _ in range(50): # Wait up to 5 seconds
            try:
                # Try to create lock file exclusively
                fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, b"locked")
                os.close(fd)
                
                try:
                    return func(*args, **kwargs)
                finally:
                    try:
                        os.remove(LOCK_FILE)
                    except:
                        pass
                break
            except FileExistsError:
                time.sleep(0.1)
        else:
            raise Exception("Timeout waiting for queue lock")
    return wrapper

@_lock_and_run
def push_task(task_type, case_num, date_obj, payload):
    """
    Pushes a task to the queue.
    task_type: "UPDATE_WORD_COUNT", "UPDATE_STATUS", "MARK_SENT", etc.
    payload: dict of additional data.
    """
    queue = []
    if os.path.exists(QUEUE_FILE):
        try:
            with open(QUEUE_FILE, "r", encoding="utf-8") as f:
                queue = json.load(f)
        except:
            pass
            
    # Serialize date
    d_str = str(date_obj) if date_obj else None
    
    task = {
        "timestamp": datetime.datetime.now().isoformat(),
        "task_type": task_type,
        "case_num": case_num,
        "date_str": d_str,
        "payload": payload
    }
    queue.append(task)
    
    with open(QUEUE_FILE, "w", encoding="utf-8") as f:
        json.dump(queue, f, ensure_ascii=False, indent=2)
        
    return True

@_lock_and_run
def pop_all_tasks():
    """Returns all tasks and clears the queue."""
    if not os.path.exists(QUEUE_FILE):
        return []
        
    try:
        with open(QUEUE_FILE, "r", encoding="utf-8") as f:
            queue = json.load(f)
            
        # Clear queue
        with open(QUEUE_FILE, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)
            
        return queue
    except Exception as e:
        print(f"[Queue Manager] Error reading queue: {e}")
        return []
