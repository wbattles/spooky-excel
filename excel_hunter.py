from pathlib import Path
import sys
import threading
import sys
import time
from itertools import cycle

def loading_animation():
    spinner = cycle(['🌑', '🌒', '🌓', '🌔', '🌕', '🌖', '🌗', '🌘'])
    while not loading_animation.done:
        sys.stdout.write('\r🔮 Hunting Excel files... ' + next(spinner))
        sys.stdout.flush()
        time.sleep(0.1)
    
    sys.stdout.write('\r')
    sys.stdout.flush()

def trick_or_treat(dir_path, visited, excel_files):
    try:
        current_dir = Path(dir_path)
        
        for item in current_dir.iterdir():
            try:
                if item.is_file() and item.suffix.lower() in ['.xlsx', '.xls', '.xlsm']:
                    excel_files.add(str(item.absolute()))
                
                elif item.is_dir() and str(item.absolute()) not in visited:
                    visited.add(str(item.absolute()))
                    trick_or_treat(item, visited, excel_files)
                    
            except PermissionError:
                continue
            except Exception as e:
                print(f"Error processing {item}: {e}", file=sys.stderr)
                continue
                
    except PermissionError:
        return
    except Exception as e:
        print(f"Error accessing directory {dir_path}: {e}", file=sys.stderr)
        return

def the_haunting():
    visited = set()
    excel_files = set()
    
    loading_animation.done = False
    loading_thread = threading.Thread(target=loading_animation)
    loading_thread.start()
    
    try:
        user_dir = str(Path.home())
        trick_or_treat(user_dir, visited, excel_files)
    finally:
        loading_animation.done = True
        loading_thread.join()
    
    return len(excel_files)
if __name__ == '__main__':
    count = the_haunting()
    
    print(f"\n👻 Found {count} spooky Excel files haunting your computer")
    print("🎃 Happy Halloween! 🎃")