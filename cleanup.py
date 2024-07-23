import os
import shutil
import subprocess

def delete_files_in_dir(directory):
    print(f"Cleaning {directory}")
    try:
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            try:
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    os.unlink(item_path)
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)
            except Exception as e:
                print(f'Failed to delete {item_path}. Reason: {e}')
    except Exception as e:
        print(f'Failed to list items in {directory}. Reason: {e}')

def clean_temp_files():
    temp_dirs = [
        os.getenv('TEMP'),
        os.getenv('TMP'),
        os.path.expandvars(r'%SYSTEMROOT%\Temp'),
        os.path.join(os.getenv('USERPROFILE'), 'AppData', 'Local', 'Temp')
    ]

    for temp_dir in temp_dirs:
        if temp_dir:
            delete_files_in_dir(temp_dir)

def run_disk_cleanup():
    try:
        # Create a new Disk Cleanup task with cleanmgr /sageset
        subprocess.run(['cleanmgr', '/sageset:1'], check=True)
        # Run the Disk Cleanup task created above
        subprocess.run(['cleanmgr', '/sagerun:1'], check=True)
        print("Disk Cleanup executed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to execute Disk Cleanup. Reason: {e}")

if __name__ == "__main__":
    clean_temp_files()
    run_disk_cleanup()