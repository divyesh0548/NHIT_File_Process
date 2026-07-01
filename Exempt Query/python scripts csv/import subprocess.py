import subprocess
import time

def run_script(script_name):
    print(f"\nRunning {script_name}...")
    start_time = time.time()        

    result = subprocess.run(["python", script_name], capture_output=True, text=True)

    end_time = time.time()
    elapsed = end_time - start_time
    print(f"Time taken for {script_name}: {elapsed:.2f} seconds")

    if result.returncode == 0:
        print(f"{script_name} completed successfully.")
    else:
        print(f"Error running {script_name}:\n{result.stderr}")

    return elapsed

def main():
    scripts = [
        "config.py",
        "data_processing.py",
        "main_script.py",
        "RF3_condition.py",
        "combined_with_rf3.py",
        "Date_validity_check.py",
        "Date_validity_check_LT.py", # Comment out this if no LT pass is available.
        "Return_Journey_Logic.py"
    ]

    total_start = time.time()
    total_time = 0

    for script in scripts:
        script_time = run_script(script)
        total_time += script_time

    total_end = time.time()
    print(f"\nTotal time for all scripts: {total_time:.2f} seconds (measured individually)")
    print(f"Actual wall-clock time: {total_end - total_start:.2f} seconds")

if __name__ == "__main__":
    main()
