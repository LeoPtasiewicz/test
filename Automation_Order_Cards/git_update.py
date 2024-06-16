import subprocess
import sys

def run_git_command(command):
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error running command: {command}")
        print(result.stderr)
    else:
        print(result.stdout)

def main():
    if len(sys.argv) < 2:
        print("Usage: python git_update.py <commit_message>")
        sys.exit(1)
    
    commit_message = sys.argv[1]

    # Stage changes
    run_git_command("git add .")

    # Commit changes
    run_git_command(f'git commit -m "{commit_message}"')

    # Push changes
    run_git_command("git push origin main")

if __name__ == "__main__":
    main()
