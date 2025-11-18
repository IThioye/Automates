import os
from packaging.version import Version
from dotenv import load_dotenv

# --- Load .env file ---
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(os.path.dirname(PROJECT_DIR),"config.env")
load_dotenv(dotenv_path=ENV_PATH)

current_version_str = os.getenv("VERSION")

if not current_version_str:
    raise ValueError("VERSION variable not found in .env")

current_version = Version(current_version_str)

# --- Read past_versions.txt ---
past_versions_path = f"{PROJECT_DIR}\past_versions.txt"
lines = []
if not os.path.exists(past_versions_path):
    with open(past_versions_path, "w") as f:
        f.write("0.0.0\n")
        lines = ["0.0.0"]
else:
    with open(past_versions_path, "r") as f:
        lines = [line.strip() for line in f if line.strip()]

# Parse and sort versions
valid_versions = sorted((Version(v) for v in lines), reverse=True)
latest_version = valid_versions[0]
print(f"Latest version from past_versions.txt: {latest_version}")

# --- Compare and update if needed ---
if latest_version >= current_version:
    # Example logic: bump patch version
    new_version = Version(f"{latest_version.major}.{latest_version.minor}.{latest_version.micro + 1}")
    print(f"Updating .env VERSION from {current_version} to {new_version}")

    # Read .env and update the VERSION line
    with open(ENV_PATH, "r") as f:
        env_lines = f.readlines()

    with open(ENV_PATH, "w") as f:
        for line in env_lines:
            if line.startswith("VERSION="):
                f.write(f"VERSION={new_version}\n")
            else:
                f.write(line)
    print(f"Update done: .env VERSION is now {new_version}")
else:
    print(f"No update needed. Current VERSION in config.env ({current_version}) is up to date.")
