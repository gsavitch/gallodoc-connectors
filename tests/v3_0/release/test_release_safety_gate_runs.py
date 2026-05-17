import subprocess
import os

def test_gate_runs():
    # Path to script
    script_path = os.path.abspath("scripts/release_safety_gate.py")
    result = subprocess.run(["python3", script_path], capture_output=True, text=True)
    assert result.returncode == 0
    assert "GATE PASSED" in result.stdout
    assert os.path.exists("release_safety_report.json")

if __name__ == "__main__":
    test_gate_runs()
    print("test_gate_runs: PASSED")
