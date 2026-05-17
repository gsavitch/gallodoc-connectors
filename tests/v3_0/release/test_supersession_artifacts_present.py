import os

def test_supersession_artifacts():
    assert os.path.exists("docs/GALLODOC_CORE_V1_FROZEN.md")
    with open("docs/GALLODOC_CORE_V1_FROZEN.md") as f:
        content = f.read()
        assert "SUPERSEDED BY V3" in content

    assert os.path.exists("opensource/gallodoc-core/pyproject.toml")
    with open("opensource/gallodoc-core/pyproject.toml") as f:
        content = f.read()
        assert "Development Status :: 4 - Beta" in content

if __name__ == "__main__":
    test_supersession_artifacts()
    print("test_supersession_artifacts: PASSED")
