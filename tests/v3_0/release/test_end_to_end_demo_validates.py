import os
import json

def test_demo_files_valid():
    demo_dir = "examples/v3_0/full_operational_intelligence_reference/"
    files = [
        "vendor_invoice.gdoc.json",
        "employee_record.gdoc.json",
        "linker_output.json",
        "linker_output_with_embeddings.json",
        "human_review_decision.json",
        "aibi_query_receipt.json"
    ]
    for f in files:
        path = os.path.join(demo_dir, f)
        assert os.path.exists(path)
        with open(path) as jf:
            data = json.load(jf)
            assert data is not None

if __name__ == "__main__":
    test_demo_files_valid()
    print("test_demo_files_valid: PASSED")
