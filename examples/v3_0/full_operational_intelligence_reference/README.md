# Full Operational Intelligence Reference Demo (v3)

This demo walks through the end-to-end lifecycle of a GalloDoc v3 document.

## Steps
1. **Ingestion**: `vendor_invoice.gdoc.json` is created from Word.
2. **Context**: `employee_record.gdoc.json` provides HR context.
3. **Linking**: `linker_output.json` suggests a relationship between the invoice and the employee.
4. **Semantic Augmentation**: `linker_output_with_embeddings.json` shows how AI improves confidence.
5. **Human Governance**: `human_review_decision.json` confirms the link.
6. **Query**: `aibi_query_receipt.json` shows the AI/BI planner resolving the question: "Who approved the Acme Corp invoice?"
