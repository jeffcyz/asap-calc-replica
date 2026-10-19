# Local-only ASAP knowledge intake

The **real ASAP knowledge base belongs under \`knowledge/asap/\` on your local machine** and that folder is Git-ignored. Do not upload employer manuals, internal API documentation, deal scripts, database extracts or screenshots to this public repository unless expressly authorized.

Suggested local-only arrangement:

~~~text
knowledge/asap/
  objects/
  methods/
  collection_semantics/
  execution_lifecycle/
  data_model/
  index_and_dates/
  typing_and_rounding/
~~~

You may link a separately authorized, locally available knowledge base (such as your ASAP-Agent project) rather than duplicating it. Only use paths/connectors permitted by your organization's policies; do not assume a private GitHub repository or ChatGPT connection makes the data safe for all tools.

Every runtime contract should cite a concrete source version and location, plus record unknown side effects and missing input data separately.

See \`docs/knowledge_base_integration.md\`.
