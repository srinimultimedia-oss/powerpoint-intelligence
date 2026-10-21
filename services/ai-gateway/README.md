# On-prem AI gateway (not implemented)

Azure OpenAI is Azure-hosted; gateway remains on-premises. No outbound requests until endpoint, model deployment, identity, networking and data handling are approved.

Planned pipeline: request -> RBAC -> authorized context retrieval -> data minimization -> Azure OpenAI structured action plan -> schema validation -> deterministic calculations -> human approval -> host command execution -> audit.

Never execute arbitrary code/SQL/COM calls from model-generated text. Store no secrets in source control. Implement contract tests before integration.
