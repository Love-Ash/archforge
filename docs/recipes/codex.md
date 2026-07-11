# Recipe: archforge in a Codex / agent pipeline

Agents cannot see their own renders; give them the gate as a tool.

```bash
pip install archforge
```

System-prompt contract (works for any agent framework, including MCP-wrapped shells):

```
After building or editing a .pptx, run:
  archforge <file> --profile full --fail-incomplete --json
Parse the JSON. Ship only when summary.pass is true. For each finding, use
location.shape_id / location.bbox / location.paragraph / location.run to target the
fix. E1/E2/E4 may be fixed deterministically first:
  archforge fix <file> -o <file>
Re-lint after every fix round.
```

`llms.txt` at the repo root is a condensed version of this contract for context
injection, and `archforge skill` prints the full Agent Skills pack.
