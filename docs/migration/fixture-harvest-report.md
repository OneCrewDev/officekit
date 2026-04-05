# Fixture harvest report

## Included harvest
- Included files: 51
- Included bytes: 864514
- Included kinds: ci-workflow=1, example-fixture=46, readme=1, skill=1, preview-asset=2

## CI smoke flow harvested
- `officecli create test_smoke.docx`
- `officecli add test_smoke.docx /body --type paragraph --prop text="Hello from CI"`
- `officecli get test_smoke.docx /body/p[1]`

## Notable exclusions
- Large generated outputs and model assets were intentionally not copied into the curated fixture set to keep lane-1 commits reviewable.
- MCP-specific assets are tracked as excluded by design in the capability matrix/ledger.

## Largest skipped example assets
- `examples/ppt/models/sun.glb` (4343656 bytes) — excluded-large-binary-or-generated-output
- `examples/ppt/outputs/3d-sun.pptx` (34368562 bytes) — excluded-large-binary-or-generated-output
