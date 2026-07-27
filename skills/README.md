# The agent skill moved into the package

`skills/docx-plus/` is now **`docx_plus/skill/`**, so it ships in the wheel
instead of living only in this repository. Anyone who `pip install`s
`docx-plus` gets it on disk.

Install it into an agent's skills directory with the CLI:

```bash
docx-plus skill install            # -> ./.claude/skills/docx-plus
docx-plus skill install --user     # -> ~/.claude/skills/docx-plus
docx-plus skill path               # where the packaged copy lives
docx-plus skill show tables        # read one topic without installing
```

See [`docs/SKILLS.md`](../docs/SKILLS.md) for the full guide.
