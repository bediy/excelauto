# Repository Guidelines

## Project Structure & Module Organization
- `create_person_sheets.py` copies the `Ф    ` template in `      .xlsx`, populating each sheet with people data from `2025   У      Ա  Ϣ    .xlsx`.
- `insert_images.py` embeds up to two photos per person into the target workbook; images are sourced from `images/` and are matched by filename prefix.
- `remove_extra_sheets.py` keeps only the first worksheet in a workbook, useful for resetting templates before regeneration.
- The `images/` directory stores portrait assets; keep filenames in `<Name><index>.jpg` format so ordering logic remains stable.

## Build, Test, and Development Commands
- `python create_person_sheets.py` generates or refreshes per-person sheets using the configured source and target workbooks.
- `python insert_images.py       .xlsx images/` attaches the first two images for each matching person sheet.
- `python remove_extra_sheets.py       .xlsx` strips surplus sheets before re-running other scripts.
- Run scripts inside a virtual environment with `python -m venv .venv` and `.\.venv\Scripts\activate` to isolate dependencies (`openpyxl`, `Pillow`).

## Coding Style & Naming Conventions
- Follow PEP 8: four-space indentation, snake_case for functions, UpperCamelCase only for classes, and UPPER_SNAKE_CASE for constants (see `TARGET_CELL`).
- Prefer `pathlib.Path` for filesystem paths and add type hints for public functions, as done in existing modules.
- Use descriptive docstrings and inline comments in Chinese when clarifying domain-specific Excel logic.
- Keep IO boundaries narrow: centralize file paths at the top of each script so deployments can adjust locations easily.

## Testing Guidelines
- Before modifying scripts, clone the production workbook and run scripts against the copy: `python create_person_sheets.py` → verify sheet count, then `python insert_images.py ...` → confirm images anchor correctly.
- When adding logic, extract helper functions and cover them with `pytest` or `unittest` cases; place tests under `tests/` (create if absent).
- Verify regenerated workbooks with Excel’s “Inspect Document” to ensure no hidden warnings and spot-check a few random sheets for data fidelity.

## Commit & Pull Request Guidelines
- Use imperative, scoped commit messages such as `feat: automate image anchoring` or `fix: handle missing id card values`.
- Reference related workbooks or issue IDs in the body and describe manual verification performed (e.g., “Ran insert_images.py on sample workbook”).
- Pull requests should summarize the workflow change, note any dependency updates, attach before/after screenshots when UI-facing, and request a peer review before merging.

## Excel & Asset Handling
- Back up `      .xlsx` before running automation; scripts overwrite sheets in-place.
- Keep portrait images at ~500 KB for faster embedding and consistent resizing; prefer JPG for photos, PNG for signatures.
- When onboarding new data files, validate column placements (name, soldier ID, ID card) to avoid silent skips during processing.

##其他要求
-编写代码时尽可能的加上中文注释
-注意清理废弃的代码
