# SheetGuard development workflow

- Use this repository for SheetGuard AI development. Start changes on a development branch based on `main`.
- Run the relevant automated tests before requesting review. Record the results and distinguish Excel test doubles from live Excel verification.
- Open a pull request targeting `main` for review before merging. Validate changes in staging where applicable. Deploy to production only after review and approval.
- Preserve the working production manifest and Render service during development. Keep credentials in Render secret settings.
