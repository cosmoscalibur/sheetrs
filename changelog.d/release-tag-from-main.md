### Changed

- The release workflow creates the `vX.Y.Z` tag itself: when the workspace version reaches `main` and has no tag, the job `plan` tags the pushed commit and the release proceeds as it did for a tag pushed by hand. The dry run (manual run of the workflow) is optional (`CONTRIBUTING.md`, *Releases*).
