### Added

- Add `registry::create_walker_rules`, which builds fresh rule instances for a list of rule IDs and merges INT401–403 into one `InterruptionRule`.
- Add `InterruptionKind::from_rule_id` and `InterruptionKind::rule_id`.

### Changed

- **Breaking:** `WalkerRule` hooks take `&mut self`, and the trait requires `Send` instead of `Send + Sync`. Rules keep their collected state in plain fields instead of `Mutex`-guarded ones. Production benchmark: about 0.3% fewer user-space instructions, same peak RAM.
- **Breaking:** `InterruptionRule::new` takes the list of `InterruptionKind`s to report. One instance covers every enabled INT401–403 rule.

### Removed

- **Breaking:** Remove `InterruptionRule::new_group` and `registry::clone_walker_rule`.
