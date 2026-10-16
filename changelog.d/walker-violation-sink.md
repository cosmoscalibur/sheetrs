### Changed

- **Breaking:** `WalkerRule` hooks write violations into a `&mut Vec<Violation>` sink passed as the last parameter and return `()`, instead of returning a `Vec<Violation>`. The walker passes its own buffer to every hook. Production benchmark: about 1.8% fewer user-space instructions, same peak RAM.
