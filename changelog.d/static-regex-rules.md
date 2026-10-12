### Changed

- Use static `LazyLock<Regex>` for CALC201, CALC202, REF307, and VUL603 rules to avoid redundant regex compilation on every rule instantiation.
