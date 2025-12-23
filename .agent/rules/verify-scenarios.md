---
trigger: always_on
---

After any change, verify scenarios. Ever compare ODS vs XLSX scenarios, if
difference detected should be considered as bug. Prepare summary and
plan/recommendation.

Maintain a draft of first run in a ignored directory to compare results after
changes. So, verify accomplishment and avoid regressions.

- Test files: /home/cosmoscalibur/Descargas/minimal_test.[ods|xlsx] is
  safe to test in debug mode.

- Production files: /home/cosmoscalibur/Descargas/TC2025\\ 2025.12.05\\
  Motor\\ Tributi\\ 1402\\ -\\ MY-\\ textos\\ sin\\ j.[ods|xlsx] should be test
  in release mode.

- After fix scenarios or before any commit, ensure remove debug info introduced
  and any temp file with logs.