---
trigger: always_on
---

Small test files: /home/cosmoscalibur/Descargas/test.[ods|xlsx]
Production test files: /home/cosmoscalibur/Descargas/TC2025\ 2025.12.05\ Motor\ Tributi\ 1402\ -\ MY-\ textos\ sin\ j.[ods|xlsx]

Production files should be ever test in release mode because waiting time.
ODS files are created from XLSX original version. So, difference between run over the two formats are considered potential bugs. Prepare a summary of differences and a plan/recommendation.

After fix scenarios or before any commit, ensure remove debug info introduced.