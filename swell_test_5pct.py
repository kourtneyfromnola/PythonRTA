"""
Runs a deterministic 5% sample of all valid SWELL cases.

Keep this file in the same folder as swell_bulk_parallel.py.

Run:
    python swell_test_5pct.py

After the 5% results match manual SWELL cases, run:
    python swell_bulk_parallel.py
"""

from swell_bulk_parallel import main

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    main(
        sample_fraction=0.05,
        sample_seed=42,
        output_suffix="SWELL_TEST_5PCT",
        default_workers_override=2,
    )
