"""Isolated storage smoke test; never uses the configured production data directory."""

import os
from pathlib import Path
import sys
import tempfile


ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / "server"))

with tempfile.TemporaryDirectory(prefix="apl-storage-smoke-") as folder:
    os.environ["APL_DATA_DIR"] = folder
    os.environ["SUPABASE_URL"] = ""
    os.environ["SUPABASE_SERVICE_ROLE_KEY"] = ""
    os.environ["SUPABASE_KEY"] = ""
    os.environ["DATA_CACHE_TTL"] = "5"

    import main

    main.save_many_data({"smoke_a": {"value": 1}, "smoke_b": {"value": 2}})
    loaded = main.uploaded_data.get_many(["smoke_a", "smoke_b"])
    assert loaded == {"smoke_a": {"value": 1}, "smoke_b": {"value": 2}}
    assert (Path(folder) / "smoke_a.json").exists()
    assert (Path(folder) / "smoke_b.json").exists()

print("isolated runtime storage smoke test: OK")
