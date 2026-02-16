# PATCH: Update eu_lobbying_core.py to load gzipped UK index
# ============================================================
#
# You need to make TWO changes in eu_lobbying_core.py:
#
# 1. Add gzip to imports (near the top of the file):
#
#    import gzip
#
# 2. Find where uk_meetings_index.json is loaded (in search_uk_ministerial_meetings).
#    It will look something like:
#
#    BEFORE:
#    -------
#    index_path = Path(__file__).parent / "uk_meetings_index.json"
#    with open(index_path, "r", encoding="utf-8") as f:
#        data = json.load(f)
#
#    AFTER:
#    ------
#    index_path = Path(__file__).parent / "uk_meetings_index.json.gz"
#    with gzip.open(index_path, "rt", encoding="utf-8") as f:
#        data = json.load(f)
#
# That's it. The only differences are:
#   - Filename: .json -> .json.gz
#   - Function: open() -> gzip.open()
#   - Mode: "r" -> "rt" (rt = read text from gzip)
#
# After making these changes:
#   1. Delete the old uk_meetings_index.json (40MB)
#   2. Run: python build_uk_index.py (creates uk_meetings_index.json.gz ~4MB)
#   3. git add . && git commit -m "Compress UK index with gzip" && git push
